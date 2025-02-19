import logging
from typing import Dict, List
from .template_parser import TemplateParser
from .content_loader import ContentLoader
from .rule_engine import RuleEngine
from .content_populator import ContentPopulator

class OutputGenerator:
    """输出生成器，协调各个模块完成PPT生成"""
    
    def __init__(self, template_path: str, content_dir: str, rules_config: str = None):
        """
        初始化输出生成器
        
        Args:
            template_path: PPT模板文件路径
            content_dir: 资源目录路径
            rules_config: 规则配置文件路径（可选）
        """
        self.template_parser = TemplateParser(template_path)
        self.content_loader = ContentLoader(content_dir)
        self.rule_engine = RuleEngine(rules_config)
        self.content_populator = ContentPopulator(template_path)
        
    def generate(self, output_path: str):
        """
        生成PPT文件
        
        Args:
            output_path: 输出文件路径
        """
        try:
            # 1. 解析模板
            layouts_info = self.template_parser.parse()
            logging.info("模板解析完成")
            
            # 2. 加载内容
            content_groups = self.content_loader.scan_content()
            logging.info("内容加载完成")
            
            # 3. 处理每个内容组
            for group_name, content in content_groups.items():
                try:
                    # 根据规则匹配布局
                    layout_name = self.rule_engine.match_layout(content)
                    
                    # 创建新幻灯片，使用组名作为标题
                    slide, layout = self.content_populator.add_slide(layout_name, title=group_name)
                    
                    # 填充图片
                    for idx, image_path in enumerate(content['images']):
                        self.content_populator.fill_image(slide, idx, image_path)
                    
                    # 填充文本
                    for idx, text_path in enumerate(content['texts']):
                        text_content = self.content_loader.load_text_content(text_path)
                        print(text_content)
                        self.content_populator.fill_text(slide, len(content['images']) + idx, 
                                                       text_content)
                    
                    # 填充视频
                    for idx, video_path in enumerate(content['videos']):
                        self.content_populator.fill_video(slide, 
                                                        len(content['images']) + 
                                                        len(content['texts']) + idx,
                                                        video_path)
                        
                except Exception as e:
                    logging.error(f"处理内容组 {group_name} 时发生错误: {str(e)}")
                    continue
            
            # 4. 保存文件
            self.content_populator.save(output_path)
            logging.info(f"PPT生成完成，已保存至: {output_path}")
            
        except Exception as e:
            logging.error(f"生成PPT时发生错误: {str(e)}")
            raise
