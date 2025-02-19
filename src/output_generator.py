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
        
    def _process_content(self, slide, content_list):
        """处理内容列表"""
        image_idx = 0
        text_idx = 0
        video_idx = 0
        
        for content in content_list:
            content_type = content.get('type')
            content_path = content.get('path')
            
            if content_type == 'image':
                self.content_populator.fill_image(slide, image_idx, content_path)
                image_idx += 1
            elif content_type == 'text':
                # 读取文本文件内容
                try:
                    with open(content_path, 'r', encoding='utf-8') as f:
                        text_content = f.read().strip()
                    self.content_populator.fill_text(slide, text_idx, text_content)
                    text_idx += 1
                except Exception as e:
                    logging.error(f"处理文本文件时发生错误: {str(e)}")
            elif content_type == 'video':
                self.content_populator.fill_video(slide, video_idx, content_path)
                video_idx += 1
            else:
                logging.warning(f"未知的内容类型: {content_type}")
                
    def generate(self, output_path: str):
        """
        生成PPT文件
        
        Args:
            output_path: 输出文件路径
        """
        try:
            # 扫描内容
            content_groups = self.content_loader.scan_content()
            logging.info(f"找到 {len(content_groups)} 个内容组")
            
            # 获取布局规则
            layout_rules = self.rule_engine.get_rules()
            
            # 处理每个内容组
            for group_name, content in content_groups.items():
                try:
                    logging.info(f"\n开始处理内容组: {group_name}")
                    
                    # 根据内容类型和数量选择布局
                    layout_name = self.rule_engine.select_layout(content)
                    logging.info(f"选择布局: {layout_name}")
                    
                    # 创建新幻灯片，使用组名作为标题
                    slide, layout = self.content_populator.add_slide(layout_name, title=group_name)
                    
                    # 处理内容
                    self._process_content(slide, content)
                    
                except Exception as e:
                    logging.error(f"处理内容组 {group_name} 时发生错误: {str(e)}")
                    continue
                
            # 保存文件
            self.content_populator.save(output_path)
            logging.info(f"PPT文件已保存: {output_path}")
            
        except Exception as e:
            logging.error(f"生成PPT时发生错误: {str(e)}")
            raise
