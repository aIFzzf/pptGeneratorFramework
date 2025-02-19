from typing import Dict, List
import yaml
import logging

class RuleEngine:
    """规则引擎，用于根据内容匹配最佳布局"""
    
    def __init__(self, rules_config: str = None):
        """
        初始化规则引擎
        
        Args:
            rules_config: 规则配置文件路径（可选）
        """
        self.rules = self._load_default_rules()
        if rules_config:
            self.load_custom_rules(rules_config)
    
    def _load_default_rules(self) -> Dict:
        """
        加载默认规则
        
        Returns:
            Dict: 默认规则集
        """
        return {
            'image_rules': {
                '1': 'layout_single_image',
                '2': 'layout_two_images',
                '3': 'layout_three_images',
            },
            'video_rules': {
                'any': 'layout_video'
            },
            'text_rules': {
                'single': 'layout_text',
                'with_image': 'layout_text_image'
            }
        }
    
    def load_custom_rules(self, config_path: str):
        """
        加载自定义规则配置
        
        Args:
            config_path: YAML配置文件路径
        """
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                custom_rules = yaml.safe_load(f)
                self.rules.update(custom_rules)
        except Exception as e:
            logging.error(f"加载自定义规则时发生错误: {str(e)}")
            raise
    
    def match_layout(self, content: Dict) -> str:
        """
        根据内容匹配最佳布局
        
        Args:
            content: 内容信息字典
            
        Returns:
            str: 匹配的布局名称
        """
        try:
            # 优先处理视频内容
            if content.get('videos'):
                return self.rules['video_rules']['any']
            
            # 处理图片内容
            image_count = len(content.get('images', []))
            if image_count > 0:
                layout = self.rules['image_rules'].get(str(image_count))
                if layout:
                    return layout
            
            # 处理文本内容
            if content.get('texts'):
                if image_count > 0:
                    return self.rules['text_rules']['with_image']
                return self.rules['text_rules']['single']
            
            raise ValueError("无法找到匹配的布局")
            
        except Exception as e:
            logging.error(f"匹配布局时发生错误: {str(e)}")
            raise
