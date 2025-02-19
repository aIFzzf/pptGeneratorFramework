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
    
    def get_rules(self) -> Dict:
        """
        获取当前规则集
        
        Returns:
            Dict: 当前规则集
        """
        return self.rules
    
    def select_layout(self, content_list: List[Dict]) -> str:
        """
        根据内容列表选择合适的布局
        
        Args:
            content_list: 内容项列表，每项包含 type 和 path
            
        Returns:
            str: 匹配的布局名称
        """
        try:
            # 统计各类型内容的数量
            type_counts = {
                'image': 0,
                'text': 0,
                'video': 0
            }
            
            for item in content_list:
                item_type = item.get('type')
                if item_type in type_counts:
                    type_counts[item_type] += 1
            
            logging.info(f"内容统计: {type_counts}")
            
            # 优先处理视频内容
            if type_counts['video'] > 0:
                return self.rules['video_rules']['any']
            
            # 处理图片内容
            if type_counts['image'] > 0:
                layout = self.rules['image_rules'].get(str(type_counts['image']))
                if layout:
                    return layout
            
            # 处理文本内容
            if type_counts['text'] > 0:
                if type_counts['image'] > 0:
                    return self.rules['text_rules']['with_image']
                return self.rules['text_rules']['single']
            
            # 如果没有找到合适的布局，使用默认布局
            return 'layout_text'  # 默认使用文本布局
            
        except Exception as e:
            logging.error(f"选择布局时发生错误: {str(e)}")
            raise
