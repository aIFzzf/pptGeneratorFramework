import os
from typing import Dict, List
import logging
from pathlib import Path

class ContentLoader:
    """内容加载器，用于扫描和加载用户提供的资源"""
    
    SUPPORTED_IMAGE_FORMATS = ['.png', '.jpg', '.jpeg', '.gif']
    SUPPORTED_VIDEO_FORMATS = ['.mp4', '.avi', '.mov']
    
    def __init__(self, content_dir: str):
        """
        初始化内容加载器
        
        Args:
            content_dir: 资源目录路径
        """
        self.content_dir = Path(content_dir)
        self.content_groups = {}
        
    def scan_content(self) -> Dict:
        """
        扫描资源目录，按规则分类内容
        
        Returns:
            Dict: 分类后的内容组
        """
        try:
            for item in self.content_dir.iterdir():
                if item.is_dir():
                    group_content = {
                        'images': [],
                        'texts': [],
                        'videos': [],
                    }
                    
                    # 扫描图片
                    for img_file in item.glob('**/*'):
                        if img_file.suffix.lower() in self.SUPPORTED_IMAGE_FORMATS:
                            group_content['images'].append(str(img_file))
                            
                    # 扫描文本
                    for text_file in item.glob('**/*.txt'):
                        group_content['texts'].append(str(text_file))
                        
                    # 扫描视频
                    for video_file in item.glob('**/*'):
                        if video_file.suffix.lower() in self.SUPPORTED_VIDEO_FORMATS:
                            group_content['videos'].append(str(video_file))
                    
                    self.content_groups[item.name] = group_content
            
            return self.content_groups
            
        except Exception as e:
            logging.error(f"扫描内容时发生错误: {str(e)}")
            raise
    
    def load_text_content(self, text_file: str) -> str:
        """
        加载文本文件内容
        
        Args:
            text_file: 文本文件路径
            
        Returns:
            str: 文本内容
        """
        try:
            with open(text_file, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            logging.error(f"读取文本文件时发生错误: {str(e)}")
            raise
