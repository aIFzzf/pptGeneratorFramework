import os
from src.output_generator import OutputGenerator
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 设置路径
    template_path = os.path.join(current_dir, "templates", "JK专用PPT-中文版.pptx")
    content_dir = os.path.join(current_dir, "content")
    rules_config = os.path.join(current_dir, "config", "rules.yaml")
    output_dir = os.path.join(current_dir, "output")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    try:
        logging.info("开始创建PPT")
        
        # 创建一个PPT生成器实例
        generator = OutputGenerator(
            template_path=template_path,
            content_dir=content_dir,  # 传入主content目录
            rules_config=rules_config
        )
        
        # 设置输出文件路径
        output_path = os.path.join(output_dir, "combined_presentation.pptx")
        
        # 生成包含所有内容的PPT
        generator.generate(output_path)
        
        logging.info(f"成功生成PPT: {output_path}")
        
    except Exception as e:
        logging.error(f"生成PPT时发生错误: {str(e)}")

if __name__ == "__main__":
    main()
