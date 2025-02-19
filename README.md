# PPT Generator

一个基于 Python 的自动化 PPT 生成工具，可以根据模板和内容自动生成 PowerPoint 演示文稿。

## 功能特点

- 支持自定义 PPT 模板
- 自动识别和使用模板中的布局和占位符
- 支持多种内容类型：
  - 图片（支持 PNG、JPG、JPEG、GIF）
  - 文本
  - 视频（支持 MP4、AVI、MOV）
- 智能内容排版：
  - 自动缩放和居中对齐图片
  - 保持图片原始比例
  - 自动适应占位符大小
- 灵活的目录结构支持：
  - 按文件夹组织内容
  - 自动使用文件夹名称作为幻灯片标题

## 安装要求

```bash
pip install -r requirements.txt
```

主要依赖：
- python-pptx >= 0.6.21
- Pillow >= 9.0.0
- pyyaml >= 6.0.0

## 项目结构

```
ppt/
├── config/          # 配置文件目录
├── content/         # 内容资源目录
├── output/          # 生成的PPT输出目录
├── src/            # 源代码目录
│   ├── content_loader.py     # 内容加载模块
│   ├── content_populator.py  # 内容填充模块
│   ├── output_generator.py   # 输出生成模块
│   ├── rule_engine.py       # 规则引擎模块
│   └── template_parser.py   # 模板解析模块
├── templates/       # PPT模板目录
├── main.py         # 主程序入口
└── requirements.txt # 项目依赖
```

## 使用方法

1. 准备 PPT 模板：
   - 将 PowerPoint 模板文件放入 `templates` 目录
   - 确保模板中包含所需的布局和占位符

2. 组织内容：
   - 在 `content` 目录下创建文件夹
   - 文件夹名称将作为幻灯片标题
   - 在文件夹中放入图片、文本等内容

3. 运行程序：
   ```bash
   python main.py
   ```

4. 查看结果：
   - 生成的 PPT 文件将保存在 `output` 目录中

## 配置说明

在 `config` 目录中可以配置：
- 布局匹配规则
- 内容处理规则
- 其他自定义设置

## 注意事项

- 支持的图片格式：PNG、JPG、JPEG、GIF
- 支持的视频格式：MP4、AVI、MOV
- 文本文件应使用 UTF-8 编码
- 建议使用 16:9 比例的 PPT 模板

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
