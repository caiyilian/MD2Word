# MD2Word

Markdown 转 Word (.docx) 工具，支持图片插入、精确页数统计。

## 功能

- Markdown → Word (.docx) 转换
- 支持标题、段落、列表、代码块、引用、表格等常见 Markdown 语法
- 图片插入（支持定位/尺寸控制）
- 精确获取 Word 页数（通过 win32com 调用 MS Word）
- YAML Frontmatter 元数据支持
- 命令行和 Python API 双模式

## 安装

```bash
pip install md2word
```

## 使用

### CLI

```bash
md2word input.md -o output.docx
```

### Python API

```python
from md2word import MD2Word

converter = MD2Word()
result = converter.convert_file("input.md", "output.docx")
print(f"生成文件: {result.path}")
```

## 依赖

- Python >= 3.8
- mistune
- python-docx
- pyyaml

## 协议

MIT
