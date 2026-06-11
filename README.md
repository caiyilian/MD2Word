# MD2Word

Markdown 与 Word (.docx) 双向转换工具，支持图片插入、精确页数统计、LaTeX 公式渲染。

## 功能

- **Markdown → Word** (.docx) 转换
  - 标题、段落、列表、代码块、引用、表格等常见 Markdown 语法
  - 图片插入（支持定位/尺寸控制）
  - LaTeX 数学公式渲染（`$...$`、`$$...$$`、`\begin{align}...\end{align}`）
  - 文本增强：删除线、下划线、上标/下标、自定义字体字号
  - 精确获取 Word 页数（通过 win32com 调用 MS Word）
  - YAML Frontmatter 元数据支持
  - 样式自定义（YAML 配置文件）
- **Word → Markdown** (.docx → .md)
  - 标题、段落、列表（有序/无序）、水平线
  - 加粗、斜体、下划线、删除线、行内代码
  - 上标/下标、自定义字体字号还原
- 命令行和 Python API 双模式

## 安装

```bash
pip install md2word
```

## 使用

### CLI

```bash
# Markdown 转 Word
md2word input.md -o output.docx

# 获取页数
md2word input.md -o output.docx --count-pages

# 指定样式配置
md2word input.md -o output.docx --style style.yaml

# Word 转 Markdown
md2word input.docx -r -o output.md
```

### Python API

```python
from md2word import MD2Word

# Markdown 转 Word
converter = MD2Word()
result = converter.convert_file("input.md", "output.docx")
print(f"生成文件: {result.path}")

# Word 转 Markdown
from md2word.extractor.docx_extractor import DocxExtractor
from md2word.writer.md_writer import MdWriter

extractor = DocxExtractor()
doc = extractor.extract("input.docx")
writer = MdWriter()
md_text = writer.write(doc)
```

## 依赖

- Python >= 3.8
- mistune
- python-docx
- pyyaml
- latex2mathml
- lxml
- pywin32（可选，用于页数统计）

## 协议

MIT
