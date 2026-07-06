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
- **Word → 结构化元数据 → Word** (.docx → JSON + 资源 → .docx)
  - 将 docx ZIP 包中的 XML 部件提取为可读 JSON 树
  - 生成语义化 document 索引，覆盖正文、段落/Run 格式、表格、图片、节属性、脚注/尾注/批注等信息
  - 提取 DrawingML 图片裁剪、旋转/翻转、边框线条、效果和图片超链接
  - 提取 VML 形状、文本框、填充/线条和形状内文本
  - 提取样式、编号、设置、字体表、主题和 docProps 文档属性的语义摘要
  - 提取域代码、内容控件和基础修订追踪的语义摘要
  - 识别图表、SmartArt/diagrams、OLE 嵌入、VBA、ActiveX、glossary、people、customXml 等高级包部件
  - 提取图表类型、系列名称、分类/数值引用、缓存点、坐标轴和图例摘要
  - 提取 SmartArt 数据点/连接、布局节点、样式标签和颜色定义摘要
  - 将图片、嵌入对象等二进制资源保存到独立资源目录
  - 保存 ZIP 容器级 exact payload cache；结构化 XML/资源未改动时可复现原始压缩流，实现真实 docx 的字节级一致还原
  - 可从元数据和资源目录重建可打开的 docx 包
  - 详细阶段计划见 `docs/issue41_roundtrip_plan.md`
- **Word → 结构化元数据 → HTML** (.docx → JSON + 资源 → .html)
  - `--to-html` 先提取 metadata，再从 `document.json` 语义索引渲染 HTML
  - 输出独立 HTML，图片资源以内嵌 data URI 呈现，公式加载 MathJax
  - 覆盖标题、段落/run 样式、图片、表格合并、列表层级、页眉页脚、脚注/尾注/批注
  - 标准覆盖样例：`examples/full_coverage.md` 和 `examples/full_coverage.docx`
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

# Word 转结构化元数据
md2word input.docx --extract-meta -o output/input_meta

# 元数据还原为 Word
md2word output/input_meta --restore-meta -o restored.docx

# 提取、还原并验证字节级一致性
md2word input.docx --roundtrip-meta -o output/input_roundtrip

# 从 docx metadata 渲染 HTML
md2word input.docx --to-html -o output/input.html

# 从已提取的 metadata 目录渲染 HTML
md2word output/input_meta --to-html -o output/input.html
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

# Word 结构化元数据闭环
from md2word.meta import extract_docx_metadata, restore_docx_from_metadata

extract_docx_metadata("input.docx", "output/input_meta")
restore_docx_from_metadata("output/input_meta", "restored.docx")

from md2word.meta import verify_docx_metadata_roundtrip

result = verify_docx_metadata_roundtrip("input.docx", "output/input_roundtrip")
print(result["byte_identical"])

from md2word.html import render_docx_to_html, render_metadata_to_html

render_docx_to_html("input.docx", "output/input.html")
render_metadata_to_html("output/input_meta", "output/input-from-meta.html")
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
