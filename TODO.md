# 项目待办事项

## 核心目标重定义

**docx → 元数据 → 完美还原 docx**

不是简单的 docx ↔ md 转换，而是：

1. **docx → meta**: 把 docx 里所有信息提取为人类可读的元数据（md/txt/yaml/json 格式），图片等二进制资源保存到单独文件夹
2. **meta → docx**: 从元数据完全一模一样的还原回 docx，包括：
   - 字体、字号、颜色
   - 段落样式、对齐、缩进、间距
   - 所有文本装饰（加粗、斜体、下划线、删除线、上标/下标）
   - 图片（尺寸、位置、环绕方式）
   - 表格（合并单元格、边框、底纹）
   - 公式
   - 页眉/页脚
   - 脚注/尾注/批注
   - 超链接
   - 列表（多级编号）
   - 页面设置（纸张、边距、分栏）

**不能偷懒：** 不允许直接保存 docx 的二进制数据/XML blob，meta 信息必须是结构化的、人能看懂的内容。

## 已知问题

### 字体字号还原不准确
- 当前 `DocxExtractor`/`MdWriter` 对字体字号的提取和还原不够精准
- 颜色信息完全没有提取
- 某些特殊字体名可能解析失败

### 图片信息不完整
- 图片的对齐方式（左/中/右）未正确提取
- 图片的环绕方式（嵌入型/上下型/四周型）未提取
- 单元格内图片处理不完善

## 未完成功能（Phase 8+）

### 1. 完整元数据提取层
- 设计通用的元数据格式（YAML/JSON/Toml 或其他结构化格式）
- 覆盖所有 docx 元素的完整信息提取
- 需要一个 `MetaExtractor`（取代当前的 `DocxExtractor` + `MdWriter`）

### 2. 完美还原层
- 从元数据重建 docx
- 需要一个 `MetaRenderer`（增强/取代当前的 `docx_renderer.py`）
- 闭环验证：docx → meta → docx，两个 docx 必须二进制一致（字节级对比）

### 3. 公式还原（OMML → LaTeX）
- 与 Phase 5 的双向对称
- 将 docx 中的 OMML 数学公式转换回 LaTeX 语法
- 需要实现 OMML XML → LaTeX 字符串的反向转换器

### 4. 脚注/尾注提取
- 从 `word/footnotes.xml` 读取脚注定义
- 在段落中检测 `w:footnoteReference` 引用
- 输出为 Markdown 脚注语法 `[^id]: text`

### 5. 批注提取
- 从 `word/comments.xml` 读取批注定义（作者、日期、内容）
- 在段落中检测 `w:commentReference` 引用
- 输出格式：`<!-- author: text -->`

### 6. 超链接还原
- 当前已能提取超链接文本和 URL，但还原到 docx 时未正确重建
- 需要在前向渲染器里支持 `Hyperlink` 模型 → docx 超链接

### 7. 代码块闭环
- 当前 `CodeBlock` 在 md→docx→md 转换后丢失格式

### 8. 文档加密功能
- 目前项目无任何文档加密/密码保护功能
- 需要支持：
  - 读取加密 docx 文档（需要密码才能打开）
  - 写入加密 docx 文档（设置打开密码）
  - 可能涉及 `python-docx` 不支持的功能，需要 `pywin32` COM 或底层 XML 操作

## 架构设计思路

```
docx  →  DocxExtractor  →  元数据（YAML/JSON/md）  →  MetaRenderer  →  docx
                                                           ↑
                                                  Meta中引用的图片等资源
```

元数据格式示例（YAML）：
```yaml
- type: paragraph
  alignment: center
  runs:
    - text: "Hello World"
      bold: true
      font:
        name: "Arial"
        size: 12pt
        color: "FF0000"
      spacing:
        before: 0
        after: 6

- type: image
  src: images/image_001.png
  width: 5.77in
  height: 1.60in
  align: center
  wrap: inline

- type: table
  rows:
    - cells: ["A1", "B1", "C1"]
      shading: "D9E2F3"
    - cells: ["A2", "B2", "C2"]
  column_widths: [2in, 3in, 2in]
```