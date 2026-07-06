---
title: Full Coverage Test Document
author: MD2Word
description: Markdown source for issue 43 full-coverage DOCX and HTML rendering tests.
---

# H1 标题

## H2 标题

### H3 标题

#### H4 标题

##### H5 标题

###### H6 标题

左对齐段落：这是一段包含普通中文、English words、数字 123 和标点符号的段落。

<p style="text-align:center">居中段落：用于验证段落居中对齐。</p>

<p style="text-align:right">右对齐段落：用于验证段落右对齐。</p>

<p style="text-align:justify">两端对齐段落：这一段故意写得较长，用于验证 HTML 输出时可以根据 metadata 还原 justify 对齐、缩进和间距。</p>

文本格式覆盖：**加粗**、*斜体*、<u>单下划线</u>、<span data-underline="double">双下划线</span>、<span data-underline="wave">波浪下划线</span>、~~删除线~~、<span data-strike="double">双删除线</span>、<sup>上标</sup>、<sub>下标</sub>。

字体覆盖：<span style="font-family:Arial;font-size:14pt;color:#C00000">Arial 14pt 红色</span>、<span style="font-family:宋体;font-size:12pt;color:#00B050">宋体 12pt 绿色</span>、<span data-theme-color="accent1">主题色 accent1</span>。

高亮和底纹：<span style="background-color:#FFFF00">黄色高亮</span>、<span data-shading="D9EAF7">浅蓝底纹</span>。

字符高级属性：<span data-character-spacing="40">字符间距加宽</span>、<span data-character-scale="80">字符缩放 80%</span>、<span data-position="raised">字符上移</span>。

1. 有序列表 decimal 第一项
2. 有序列表 decimal 第二项
   1. 嵌套有序列表第一项
   2. 嵌套有序列表第二项

- 无序列表 bullet 第一项
- 无序列表 bullet 第二项
  - 嵌套无序列表第一项
  - 嵌套无序列表第二项

编号格式覆盖：

1. upperRoman: I / II / III
2. lowerRoman: i / ii / iii
3. upperLetter: A / B / C
4. lowerLetter: a / b / c

> 引用块第一层。
>
> > 嵌套引用块第二层。

行内代码：`print("inline code")`

```python
def fenced_code(name: str) -> str:
    return f"hello {name}"
```

---

[OpenAI](https://openai.com) 超链接。

![示例图片](image.png){width=2in align=center}

| 左对齐 | 居中 | 右对齐 |
|:---|:---:|---:|
| A1 | B1 | C1 |
| A2 | B2 | C2 |

复杂表格覆盖项：合并单元格、单元格底纹、单元格宽度、嵌套表格、表格内图片会在标准 DOCX 中以 OpenXML 形式补充。

行内公式：$E=mc^2$。

块级公式：

$$
\int_0^1 x^2 dx = \frac{1}{3}
$$

矩阵、求和、根号：

$$
\begin{bmatrix}1 & 2 \\ 3 & 4\end{bmatrix}
\quad
\sum_{i=1}^{n} i
\quad
\sqrt{x^2+y^2}
$$

脚注引用[^fn1]，尾注、批注、书签、分页符、首页不同页眉页脚、奇偶页不同页眉页脚、多节文档、横向页面、不同页边距和分栏会在标准 DOCX 中补充。

[^fn1]: 这是 Markdown 脚注文本。

- [ ] 未完成任务
- [x] 已完成任务
