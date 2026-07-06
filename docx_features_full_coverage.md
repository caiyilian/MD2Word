# DOCX 功能全覆盖清单 — 提取与还原

> 本文档列出 docx 格式中所有可能被提取并还原的功能/信息，按 OpenXML 的 ZIP 包内文件结构组织。
>
> 标记说明：
> - ✅ 已实现（提取 + 还原均完成）
> - ⚠️ 部分实现（提取/还原一侧缺失，或信息不完整）
> - ❌ 未实现
> - 🔲 不适用/无需提取

---

## 目录

1. [word/document.xml — 正文内容](#1-worddocumentxml--正文内容)
2. [word/styles.xml — 样式系统](#2-wordstylesxml--样式系统)
3. [word/numbering.xml — 编号定义](#3-wordnumberingxml--编号定义)
4. [word/header*.xml + word/footer*.xml — 页眉页脚](#4-wordheaderxml--wordfooterxml--页眉页脚)
5. [word/footnotes.xml + word/endnotes.xml — 脚注尾注](#5-wordfootnotesxml--wordendnotesxml--脚注尾注)
6. [word/comments.xml — 批注](#6-wordcommentsxml--批注)
7. [word/media/ — 媒体资源](#7-wordmedia--媒体资源)
8. [word/embeddings/ — 嵌入对象](#8-wordembeddings--嵌入对象)
9. [word/theme/ — 主题](#9-wordtheme--主题)
10. [word/settings.xml — 文档设置](#10-wordsettingsxml--文档设置)
11. [word/fontTable.xml — 字体表](#11-wordfonttablexml--字体表)
12. [word/glossary/ — 构建基块](#12-wordglossary--构建基块)
13. [word/charts/ — 图表](#13-wordcharts--图表)
14. [word/diagrams/ — 图表数据](#14-worddiagrams--图表数据)
15. [word/activeX*.xml — ActiveX 控件](#15-wordactivexxml--activex-控件)
16. [docProps/ — 文档属性](#16-docprops--文档属性)
17. [word/vbaProject.bin — 宏](#17-wordvbaprojectbin--宏)
18. [word/people.xml — 人员](#18-wordpeoplexml--人员)
19. [word/revisions/ — 修订](#19-wordrevisions--修订)
20. [word/mailMerge/ — 邮件合并](#20-wordmailmerge--邮件合并)
21. [跨文档/多节属性](#21-跨文档多节属性)
22. [元数据格式设计建议](#22-元数据格式设计建议)
23. [闭环验证策略](#23-闭环验证策略)

---

## 1. word/document.xml — 正文内容

### 1.1 Run 级别属性 (w:rPr) — 当前仅提取 bold/italic/underline/strike/vertAlign/rFonts/sz

| 属性 | XML 路径 | 状态 | 说明 |
|------|----------|------|------|
| **加粗** | w:rPr/w:b | ✅ | 已实现 |
| **斜体** | w:rPr/w:i | ✅ | 已实现 |
| **下划线** | w:rPr/w:u (val 类型: single, double, words, dotted, dash, dotDash, wave, etc.) | ⚠️ | 只提取了是否存在，未提取下划线类型、颜色、粗细 |
| **删除线** | w:rPr/w:strike | ✅ | 已实现 |
| **双删除线** | w:rPr/w:dstrike | ❌ | 未实现 |
| **上标/下标** | w:rPr/w:vertAlign | ✅ | 已实现 |
| **字体名** | w:rPr/w:rFonts (ascii, hAnsi, eastAsia, cs) | ⚠️ | 只提取了 ascii/hAnsi，未提取 eastAsia 字体和 cs 字体 |
| **字号** | w:rPr/w:sz (半号单位) | ✅ | 已实现 |
| **字号(东亚)** | w:rPr/w:szCs | ❌ | 东亚文字字号可能独立设置 |
| **字体颜色** | w:rPr/w:color (val, themeColor, themeShade, themeTint) | ❌ | **严重缺失** — 颜色信息完全未提取 |
| **高亮** | w:rPr/w:highlight (val: yellow, green, cyan, magenta, red, blue, etc.) | ❌ | 未实现 |
| **底纹/背景** | w:rPr/w:shd (fill, val, color) | ❌ | 文字级别底纹（区别于段落底纹） |
| **字符间距** | w:rPr/w:spacing (val: 微调宽度, twips 单位) | ❌ | 字符之间的间距调整 |
| **字符缩放** | w:rPr/w:w (val: 百分比) | ❌ | 字符宽度缩放 |
| **字符位置** | w:rPr/w:position (val: 提升/降低, twips) | ❌ | 区别于 vertAlign 的精细位置调整 |
| **文字效果** | w:rPr/w:effect (val: antialias, shadow, outline, emboss, engrave, etc.) | ❌ | 完全未实现 |
| **文字边框** | w:rPr/w:bdr (val, sz, color, space) | ❌ | 字符级别的边框 |
| **轮廓** | w:rPr/w:outline | ❌ | 空心文字 |
| **阴影** | w:rPr/w:shadow | ❌ | 文字阴影 |
| **阳文/阴文** | w:rPr/w:imprint / w:rPr/w:emboss | ❌ | 浮雕效果 |
| **全部大写** | w:rPr/w:caps / w:rPr/w:smallCaps | ❌ | 字母大小写转换 |
| **隐藏文字** | w:rPr/w:vanish | ❌ | 隐藏文本（不显示不打印） |
| **语言** | w:rPr/w:lang (val, eastAsia, bidi) | ❌ | 校对语言 |
| **下划线颜色** | w:rPr/w:u 的 uColor 属性 | ❌ | 下划线可独立设置颜色 |
| **RTF 字符样式引用** | w:rPr/w:rStyle (val: 引用样式ID) | ❌ | 字符级样式引用 |
| **高亮(theme-aware)** | w:rPr/w:highlight 关联 themeColor | ❌ | 主题色高亮 |

### 1.2 段落级别属性 (w:pPr) — 当前仅提取对齐

| 属性 | XML 路径 | 状态 | 说明 |
|------|----------|------|------|
| **对齐方式** | w:pPr/w:jc (val: left, center, right, both) | ✅ | 已实现 |
| **左缩进** | w:pPr/w:ind (left, leftChars) | ❌ | 段落左缩进 |
| **右缩进** | w:pPr/w:ind (right, rightChars) | ❌ | 段落右缩进 |
| **首行缩进** | w:pPr/w:ind (firstLine, firstLineChars) | ❌ | 首行缩进 |
| **悬挂缩进** | w:pPr/w:ind (hanging, hangingChars) | ❌ | 悬挂缩进 |
| **段前间距** | w:pPr/w:spacing (before, beforeLines, beforeAutoSpacing) | ❌ | 段前间距 |
| **段后间距** | w:pPr/w:spacing (after, afterLines, afterAutoSpacing) | ❌ | 段后间距 |
| **行间距** | w:pPr/w:spacing (line, lineRule: auto, atLeast, exactly, multiple) | ❌ | 行距类型和值 |
| **段落底纹** | w:pPr/w:shd (fill, val, color) | ❌ | 段落背景色 |
| **段落边框** | w:pPr/w:pBdr (top, bottom, left, right, between, bar) | ❌ | 存在水平线检测，但其他边框完全未提取 |
| **与下段同页** | w:pPr/w:keepNext | ❌ | 保持与下段不分页 |
| **段内不分页** | w:pPr/w:keepLines | ❌ | 段内行不分页 |
| **段前分页** | w:pPr/w:pageBreakBefore | ❌ | 段前强制分页 |
| **孤行控制** | w:pPr/w:widowControl | ❌ | 孤行/寡行控制 |
| **大纲级别** | w:pPr/w:outlineLvl (val: 0-9) | ❌ | 非标题样式的大纲级别 |
| **段落样式引用** | w:pPr/w:pStyle (val: 样式ID) | ✅ | 已通过 style.name 实现 |
| **编号引用** | w:pPr/w:numPr (w:ilvl, w:numId) | ⚠️ | 部分实现，但缺少详细的抽象编号属性 |
| **制表位** | w:pPr/w:tabs (w:tab: val, pos, leader) | ❌ | 完全未实现 |
| **段落方向** | w:pPr/w:bidi | ❌ | 从右到左段落 |
| **自动调整间距** | w:pPr/w:autoSpaceDE / w:autoSpaceDN | ❌ | 东亚文字间距调整 |
| **文字方向** | w:pPr/w:textDirection (val: btLr, tbRl, lrTb, etc.) | ❌ | 段落文字方向 |
| **对齐网格** | w:pPr/w:snapToGrid | ❌ | 对齐页面网格 |
| **上下文间距** | w:pPr/w:contextualSpacing | ❌ | 相同样式段落间自动删除间距 |
| **镜像缩进** | w:pPr/w:mirrorIndents | ❌ | 奇偶页镜像缩进 |
| **禁止自动连字符** | w:pPr/w:suppressAutoHyphens | ❌ | 禁止自动断字 |
| **自动换行** | w:pPr/w:wordWrap | ❌ | 允许断字换行 |
| **溢出标点** | w:pPr/w:overflowPunct | ❌ | 允许标点溢出边界 |
| **行首标点** | w:pPr/w:topLinePunct | ❌ | 行首标点调整 |
| **禁则处理** | w:pPr/w:kinsoku | ❌ | 东亚文字禁则处理 |
| **段落编号(行号)** | w:pPr/w:lnNumType (val: restartPage, restartSect, suppress) | ❌ | 行号控制 |
| **div 引用** | w:pPr/w:divId (val: div ID) | ❌ | HTML 风格的 div 布局引用 |
| **段落级样式变体** | w:pPr/w:pPrChange (rPr 修订) | ❌ | 段落样式修订 |

### 1.3 列表与编号 — 当前仅提取 ordered/无序/bullet

| 属性 | XML 路径 | 状态 | 说明 |
|------|----------|------|------|
| **有序/无序检测** | w:pPr/w:numPr/w:numId | ✅ | 已实现 |
| **列表层级** | w:pPr/w:numPr/w:ilvl | ✅ | 已实现 |
| **编号格式** | w:numFmt (val: decimal, upperRoman, lowerRoman, upperLetter, lowerLetter, ordinal, bullet, etc.) | ❌ | 未提取编号格式 |
| **编号起始值** | w:start (val: 起始数值) | ❌ | 列表可从非 1 开始 |
| **编号是否继续** | 上一个 numId 匹配表示继续 | ❌ | 需要判断编号连续性 |
| **项目符号字符** | w:lvl/w:lvlText (val: 符号字符) | ❌ | 自定义项目符号 |
| **项目符号字体** | w:lvl/w:rPr (字体信息) | ❌ | 符号的字体 |
| **列表缩进/对齐** | w:lvl/w:pPr/w:ind / w:jc | ❌ | 每级列表的缩进和对齐 |
| **图片项目符号** | w:lvl/w:lvlPicBulletId | ❌ | 图片作为项目符号 |
| **编号重启** | w:numPr/w:numId 不同实例 | ❌ | 列表重新开始编号 |
| **列表样式引用** | w:pPr/w:numPr/w:numId 关联 w:num 定义 | ❌ | 完整样式链 |

### 1.4 表格 — 当前仅提取单元格文本 + 列对齐

| 属性 | XML 路径 | 状态 | 说明 |
|------|----------|------|------|
| **单元格文本** | w:tc/w:p/w:r/w:t | ✅ | 已实现 |
| **列对齐** | w:tc/w:p/w:pPr/w:jc | ⚠️ | 首行对齐，但未提取到每行 |
| **单元格合并(水平)** | w:tc/w:tcPr/w:gridSpan (val: 跨列数) | ❌ | 完全未实现 |
| **单元格合并(垂直)** | w:tc/w:tcPr/w:vMerge (val: continue, restart) | ❌ | 完全未实现 |
| **单元格底纹** | w:tc/w:tcPr/w:shd (fill, val, color) | ❌ | 单元格背景色 |
| **单元格边框** | w:tc/w:tcPr/w:tcBorders (top, bottom, left, right, insideH, insideV) | ❌ | 单个单元格边框可独立设置 |
| **单元格边距** | w:tc/w:tcPr/w:tcMar (top, bottom, left, right) | ❌ | 单元格内边距 |
| **单元格宽度** | w:tc/w:tcPr/w:tcW (w, type: dxa, auto, nil, pct) | ❌ | 单元格宽度 |
| **单元格垂直对齐** | w:tc/w:tcPr/w:vAlign (val: top, center, bottom) | ❌ | 垂直对齐方式 |
| **单元格文字方向** | w:tc/w:tcPr/w:textDirection | ❌ | 单元格内文字方向 |
| **单元格内图片** | w:tc/w:p/w:r/w:drawing 中的图片 | ❌ | 表格内图片完全未处理 |
| **单元格内嵌套表格** | w:tc/w:tbl | ❌ | 嵌套表格（表格中再套表格） |
| **行高** | w:tr/w:trPr/w:trHeight (val, hRule: atLeast, exactly, auto) | ❌ | 行高设置 |
| **表头行** | w:tr/w:trPr/w:tblHeader | ❌ | 跨页重复表头 |
| **行分页控制** | w:tr/w:trPr/w:cantSplit / w:trPr/w:tblHeader | ❌ | 禁止行跨页打断 |
| **表格宽度** | w:tbl/w:tblPr/w:tblW (w, type) | ❌ | 表格整体宽度 |
| **表格对齐** | w:tbl/w:tblPr/w:jc (val: left, center, right) | ❌ | 表格在页面上的对齐 |
| **表格缩进** | w:tbl/w:tblPr/w:tblInd (w, type) | ❌ | 表格整体左缩进 |
| **表格边框** | w:tbl/w:tblPr/w:tblBorders (top, bottom, left, right, insideH, insideV) | ❌ | 未提取（渲染时用了 Table Grid 但未从源提取） |
| **表格布局** | w:tbl/w:tblPr/w:tblLayout (type: fixed, autofit) | ❌ | 固定宽度 vs 自动适应 |
| **表格样式引用** | w:tbl/w:tblPr/w:tblStyle (val: 样式ID) | ❌ | 表格样式 |
| **表格样式覆盖** | w:tbl/w:tblPr/w:tblLook (val, firstRow, lastRow, firstColumn, lastColumn, noHBand, noVBand) | ❌ | 样式覆盖位掩码 |
| **行/列带区** | w:tblPr/w:tblLook 的 banding 属性 | ❌ | 斑马纹控制 |
| **表格描述** | w:tbl/w:tblPr/w:tblPrEx | ❌ | 例外表格属性 |
| **单元格内公式** | 单元格内包含 w:oMath | ❌ | 表格内公式 |
| **单元格内列表** | 单元格内包含 w:numPr 的段落 | ❌ | 表格单元格内的列表 |

### 1.5 图片与绘图对象 (w:drawing / wp:inline / wp:anchor)

| 属性 | XML 路径 | 状态 | 说明 |
|------|----------|------|------|
| **图片尺寸** | wp:extent (cx, cy) | ✅ | 已提取 |
| **替代文本** | wp:docPr (descr, name, title) | ⚠️ | 只提取了 descr/name，未提取 title |
| **图片嵌入方式** | wp:inline / wp:anchor | ❌ | 嵌入型 vs 非嵌入型未区分 |
| **环绕方式** | wp:anchor/wrapSquare, wrapNone, wrapThrough, wrapTight, wrapTopAndBottom | ❌ | 完全未实现 |
| **环绕文字边** | wp:anchor/wrapText (val: both, left, right, largest) | ❌ | 文字环绕在哪一侧 |
| **水平位置** | wp:anchor/positionH (posOffset, relativeFrom: page, column, margin, character, leftMargin, rightMargin, insideMargin, outsideMargin) | ❌ | 完全未实现 |
| **垂直位置** | wp:anchor/positionV (posOffset, relativeFrom: page, paragraph, line, margin, topMargin, bottomMargin, insideMargin, outsideMargin) | ❌ | 完全未实现 |
| **图片旋转** | wp:anchor/wp:extent 的 rot 属性（0.1毫弧度） | ❌ | 未实现 |
| **图片翻转** | a:xfrm/hflip / vflip | ❌ | 水平/垂直翻转 |
| **图片裁剪** | a:srcRect (l, t, r, b: 百分比) | ❌ | 裁剪区域 |
| **图片边框** | a:ln/a:prstDash/a:solidFill (线型、颜色、宽度) | ❌ | 图片边框样式 |
| **图片阴影/发光/柔化/3D** | a:effectLst (a:shadow, a:glow, a:softEdge, a:presetShadow, a:reflection) | ❌ | 图片效果 |
| **图片超链接** | wp:docPr/a:hlinkClick (r:id, tooltip) | ❌ | 可点击的图片链接 |
| **图片 ID** | wp:docPr (id, name) | ❌ | 未提取图片唯一标识 |
| **图片在表格内** | 图片位于 w:tc/w:p/w:r/w:drawing | ❌ | 未处理 |
| **图片在页眉页脚** | 页眉页脚中的图片 | ❌ | 未处理 |
| **SVG 图片** | 通过 a:blip 的 r:embed 链接到 SVG 部分 | ❌ | SVG 的支持 |
| **多图片选择** | 图片可能有多个嵌入链接（不同分辨率） | ❌ | 未处理 |

### 1.6 绘图形状 (VML / DrawingML)

| 形状类型 | XML 路径 | 说明 |
|----------|----------|------|
| **矩形/圆角矩形** | w:object/v:rect / v:roundrect | ❌ |
| **椭圆/圆形** | w:object/v:oval | ❌ |
| **线条/箭头** | w:object/v:line | ❌ |
| **任意多边形** | w:object/v:polyline / v:curve | ❌ |
| **自选图形** | w:object/v:shape (各种预设形状) | ❌ |
| **文本框** | w:object/v:textbox / w:p (在形状内) | ❌ |
| **链接文本框** | 多个文本框内容串联 | ❌ |
| **艺术字** | w:object/v:shape 的 style 包含旋转/渐变/3D 等 | ❌ |
| **组合图形** | v:group (多个形状组合) | ❌ |
| **画布** | w:object/v:canvas | ❌ |
| **墨迹注释** | w:object/v:ink | ❌ |
| **形状填充** | 渐变、图片、图案、纯色填充 | ❌ |
| **形状线条** | 实线、虚线、渐变线、多色线 | ❌ |
| **形状效果** | 阴影、发光、柔化边缘、3D 格式、3D 旋转 | ❌ |
| **形状中的文本** | 形状内部的文字 | ❌ |

### 1.7 公式 (OMML) — 当前仅正向渲染

| 属性 | XML 路径 | 状态 | 说明 |
|------|----------|------|------|
| **LaTeX → OMML** | 正向渲染 | ✅ | Phase 5 完成 |
| **OMML → LaTeX** | 反向提取 | ❌ | **严重缺失** — 完全未实现从 docx 提取公式 |
| **行内公式** | w:oMathPara / w:oMath | ❌ | 反向提取未实现 |
| **块级公式** | w:oMathPara (display="true") | ❌ | 反向提取未实现 |
| **公式编号** | 公式后的 SEQ 域或制表位编号 | ❌ | 未提取公式编号 |
| **公式中的文本** | w:oMath 中的 w:t 元素 | ❌ | 公式中的普通文本 |
| **公式中的矩阵** | m:matrix | ❌ | 矩阵结构 |
| **公式专业格式** | 分数、根号、积分、求和、极限、大括号、函数等 | ❌ | 完整的 OMML AST 解析 |
| **公式字符格式** | 公式内字符的字体、大小、颜色 | ❌ | 公式中的字体信息 |

### 1.8 域代码 (Fields)

| 域类型 | XML 路径 | 说明 |
|--------|----------|------|
| **TOC 域** | w:fldChar (begin, separate, end) + w:instrText | ❌ |
| **PAGE 域** | 页码 | ❌ |
| **NUMPAGES 域** | 总页数 | ❌ |
| **DATE / TIME 域** | 日期时间 | ❌ |
| **AUTHOR 域** | 作者 | ❌ |
| **TITLE 域** | 标题 | ❌ |
| **FILENAME 域** | 文件名 | ❌ |
| **REF 域** | 交叉引用 | ❌ |
| **SEQ 域** | 自动编号序列 | ❌ |
| **HYPERLINK 域** | 超链接字段 | ❌ |
| **INCLUDEPICTURE 域** | 外部图片引用 | ❌ |
| **INCLUDETEXT 域** | 外部文本文件 | ❌ |
| **MACROBUTTON 域** | 宏按钮 | ❌ |
| **MAILMERGE 域** | 邮件合并字段 | ❌ |
| **FORMULA 域** | 表格公式 (=SUM(LEFT)) | ❌ |
| **PAGEREF 域** | 引用页码 | ❌ |
| **NOTEREF 域** | 引用脚注 | ❌ |
| **域代码字符串** | w:instrText 中的指令 | ❌ |
| **域结果** | w:fldChar 后的 w:r/w:t 中的计算结果 | ❌ |
| **域锁** | w:fldLock | ❌ |
| **域脏标** | w:dirty | ❌ |

### 1.9 内容控件 (Structured Document Tags / SDT)

| 控件类型 | XML 路径 | 说明 |
|----------|----------|------|
| **纯文本控件** | w:sdt/w:sdtPr/w:text | ❌ |
| **富文本控件** | w:sdt/w:sdtPr/w:richText | ❌ |
| **图片控件** | w:sdt/w:sdtPr/w:date | ❌ |
| **下拉列表控件** | w:sdt/w:sdtPr/w:dropDownList | ❌ |
| **组合框控件** | w:sdt/w:sdtPr/w:comboBox | ❌ |
| **日期选择器** | w:sdt/w:sdtPr/w:date | ❌ |
| **复选框控件** | w:sdt/w:sdtPr/w:checkBox | ❌ |
| **重复节控件** | w:sdt/w:sdtPr/w:repeatingSection | ❌ |
| **分组控件** | w:sdt/w:sdtPr/w:group | ❌ |
| **控件标签/别名** | w:sdt/w:sdtPr/w:alias | ❌ |
| **控件锁定** | w:sdt/w:sdtPr/w:lock (canDelete, canEdit, noRemove, noSelect) | ❌ |
| **控件占位符** | w:sdt/w:sdtPr/w:placeholder | ❌ |
| **控件数据绑定** | w:sdt/w:sdtPr/w:dataBinding (xpath, storeItemID) | ❌ |
| **控件标记** | w:sdt/w:sdtPr/w:tag | ❌ |

### 1.10 交叉引用

| 类型 | 说明 |
|------|------|
| **标题交叉引用** | 引用文档中的标题文字 | ❌ |
| **书签交叉引用** | 引用书签位置 | ❌ |
| **图表交叉引用** | 引用图/表编号 | ❌ |
| **公式交叉引用** | 引用公式编号 | ❌ |
| **脚注/尾注交叉引用** | 引用脚注 | ❌ |
| **页码交叉引用** | 引用页面位置 | ❌ |

### 1.11 题注 (Caption)

| 属性 | 说明 |
|------|------|
| **图题注** | 自动编号的图题注 | ❌ |
| **表题注** | 自动编号的表题注 | ❌ |
| **公式题注** | 自动编号的公式题注 | ❌ |
| **自定义题注标签** | 用户自定义题注标签 | ❌ |
| **题注编号格式** | 1, 2, 3 / 一、二、三 / 1.1, 1.2 | ❌ |

### 1.12 索引与引文表

| 类型 | 说明 |
|------|------|
| **索引条目** | XE 域标记的索引条目 | ❌ |
| **索引表** | INDEX 域生成的索引 | ❌ |
| **引文目录** | TOA 域 | ❌ |
| **图表目录** | TOC 域 + SEQ 域 | ❌ |
| **引文/书目** | word/citations.xml, word/bibliography.xml | ❌ |

### 1.13 书签

| 属性 | XML 路径 | 说明 |
|------|----------|------|
| **书签名称** | w:bookmarkStart (w:name) | ⚠️ | Phase 13 标记了，但实际提取代码未确认 |
| **书签范围** | w:bookmarkStart + w:bookmarkEnd 配对 | ⚠️ | 需确认实现程度 |
| **隐藏书签** | w:bookmarkStart 的 _name 以下划线开头 | ❌ | Word 内部隐藏书签 |
| **书签列** | w:bookmarkStart 的 column 属性 | ❌ | 表格列书签 |

### 1.14 特殊内容

| 类型 | 说明 |
|------|------|
| **首字下沉** | w:pPr/w:framePr 或 w:dropCap | ❌ |
| **批注框** | 批注可能以浮动框形式显示在文档中 | ❌ |
| **Ruby 注音** | w:ruby (东亚文字拼音/注音) | ❌ |
| **双行合一** | w:twoLineOne | ❌ |
| **竖排中文** | w:textDirection 垂直书写 | ❌ |
| **制表符** | w:tab / w:tab 作为特殊字符 | ❌ |
| **不间断空格** | w:noBreakHyphen / w:softHyphen / w:br (clear, type) | ❌ |
| **分节符** | w:pPr/w:sectPr (不仅是分页符) | ❌ |
| **分栏符** | w:br (type="column") | ❌ |
| **换行符** | w:br (type="textWrapping") | ⚠️ | 可能被 mistune 处理但提取代码未确认 |
| **特殊字符** | 特殊 Unicode 字符、符号、表情 | ❌ |

---

## 2. word/styles.xml — 样式系统

| 样式类型 | 属性 | 说明 |
|----------|------|------|
| **段落样式** | 完整样式定义：名称、别名、基于、后续、快捷键、优先级 | ❌ |
| **字符样式** | 字符级别的样式定义 | ❌ |
| **表格样式** | 表格样式定义（含条件格式：整行、整列、首行、末行、首列、末列、斑马行、斑马列） | ❌ |
| **编号样式** | 列表样式定义 | ❌ |
| **链接样式** | 段落+字符双属性样式 | ❌ |
| **样式继承链** | w:basedOn (val: 基于哪个样式) | ❌ |
| **后续段落样式** | w:next (val: 下一段自动应用样式) | ❌ |
| **样式快捷键** | w:qFormat (主样式), w:uiPriority (优先级), w:keyCmd (快捷键) | ❌ |
| **样式隐藏** | w:hidden, w:semiHidden | ❌ |
| **默认样式** | w:docDefaults (w:rPrDefault, w:pPrDefault) | ❌ |
| **样式ID vs 名称** | w:styleId (内部ID) 和 w:name (显示名称) 的映射 | ❌ |
| **样式中的格式** | 样式中包含的 rPr/pPr/tblPr 等完整格式定义 | ❌ |
| **主题样式引用** | 样式中引用主题色/主题字体 | ❌ |
| **LatentStyles** | 潜在样式信息（是否自动启用） | ❌ |

---

## 3. word/numbering.xml — 编号定义

| 属性 | 说明 |
|------|------|
| **抽象编号定义** | w:abstractNum (可被多个编号实例引用) | ❌ |
| **编号实例** | w:num (引用 w:abstractNumId 并添加覆盖) | ❌ |
| **每级编号格式** | w:lvl (w:start, w:numFmt, w:lvlText, w:lvlJc, w:pPr, w:rPr) | ❌ |
| **编号模板** | w:lvlText (val: "%1", "%2" 等占位符) | ❌ |
| **编号样式链接** | w:numStyleLink | ❌ |
| **编号字体** | 每级编号的字体可独立设置 | ❌ |
| **手写编号覆盖** | w:lvlOverride / w:startOverride | ❌ |

---

## 4. word/header*.xml + word/footer*.xml — 页眉页脚

| 属性 | 状态 | 说明 |
|------|------|------|
| **页眉内容** | ⚠️ | Phase 12 标记了，但需确认提取代码存在 |
| **页脚内容** | ⚠️ | Phase 12 标记了，但需确认提取代码存在 |
| **首页不同** | w:sectPr/w:titlePg | ❌ | 首页是否使用独立页眉页脚 |
| **奇偶页不同** | w:sectPr 的 evenAndOddHeaders | ❌ | 奇偶页使用不同页眉页脚 |
| **链接到前节** | w:headerReference/w:footerReference 的 type 属性 | ❌ | 是否继承上节设置 |
| **页眉中的图片** | 页眉区域的图片（Logo、水印等） | ❌ | 未处理 |
| **页眉中的表格** | 页眉中的表格 | ❌ | 未处理 |
| **页眉中的页码** | 页眉中的 PAGE 域 | ❌ | 未处理 |
| **页眉中的域** | 页眉中各种域代码 | ❌ | 未处理 |
| **页眉高度** | w:headerReference 的 空间 | ❌ | 页眉区域大小 |
| **页脚边距** | w:sectPr/w:pgMar (header, footer) | ❌ | 页眉页脚到边界的距离 |

---

## 5. word/footnotes.xml + word/endnotes.xml — 脚注尾注

| 属性 | 状态 | 说明 |
|------|------|------|
| **脚注文本** | ✅ | 已提取文本 |
| **尾注文本** | ❌ | 完全未实现 |
| **脚注编号格式** | w:footnotePr/w:numFmt (decimal, upperRoman, lowerRoman, upperLetter, lowerLetter, etc.) | ❌ |
| **脚注起始编号** | w:footnotePr/w:numStart | ❌ |
| **脚注重新开始** | w:footnotePr/w:footnote (每页/每节/每文档) | ❌ |
| **脚注位置** | w:footnotePr/w:pos (pageBottom, beneathText, endOfSection, endOfDoc) | ❌ |
| **脚注分隔符** | w:footnote/w:separator / w:continuationSeparator | ❌ |
| **脚注延续通知** | w:footnote/w:continuationNotice | ❌ |
| **脚注文本格式** | 脚注中的粗体/斜体/颜色/字体等完整格式 | ❌ | 当前仅提取纯文本 |

---

## 6. word/comments.xml — 批注

| 属性 | 状态 | 说明 |
|------|------|------|
| **批注文本** | ✅ | 已提取 |
| **批注作者** | ✅ | 已提取 |
| **批注日期** | ✅ | 已提取 |
| **批注 ID** | ✅ | 已提取 |
| **批注初始化** | w:comment (w:initials) | ❌ | 批注人缩写 |
| **批注回复** | w:comment (w:parentCommentId) | ❌ | 批注的回复/线程关系 |
| **批注解析状态** | w:comment (w:done) | ❌ | 批注是否已解决 |
| **批注文本格式** | 批注中的粗体/斜体/颜色等完整格式 | ❌ | 当前仅提取纯文本 |
| **批注范围** | w:commentRangeStart / w:commentRangeEnd (标记批注覆盖的文本范围) | ❌ | 未提取批注标记范围 |
| **批注引用** | w:commentReference (段落中的引用标记) | ❌ |
| **批注关联的移动** | 与修订关联的批注 | ❌ |

---

## 7. word/media/ — 媒体资源

| 资源类型 | 说明 |
|----------|------|
| **图片 (PNG/JPEG/GIF/BMP/TIFF/SVG)** | ⚠️ | 当前提取了基本图片，格式映射较简单 |
| **SVG 矢量图** | support/ 中的 SVG 文件 | ❌ |
| **EMF/WMF 图元文件** | 增强图元文件/Windows 图元文件 | ❌ |
| **视频/音频** | 嵌入的媒体文件 | ❌ |
| **3D 模型** | 嵌入的 3D 模型 | ❌ |

---

## 8. word/embeddings/ — 嵌入对象 (OLE)

| 对象类型 | 说明 |
|----------|------|
| **嵌入 Excel** | OLEObject 嵌入式 Excel 工作表 | ❌ |
| **嵌入 PDF** | OLE 嵌入的 PDF 文档 | ❌ |
| **嵌入 PPT** | OLE 嵌入的 PowerPoint | ❌ |
| **嵌入 Visio** | OLE 嵌入的 Visio 图表 | ❌ |
| **通用 OLE 对象** | 其他 OLE 嵌入对象 | ❌ |
| **OLE 图标** | 嵌入对象的显示图标和标签 | ❌ |
| **OLE 链接** | 链接到外部文件（非嵌入） | ❌ |
| **包附件** | 作为包嵌入的任意文件 | ❌ |

---

## 9. word/theme/ — 主题

| 属性 | 说明 |
|------|------|
| **主题颜色** | a:themeClr (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) | ❌ |
| **主题色变体** | 每个主题色有 shade/tint 的变体颜色 | ❌ |
| **主题字体** | a:majorFont (latin, ea, cs) + a:minorFont | ❌ |
| **主题效果** | a:effectStyle (填充、线条、阴影、发光、3D 预设) | ❌ |
| **主题格式方案** | a:fmtScheme (填充方案、线条方案、效果方案) | ❌ |
| **主题背景填充** | 背景样式集合 | ❌ |
| **扩展主题** | word/theme/themeOverride.xml | ❌ |

---

## 10. word/settings.xml — 文档设置

| 设置项 | 说明 |
|--------|------|
| **默认制表位** | w:defaultTabStop | ❌ |
| **自动连字符** | w:autoHyphenation | ❌ |
| **兼容性设置** | w:compat (w:compatSetting) | ❌ |
| **字符间距默认** | w:defaults/w:rPrDefault | ❌ |
| **段落默认** | w:defaults/w:pPrDefault | ❌ |
| **修正模式** | w:trackRevisions | ❌ |
| **修订保护** | w:documentProtection (w:edit, w:format, w:enforcement) | ❌ |
| **只读模式** | w:documentProtection (w:edit="readOnly") | ❌ |
| **仅批注模式** | w:documentProtection (w:edit="comments") | ❌ |
| **仅填写表单** | w:documentProtection (w:edit="forms") | ❌ |
| **保护密码(哈希)** | w:documentProtection (w:cryptProviderType, w:cryptAlgorithmSid, w:hash, w:salt) | ❌ |
| **显示设置** | w:view (打印/页面视图/大纲等), w:zoom (缩放) | ❌ |
| **背景** | w:background (颜色、图片) | ❌ |
| **打印设置** | w:printSettings | ❌ |
| **拼写/语法检查** | w:proofState (spelling, grammar) | ❌ |
| **嵌入字体** | w:embedFonts, w:embedTrueTypeFonts, w:embedSystemFonts | ❌ |
| **文档变量** | w:docVars (w:docVar: name, val) | ❌ |
| **自动保存/恢复** | w:autoSaveDef, w:savePreviewPicture, w:saveXmlDataOnly | ❌ |
| **数学设置** | w:mathPr (公式字体、公式对齐) | ❌ |
| **阅读模式** | w:readModeInkLockup | ❌ |
| **主题引用** | w:themeFontLang, w:decimalSymbol, w:listSeparator | ❌ |

---

## 11. word/fontTable.xml — 字体表

| 属性 | 说明 |
|------|------|
| **字体名称** | w:font (w:name) | ❌ |
| **字体替换** | w:font/w:altName | ❌ |
| **字体嵌入** | w:font/w:embedRegular, w:embedBold, w:embedItalic, w:embedBoldItalic | ❌ |
| **字体字符集** | w:font/w:charset | ❌ |
| **字体族** | w:font/w:family (roman, swiss, modern, script, decorative) | ❌ |
| **字体间距** | w:font/w:pitch (fixed, variable, default) | ❌ |
| **Panose 分类** | w:font/w:panose1 | ❌ |

---

## 12. word/glossary/ — 构建基块

| 类型 | 说明 |
|------|------|
| **自动图文集** | 可重用的文档片段 | ❌ |
| **文档部件** | 结构化文档部件 | ❌ |
| **构建基块条目** | 按类别/类型/名称组织的基块 | ❌ |
| **构建基块行为** | 插入行为（页面/段落/内容） | ❌ |

---

## 13. word/charts/ — 图表

| 类型 | 说明 |
|------|------|
| **柱状图** | 簇状、堆积、百分比堆积、3D | ❌ |
| **折线图** | 带标记、堆积、3D | ❌ |
| **饼图** | 饼图、环形图、复合饼图 | ❌ |
| **条形图** | 簇状、堆积、3D | ❌ |
| **面积图** | 面积、堆积、3D | ❌ |
| **散点图** | 仅标记、带平滑线、带直线 | ❌ |
| **股价图** | 高低收盘、成交量 | ❌ |
| **雷达图** | 雷达、填充雷达 | ❌ |
| **组合图** | 多种图表类型混合 | ❌ |
| **图表数据** | 图表引用的数据表（可能在 Excel 嵌入对象中） | ❌ |
| **图表样式** | 颜色、字体、图例、轴线、网格线、数据标签 | ❌ |
| **图表标题** | 图表标题、轴标题 | ❌ |
| **图表格式化** | 填充、边框、效果的完整 DrawingML 支持 | ❌ |

---

## 14. word/diagrams/ — 图表数据 (SmartArt)

| 类型 | 说明 |
|------|------|
| **SmartArt 图形** | 列表、流程、循环、层次结构、关系、矩阵、棱锥图等 | ❌ |
| **SmartArt 颜色** | 预设颜色方案 | ❌ |
| **SmartArt 样式** | 3D/简单/填充等预设样式 | ❌ |
| **SmartArt 文本** | 每个形状中的文本 | ❌ |
| **SmartArt 布局** | 数据驱动布局模型（dgm:layoutNode） | ❌ |

---

## 15. word/activeX*.xml — ActiveX 控件

| 类型 | 说明 |
|------|------|
| **ActiveX 控件** | 按钮、文本框、复选框、列表框等传统 ActiveX 控件 | ❌ |
| **控件属性** | 控件的 CLSID、属性、事件 | ❌ |
| **控件持久化** | 控件的二进制数据 | ❌ |
| **表单控件** | 较旧的表单控件支持 | ❌ |
| **复选框表单域** | w:ffData/w:checkBox | ❌ |
| **文本表单域** | w:ffData/w:textInput | ❌ |
| **下拉表单域** | w:ffData/w:dropDown | ❌ |

---

## 16. docProps/ — 文档属性

### 16.1 docProps/core.xml — 核心属性

| 属性 | 状态 | 说明 |
|------|------|------|
| **标题** | ⚠️ | 可能已提取 |
| **主题** | ❌ | 未实现 |
| **作者** | ⚠️ | 可能已提取 |
| **最后修改者** | ❌ | 未实现 |
| **创建日期** | ❌ | 未实现 |
| **修改日期** | ❌ | 未实现 |
| **类别** | ❌ | 未实现 |
| **关键词/标签** | ❌ | 未实现 |
| **描述/备注** | ❌ | 未实现 |
| **版本号** | ❌ | 未实现 |
| **修订号** | ❌ | 未实现 |
| **内容状态** | ❌ | 草稿/最终/已发布等 |
| **内容类型** | ❌ | 文档类型标识 |
| **标识符** | ❌ | URI 标识符 |
| **语言** | ❌ | 文档语言 |

### 16.2 docProps/app.xml — 应用属性

| 属性 | 说明 |
|------|------|
| **模板** | 使用的模板文件 | ❌ |
| **页数** | 文档页数 | ❌ |
| **字数** | 统计字数 | ❌ |
| **字符数** | 统计字符数 | ❌ |
| **行数** | 统计行数 | ❌ |
| **段落数** | 统计段落数 | ❌ |
| **应用程序** | 创建文档的应用程序（如 Microsoft Word） | ❌ |
| **应用程序版本** | 应用程序版本号 | ❌ |
| **文档安全** | 安全级别（无/密码保护/只读等） | ❌ |
| **公司** | 公司名称 | ❌ |
| **管理器** | 管理器名称 | ❌ |
| **演示文稿格式** | 演示文稿目标格式 | ❌ |
| **幻灯片数** | 幻灯片数量 | ❌ |
| **隐藏幻灯片数** | 隐藏幻灯片 | ❌ |
| **总编辑时间** | 累计编辑时间 | ❌ |
| **文档管理** | 文档管理属性 | ❌ |

### 16.3 docProps/custom.xml — 自定义属性

| 属性 | 说明 |
|------|------|
| **自定义属性** | 用户自定义的任意属性（名称、类型、值） | ❌ |
| **属性类型** | 文本、数字、日期、是/否 | ❌ |
| **属性链接到内容** | 链接到文档中的内容控件或书签 | ❌ |

---

## 17. word/vbaProject.bin — 宏

| 内容 | 说明 |
|------|------|
| **VBA 项目** | 包含模块、类、表单的 VBA 宏 | ❌ |
| **宏代码** | VBA 源代码 | ❌ |
| **数字签名** | 宏的数字签名证书 | ❌ |
| **宏安全** | 宏安全设置 | ❌ |

---

## 18. word/people.xml — 人员

| 属性 | 说明 |
|------|------|
| **人员信息** | 文档中涉及的所有人员（作者、审阅者、批注者） | ❌ |
| **人员 ID** | 每个人员的唯一标识符 | ❌ |
| **人员名称** | 显示名称 | ❌ |
| **人员电子邮件** | 电子邮件地址 | ❌ |
| **人员图片** | 联系人的头像图片引用 | ❌ |

---

## 19. word/revisions/ — 修订

| 类型 | 说明 |
|------|------|
| **插入修订** | w:ins (插入的文本，含作者、日期、修订ID) | ❌ |
| **删除修订** | w:del (删除的文本，含作者、日期、修订ID) | ❌ |
| **移动修订** | w:moveFrom / w:moveTo (移动文本) | ❌ |
| **格式修订** | w:rPrChange / w:pPrChange / w:tblPrChange (格式变更) | ❌ |
| **段落属性修订** | 段落级属性变更追踪 | ❌ |
| **节属性修订** | 节级属性变更追踪 | ❌ |
| **表格属性修订** | 表格属性变更追踪 | ❌ |
| **修订作者信息** | 每个修订的作者、日期、ID | ❌ |
| **修订接受/拒绝** | 修订的状态（是否已接受/拒绝） | ❌ |
| **修订 ID** | 修订的唯一标识符 | ❌ |
| **修订分页** | w:revisions + w:rPrChange 等 | ❌ |
| **修订父级** | 段落修订的段落ID引用 | ❌ |

---

## 20. word/mailMerge/ — 邮件合并

| 属性 | 说明 |
|------|------|
| **数据源** | 邮件合并的数据源连接 | ❌ |
| **查询字符串** | 数据筛选查询 | ❌ |
| **收件人** | 收件人列表 | ❌ |
| **合并字段映射** | 字段名映射 | ❌ |
| **邮件合并主文档** | 主文档类型（信函、电子邮件、信封、标签、目录） | ❌ |
| **合并域** | 文档中的 MERGEFIELD 域 | ❌ |
| **条件语句** | IF/THEN/ELSE 等条件合并域 | ❌ |
| **地址块** | 标准地址块设置 | ❌ |
| **问候行** | 标准问候行设置 | ❌ |
| **排除数据源** | 排除的收件人来源 | ❌ |

---

## 21. 跨文档/多节属性

### 21.1 节属性 (sectPr)

| 属性 | 状态 | 说明 |
|------|------|------|
| **页面宽度** | ❌ | 页面尺寸 |
| **页面高度** | ❌ | 页面尺寸 |
| **页面方向** | ❌ | 纵向/横向 |
| **页边距(上下左右)** | ❌ | 页边距 |
| **装订线** | ❌ | 装订线位置和大小 |
| **页眉页脚边距** | ❌ | 页眉页脚到页边距 |
| **页面边框** | ❌ | 页面级艺术边框 |
| **分栏** | ⚠️ | Phase 12 标记，需确认实现程度 |
| **分栏数** | ❌ | 栏数 |
| **分栏宽度** | ❌ | 每栏宽度 |
| **分栏间距** | ❌ | 栏间间距 |
| **分栏分割线** | ❌ | 栏间竖线 |
| **行号** | ❌ | 行号设置 |
| **垂直对齐** | ❌ | 页面内容的垂直对齐 |
| **节类型** | ❌ | 连续/下一页/奇数页/偶数页 |
| **页码格式** | ❌ | 页码格式 |
| **页码起始编号** | ❌ | 页码起始值 |
| **纸张来源** | ❌ | 首页/其他页的纸张来源 |
| **页面边框(艺术)** | ❌ | 艺术型页面边框 |
| **页面边框偏移** | ❌ | 页面边框到页边距的距离 |
| **页面边框应用范围** | ❌ | 本节/整篇文档 |

### 21.2 水印

| 属性 | 状态 | 说明 |
|------|------|------|
| **文字水印** | ⚠️ | Phase 13 标记，需确认实现 |
| **图片水印** | ❌ | 图片水印 |
| **水印透明度** | ❌ | 水印透明度 |
| **水印缩放** | ❌ | 水印缩放比例 |
| **水印位置** | ❌ | 水印是否 washout |

### 21.3 文档保护

| 类型 | 状态 | 说明 |
|------|------|------|
| **只读保护** | ⚠️ | Phase 13 检测，但还原不工作 |
| **仅批注保护** | ❌ | 仅允许批注 |
| **仅填写表单保护** | ❌ | 仅允许填写表单 |
| **仅允许修订** | ❌ | 仅允许使用修订进行编辑 |
| **密码哈希/盐值** | ❌ | 保护密码的加密存储 |
| **加密算法类型** | ❌ | SHA-1 / SHA-512 等加密算法 |
| **格式限制** | ❌ | 限制格式更改 |
| **例外列表** | ❌ | 允许编辑的例外用户/组 |

---

## 22. 元数据格式设计建议

### 22.1 总体架构

```
docx → Z 解压 → 逐 XML 部分解析 → 结构化元数据 (YAML) → 渲染器全新构建 docx
```

### 22.2 推荐元数据顶层结构

```yaml
# 顶层
document:
  metadata:              # docProps/* 所有属性
  settings:              # word/settings.xml 所有设置
  theme:                 # word/theme/theme1.xml 主题
  styles:                # word/styles.xml 完整样式表
  numbering:             # word/numbering.xml 编号定义
  fonts:                 # word/fontTable.xml 字体表
  sections: []           # 节列表（每节有独立的 sectPr）
  body: []               # 正文元素序列
  headers_footers: {}    # 页眉页脚（按类型/节索引）
  footnotes:             # word/footnotes.xml
  endnotes:              # word/endnotes.xml
  comments:              # word/comments.xml
  people:                # word/people.xml
  revisions:             # 修订追踪
  mail_merge:            # 邮件合并设置
  glossary:              # 构建基块
  vba_project:           # 宏
  # 二进制资源
  resources:
    images: []           # 图片列表（base64 或文件引用）
    embeddings: []       # 嵌入对象
    media: []            # 其他媒体
```

### 22.3 建议优先级

**第一优先级（必须覆盖，否则还原会明显不一致）：**
1. 字体颜色 (w:color)
2. 单元格合并 (gridSpan/vMerge)
3. 图片环绕/位置 (wp:anchor)
4. 表格完整边框/底纹
5. 段落缩进/间距
6. 公式反向提取 (OMML → LaTeX)

**第二优先级（常见功能，影响文档外观）：**
7. 段落底纹和边框
8. 表格宽度/行高/单元格宽度
9. 页眉页脚完整内容（含图片/表格/域）
10. 列表编号格式
11. 高亮/底纹
12. 内容控件
13. 超链接正向还原到 docx

**第三优先级（专业文档需求）：**
14. 修订追踪
15. 页眉页脚奇偶页/首页不同
16. 主题颜色/字体引用
17. 域代码
18. 多级编号样式
19. 图片效果（旋转/裁剪/阴影）
20. 文档保护

**第四优先级（特殊场景）：**
21. 图表
22. SmartArt
23. 嵌入对象 (OLE)
24. 宏
25. 邮件合并
26. 表单域
27. 3D 模型
28. 视频/音频

---

## 23. 闭环验证策略

### 23.1 验证层级

| 层级 | 方法 | 工具 |
|------|------|------|
| **L0 - 结构相等** | 元数据 JSON 结构比对 | Python dict diff |
| **L1 - XML 相等** | 解压 docx 后逐 XML 文件语义比对（忽略 ID 差异） | lxml 树对比 |
| **L2 - 视觉相等** | docx→PDF→图片→逐像素对比 | `docx2img.py` 已有 |
| **L3 - 功能相等** | 用 Word COM 打开原始/还原 docx，比较页数、字数、可编辑性 | win32com |
| **L4 - 属性相等** | 用 python-docx 读取 docx 属性逐项比对 | 属性遍历脚本 |

### 23.2 测试策略

- **单元测试**: 每个元数据类型一个测试用例，验证提取→还原的闭环
- **黄金文件测试**: 手工创建含各种特性的 docx 作为黄金文件，定期回归
- **模糊测试**: 随机生成各种极端组合的 docx 测试边界情况
- **真实文档测试**: 用真实世界中的复杂 docx 文档测试兼容性

---

> **总结**: 当前项目约覆盖了 docx 全部功能的 **20-25%**。
> 已实现的偏重"文本内容和基础排版"，大量格式细节（颜色、间距、边框、图片位置、表格结构、域代码、图表、修订、主题等）完全未覆盖。
> 核心架构转型（从"两条独立路径"到"统一元数据层"）是最大工程挑战。