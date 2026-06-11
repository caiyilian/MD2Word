# Phase 4.5 测试 - 文本格式增强

## 删除线

这是 ~~删除线~~ 文本。

这是一段包含 ~~删除线~~ 和**粗体 ~~删除线粗体~~** 的混合文本。

## 下划线

这是 <u>下划线</u> 文本。

这是 <u>下划线</u> 和 *斜体 <u>斜体下划线</u>* 混合。

## 上标与下标

化学式：H<sub>2</sub>O，CO<sub>2</sub>

公式：E=mc<sup>2</sup>，x<sup>2</sup> + y<sup>2</sup> = z<sup>2</sup>

混合：H<sub>2</sub>SO<sub>4</sub>，a<sup>2</sup> + b<sup>2</sup> = c<sup>2</sup>

## 删除线 + 上标/下标混合

这是 ~~删除线<u>下划线</u>H<sub>2</sub>O~~ 全部删除。

**粗体 ~~删除线<sup>上标</sup><sub>下标</sub>~~** 混合。

## 自定义字体与字号

<span style="font-family:Arial">Arial 字体文本</span>

<span style="font-size:24pt">24pt 大字</span>

<span style="font-family:Impact;font-size:18pt">Impact 字体，18pt</span>

<u><span style="font-family:'Times New Roman';font-size:16pt">下划线 + Times New Roman 16pt</span></u>

## 表格中的格式

| 列1 | 列2 | 列3 |
|-----|:---:|:----|
| ~~删除~~ | <u>下划线</u> | H<sub>2</sub>O |
| <sup>上标</sup> | <sub>下标</sub> | **粗体** |
| *斜体* | `代码` | <span style="font-family:Arial">Arial</span> |

## 列表中的格式

- 普通文本
- <u>下划线列表项</u>
- ~~删除线列表项~~
- 包含 H<sub>2</sub>O 和 E=mc<sup>2</sup> 的列表项

1. 第一项
2. ~~第二项~~ <u>有下划线</u>
3. **粗体<span style="font-size:16pt">大字号</span>混合**

## 综合测试

这是一段 ~~删除线~~ <u>下划线</u> **粗体删除线 ~~粗体~~** *斜体 <u>带下划线</u>* 和 H<sub>2</sub>O 以及 E=mc<sup>2</sup> 的复杂混合文本。

<span style="font-family:Consolas;font-size:10pt">代码风格字体（Consolas 10pt）用于普通文本。</span>
