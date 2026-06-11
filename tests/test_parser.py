from md2word.parser.markdown_parser import (
    MarkdownParser, parse_frontmatter, preprocess_image_attributes,
    _parse_img_attrs,
)
from md2word.model.document import (
    Document, Heading, Paragraph, CodeBlock, TextRun,
    Image, ListBlock, Table, HorizontalRule,
)


def test_parse_frontmatter():
    meta, body = parse_frontmatter("---\ntitle: Test\n---\n# Hello")
    assert meta == {"title": "Test"}
    assert body == "# Hello"


def test_parse_frontmatter_none():
    meta, body = parse_frontmatter("# Hello")
    assert meta == {}
    assert body == "# Hello"


def test_parse_heading():
    p = MarkdownParser()
    doc = p.parse("# Heading 1\n\n## Heading 2")
    assert len(doc.elements) == 2
    assert isinstance(doc.elements[0], Heading)
    assert doc.elements[0].level == 1
    assert doc.elements[0].runs[0].text == "Heading 1"
    assert doc.elements[1].level == 2


def test_parse_paragraph():
    p = MarkdownParser()
    doc = p.parse("Hello world")
    assert len(doc.elements) == 1
    assert isinstance(doc.elements[0], Paragraph)
    assert doc.elements[0].runs[0].text == "Hello world"


def test_parse_bold_italic():
    p = MarkdownParser()
    doc = p.parse("**bold** and *italic*")
    para = doc.elements[0]
    assert isinstance(para, Paragraph)
    assert para.runs[0].bold == True
    assert para.runs[0].text == "bold"
    assert para.runs[2].italic == True
    assert para.runs[2].text == "italic"


def test_parse_inline_code():
    p = MarkdownParser()
    doc = p.parse("Text `code` here")
    para = doc.elements[0]
    assert para.runs[1].code == True
    assert para.runs[1].text == "code"


def test_parse_code_block():
    p = MarkdownParser()
    doc = p.parse("```python\nprint(1)\n```")
    assert len(doc.elements) == 1
    cb = doc.elements[0]
    assert isinstance(cb, CodeBlock)
    assert cb.language == "python"
    assert "print(1)" in cb.code


def test_parse_unordered_list():
    p = MarkdownParser()
    doc = p.parse("- item 1\n- item 2")
    assert len(doc.elements) == 1
    lst = doc.elements[0]
    assert isinstance(lst, ListBlock)
    assert lst.ordered == False
    assert len(lst.items) == 2


def test_parse_ordered_list():
    p = MarkdownParser()
    doc = p.parse("1. first\n2. second")
    lst = doc.elements[0]
    assert isinstance(lst, ListBlock)
    assert lst.ordered == True
    assert len(lst.items) == 2


def test_parse_horizontal_rule():
    p = MarkdownParser()
    doc = p.parse("---")
    assert len(doc.elements) == 1
    assert isinstance(doc.elements[0], HorizontalRule)


def test_parse_image():
    p = MarkdownParser()
    doc = p.parse("![alt](img.png)")
    assert len(doc.elements) == 1
    img = doc.elements[0]
    assert isinstance(img, Image)
    assert img.src == "img.png"
    assert img.alt == "alt"


def test_parse_image_with_attrs():
    p = MarkdownParser()
    doc = p.parse("![](img.png){:width=300px align=center}")
    img = doc.elements[0]
    assert isinstance(img, Image)
    assert img.width == "300px"
    assert img.align == "center"


def test_parse_table():
    p = MarkdownParser()
    doc = p.parse("| h1 | h2 |\n|---|---|\n| a | b |")
    assert len(doc.elements) == 1
    t = doc.elements[0]
    assert isinstance(t, Table)
    assert t.headers == ["h1", "h2"]
    assert t.rows == [["a", "b"]]


def test_parse_image_attrs():
    assert _parse_img_attrs(":width=50% align=center") == {
        "width": "50%", "align": "center"
    }
    assert _parse_img_attrs(":width=300px") == {"width": "300px"}


def test_preprocess_image_attributes():
    text, attrs_map = preprocess_image_attributes(
        "![](a.png){:width=100} ![](b.png)"
    )
    assert attrs_map[0] == {"width": "100"}
    assert attrs_map[1] == {}
    assert "![](a.png)" in text
    assert "![](b.png)" in text
