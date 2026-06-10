from __future__ import annotations

from md2word.model.document import Document
from md2word.parser.markdown_parser import MarkdownParser
from md2word.renderer.docx_renderer import DocxRenderer


class ConversionResult:
    def __init__(self, path: str, pages: int = 0):
        self.path = path
        self.pages = pages


class MD2Word:
    def __init__(self, font_name: str = "等线", font_size: int = 12):
        self.parser = MarkdownParser()
        self.renderer = DocxRenderer(font_name=font_name, font_size=font_size)

    def convert(self, md_text: str, output_path: str) -> ConversionResult:
        document = self.parser.parse(md_text)
        self.renderer.render(document, output_path)
        return ConversionResult(path=output_path)

    def convert_file(self, md_path: str, output_path: str) -> ConversionResult:
        with open(md_path, "r", encoding="utf-8") as f:
            md_text = f.read()
        return self.convert(md_text, output_path)
