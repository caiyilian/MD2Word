from __future__ import annotations
import argparse
import sys

from md2word import MD2Word
from md2word.exceptions import MD2WordError
from md2word.extractor.docx_extractor import DocxExtractor
from md2word.writer.md_writer import MdWriter


def build_parser():
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Convert between Markdown and Word (.docx) documents.",
    )
    parser.add_argument("input", help="Path to input file (.md or .docx)")
    parser.add_argument("-o", "--output", help="Path to output file")
    parser.add_argument("--font-name", default="\u7b49\u7ebf",
                        help="Document font name (default: \u7b49\u7ebf)")
    parser.add_argument("--font-size", type=int, default=12,
                        help="Document font size in pt (default: 12)")
    parser.add_argument("--count-pages", action="store_true",
                        help="Count pages using Microsoft Word")
    parser.add_argument("--style",
                        help="Path to YAML style config file")
    parser.add_argument("--reverse", "-r", action="store_true",
                        help="Reverse mode: convert .docx to .md")
    return parser


def _do_convert(args):
    input_path = args.input
    output_path = args.output
    if output_path is None:
        if input_path.lower().endswith(".md"):
            output_path = input_path[:-3] + ".docx"
        else:
            output_path = input_path + ".docx"

    converter = MD2Word(
        font_name=args.font_name,
        font_size=args.font_size,
        style_path=args.style,
    )
    result = converter.convert_file(input_path, output_path,
                                    count_pages=args.count_pages)
    if args.count_pages:
        print(f"Converted: {input_path} -> {result.path}")
        print(f"Pages: {result.pages}")
    else:
        print(f"Converted: {input_path} -> {result.path}")


def _do_reverse(args):
    input_path = args.input
    output_path = args.output
    if output_path is None:
        if input_path.lower().endswith(".docx"):
            output_path = input_path[:-5] + ".md"
        else:
            output_path = input_path + ".md"

    extractor = DocxExtractor()
    doc = extractor.extract(input_path)
    writer = MdWriter(default_font_name=args.font_name,
                      default_font_size=args.font_size)
    md_text = writer.write(doc)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(md_text)
    print(f"Converted: {input_path} -> {output_path}")


def main():
    parser = build_parser()
    args = parser.parse_args()

    try:
        if args.reverse:
            _do_reverse(args)
        else:
            _do_convert(args)
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        sys.exit(1)
    except MD2WordError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
