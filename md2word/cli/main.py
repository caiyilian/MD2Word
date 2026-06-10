from __future__ import annotations
import argparse
import sys

from md2word import MD2Word
from md2word.exceptions import MD2WordError


def build_parser():
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Convert Markdown files to Word (.docx) documents.",
    )
    parser.add_argument("input", help="Path to input Markdown file (.md)")
    parser.add_argument("-o", "--output", help="Path to output Word file (.docx)")
    parser.add_argument("--font-name", default="等线", help="Document font name (default: 等线)")
    parser.add_argument("--font-size", type=int, default=12, help="Document font size in pt (default: 12)")
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    input_path = args.input
    output_path = args.output

    if output_path is None:
        if input_path.lower().endswith(".md"):
            output_path = input_path[:-3] + ".docx"
        else:
            output_path = input_path + ".docx"

    try:
        converter = MD2Word(font_name=args.font_name, font_size=args.font_size)
        result = converter.convert_file(input_path, output_path)
        print(f"Converted: {input_path} -> {result.path}")
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        sys.exit(1)
    except MD2WordError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
