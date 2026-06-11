from __future__ import annotations
import argparse
import os
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
    parser.add_argument("--ocr", action="store_true",
                        help="Run OCR on images and extract text to md")
    return parser


def build_docx2md_parser():
    parser = argparse.ArgumentParser(
        prog="docx2md",
        description="Convert Word (.docx) documents to Markdown.",
    )
    parser.add_argument("input", help="Path to input .docx file")
    parser.add_argument("-o", "--output", help="Path to output .md file")
    parser.add_argument("--ocr", action="store_true",
                        help="Run OCR on images and extract text to md")
    parser.add_argument("--font-name", default="\u7b49\u7ebf",
                        help="Document font name (default: \u7b49\u7ebf)")
    parser.add_argument("--font-size", type=int, default=12,
                        help="Document font size in pt (default: 12)")
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
    input_base = os.path.splitext(os.path.basename(input_path))[0]
    output_target = args.output

    if output_target is None:
        subfolder = input_base
    elif output_target.lower().endswith(".md"):
        subfolder = os.path.splitext(os.path.basename(output_target))[0]
    else:
        subfolder = os.path.basename(output_target)

    output_dir = os.path.join("output", subfolder)
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    md_filename = subfolder + ".md"
    md_path = os.path.join(output_dir, md_filename)

    extractor = DocxExtractor(output_dir=images_dir, ocr=args.ocr)
    doc = extractor.extract(input_path)
    writer = MdWriter(default_font_name=args.font_name,
                      default_font_size=args.font_size)
    md_text = writer.write(doc)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_text)
    print(f"Converted: {input_path} -> {md_path}")
    img_count = len(extractor._saved_images)
    if img_count:
        print(f"Images extracted: {img_count} (saved to {images_dir}/)")


def _do_docx2md(args):
    """Direct docx to md conversion with -o support."""
    input_path = args.input
    input_base = os.path.splitext(os.path.basename(input_path))[0]
    output_path = args.output

    if output_path is None:
        # Default: output to output/{input_name}/
        output_dir = os.path.join("output", input_base)
        images_dir = os.path.join(output_dir, "images")
        os.makedirs(images_dir, exist_ok=True)
        md_path = os.path.join(output_dir, input_base + ".md")
    elif output_path.lower().endswith(".md"):
        # -o output.md: create folder output/ with the md file
        output_name = os.path.splitext(os.path.basename(output_path))[0]
        output_dir = os.path.join("output", output_name)
        images_dir = os.path.join(output_dir, "images")
        os.makedirs(images_dir, exist_ok=True)
        md_path = os.path.join(output_dir, output_name + ".md")
    else:
        # -o some_dir: use as output directory
        output_dir = output_path
        images_dir = os.path.join(output_dir, "images")
        os.makedirs(images_dir, exist_ok=True)
        md_path = os.path.join(output_dir, input_base + ".md")

    extractor = DocxExtractor(output_dir=images_dir, ocr=args.ocr)
    doc = extractor.extract(input_path)
    writer = MdWriter(default_font_name=args.font_name,
                      default_font_size=args.font_size)
    md_text = writer.write(doc)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_text)
    print(f"Converted: {input_path} -> {md_path}")
    img_count = len(extractor._saved_images)
    if img_count:
        print(f"Images extracted: {img_count} (saved to {images_dir}/)")


def main():
    parser = build_parser()
    args = parser.parse_args()

    try:
        if args.reverse:
            _do_reverse(args)
        elif args.input.lower().endswith(".docx"):
            # Auto-detect: input is .docx, do reverse conversion
            _do_docx2md(args)
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


def docx2md_main():
    """Entry point for docx2md command."""
    parser = build_docx2md_parser()
    args = parser.parse_args()

    try:
        _do_docx2md(args)
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        sys.exit(1)
    except MD2WordError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
