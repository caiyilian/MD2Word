"""
docx → 图片工具

将 docx 每页渲染为 PNG 图片，用于：
1. 可视化查看 docx 内容（无需 Word）
2. 验证 docx→meta→docx 还原的准确性（原始 vs 还原逐页对比）

流程：
  docx → win32com(Word) → PDF → pypdfium2(PDFium) → PNG 图片
"""

from __future__ import annotations
import os
import sys
import tempfile
from pathlib import Path
from typing import List, Optional, Tuple


def docx_to_pdf(docx_path: str, pdf_path: str) -> str:
    """用 Word (win32com) 将 docx 另存为 PDF。"""
    abs_docx = os.path.abspath(docx_path)
    if not os.path.exists(abs_docx):
        raise FileNotFoundError(f"Document not found: {abs_docx}")

    try:
        from win32com.client import Dispatch
    except ImportError:
        raise RuntimeError("pywin32 is required. Install: pip install pywin32")

    word = None
    try:
        word = Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(abs_docx)
        # wdFormatPDF = 17
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(False)
        print(f"PDF saved: {pdf_path}")
        return pdf_path
    except Exception as e:
        raise RuntimeError(f"Failed to convert docx to PDF: {e}") from e
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def pdf_to_images(
    pdf_path: str,
    output_dir: str,
    dpi: int = 200,
    image_format: str = "png",
) -> List[str]:
    """用 pypdfium2 将 PDF 每页渲染为图片。"""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise RuntimeError(
            "pypdfium2 is required. Install: pip install pypdfium2"
        )

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(pdf_path)
    n_pages = len(pdf)
    saved: List[str] = []
    pdf_stem = Path(pdf_path).stem

    # Scale factor: 72pt/in → dpi pixels/in
    scale = dpi / 72.0

    for i in range(n_pages):
        page = pdf[i]
        bitmap = page.render(scale=scale)
        # pypdfium2 returns a PIL Image
        pil_image = bitmap.to_pil()
        img_path = output_dir / f"{pdf_stem}_page_{i + 1:03d}.{image_format}"
        pil_image.save(str(img_path))
        saved.append(str(img_path))
        print(f"  Saved: {img_path}")

    pdf.close()
    return saved


def docx_to_images(
    docx_path: str,
    output_dir: str = "output_images",
    dpi: int = 200,
    image_format: str = "png",
) -> List[str]:
    """完整流程：docx → PDF → 多张 PNG。"""
    docx_stem = Path(docx_path).stem

    with tempfile.TemporaryDirectory(prefix="docx2img_") as tmpdir:
        pdf_path = os.path.join(tmpdir, f"{docx_stem}.pdf")
        docx_to_pdf(docx_path, pdf_path)
        images = pdf_to_images(pdf_path, output_dir, dpi=dpi, image_format=image_format)

    print(f"\nDone! {len(images)} images generated in: {output_dir}")
    return images


def compare_docx(
    docx_a: str,
    docx_b: str,
    output_dir: str = "compare_output",
    dpi: int = 200,
) -> bool:
    """将两个 docx 分别转图片，逐页对比，报告差异。"""
    dir_a = os.path.join(output_dir, "original")
    dir_b = os.path.join(output_dir, "restored")

    print(f"Converting: {docx_a}")
    imgs_a = docx_to_images(docx_a, dir_a, dpi=dpi, image_format="png")
    print(f"\nConverting: {docx_b}")
    imgs_b = docx_to_images(docx_b, dir_b, dpi=dpi, image_format="png")

    if len(imgs_a) != len(imgs_b):
        print(f"\n[FAIL] Page count mismatch: {len(imgs_a)} vs {len(imgs_b)}")
        return False

    all_match = True
    from PIL import Image

    for i, (img_a, img_b) in enumerate(zip(imgs_a, imgs_b)):
        a = Image.open(img_a)
        b = Image.open(img_b)
        if a.size != b.size:
            print(f"  Page {i + 1}: [FAIL] Size mismatch {a.size} vs {b.size}")
            all_match = False
            continue
        # Pixel-by-pixel comparison
        pixels_a = list(a.getdata())
        pixels_b = list(b.getdata())
        if pixels_a != pixels_b:
            diff_count = sum(1 for pa, pb in zip(pixels_a, pixels_b) if pa != pb)
            total = len(pixels_a)
            pct = diff_count / total * 100
            print(f"  Page {i + 1}: [FAIL] {diff_count}/{total} pixels differ ({pct:.2f}%)")
            # Save diff image
            diff_img = Image.new("RGB", a.size)
            diff_pixels = [
                (255, 0, 0) if pa != pb else (128, 128, 128)
                for pa, pb in zip(pixels_a, pixels_b)
            ]
            diff_img.putdata(diff_pixels)
            diff_path = os.path.join(output_dir, f"diff_page_{i + 1:03d}.png")
            diff_img.save(diff_path)
            print(f"    Diff image saved: {diff_path}")
            all_match = False
        else:
            print(f"  Page {i + 1}: [OK] Identical")

    if all_match:
        print(f"\n[OK] All {len(imgs_a)} pages match perfectly!")
    else:
        print(f"\n[FAIL] Differences found. Check {output_dir}/ for details.")

    return all_match


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert docx pages to images")
    parser.add_argument("input", nargs="?", help="Path to input .docx file")
    parser.add_argument("-o", "--output", default="output_images",
                        help="Output directory for images (default: output_images)")
    parser.add_argument("--dpi", type=int, default=200,
                        help="Image DPI (default: 200)")
    parser.add_argument("--format", default="png",
                        choices=["png", "jpg", "tiff"],
                        help="Image format (default: png)")
    parser.add_argument("--compare", nargs=2, metavar=("DOCX_A", "DOCX_B"),
                        help="Compare two docx files page by page")

    args = parser.parse_args()

    if args.compare:
        compare_docx(args.compare[0], args.compare[1], output_dir=args.output, dpi=args.dpi)
    elif args.input:
        docx_to_images(args.input, args.output, dpi=args.dpi, image_format=args.format)
    else:
        parser.print_help()