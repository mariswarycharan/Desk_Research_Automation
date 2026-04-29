"""
PDF Research Paper Image Extractor using Docling
-------------------------------------------------
Extracts all figures/images from a research paper PDF.
Output folder is named after the paper.
Each image is named using its figure caption (e.g., "Figure 1. Overview of...").
Only saves figures that have a matching caption nearby.

Usage:
  1. Set PDF_PATH below to your research paper PDF.
  2. pip install docling
  3. python pdf_image_extractor.py
"""

import logging
import re
from pathlib import Path

from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
from docling.document_converter import DocumentConverter, PdfFormatOption
from docling_core.types.doc import PictureItem, TextItem


# ╔══════════════════════════════════════════════════════════════╗
# ║  SET YOUR PDF PATH HERE                                     ║
# ╚══════════════════════════════════════════════════════════════╝
PDF_PATH = r"C:\Users\maris\Downloads\1-s2.0-S168411822100092X-main.pdf"

# Minimum image dimensions (pixels) to filter out logos/icons
MIN_WIDTH = 200
MIN_HEIGHT = 150

# Caption patterns — text items starting with these are figure captions
# Add more patterns if your paper uses different conventions
CAPTION_PATTERNS = [
    r"^Figure\s+\d+",       # English: "Figure 1. ..."
    r"^Figura\s+\d+",       # Spanish: "Figura 1. ..."
    r"^Fig\.\s*\d+",        # Abbreviated: "Fig. 1 ..."
    r"^Gráfico\s+\d+",      # Spanish alt: "Gráfico 1. ..."
    r"^Gr[aá]fico\s+\d+",   # With/without accent
    r"^Chart\s+\d+",        # "Chart 1. ..."
    r"^Image\s+\d+",        # "Image 1. ..."
    r"^Illustration\s+\d+", # "Illustration 1. ..."
]

CAPTION_RE = re.compile("|".join(CAPTION_PATTERNS), re.IGNORECASE)


def sanitize_filename(text: str, max_len: int = 150) -> str:
    """
    Clean arbitrary text so it can be used safely as a filename.
    Removes illegal characters, collapses whitespace, and truncates.
    """
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip(' .')
    if len(text) > max_len:
        text = text[:max_len].rsplit(' ', 1)[0]
    return text if text else "untitled"


def get_page_no(item) -> int:
    """Get the page number from a document item's provenance."""
    if hasattr(item, 'prov') and item.prov:
        return item.prov[0].page_no
    return -1


def get_vertical_center(item) -> float:
    """Get the vertical center of an item's bounding box."""
    if hasattr(item, 'prov') and item.prov:
        bbox = item.prov[0].bbox
        return (bbox.t + bbox.b) / 2.0
    return 0.0


def extract_images(pdf_path: str):
    """Main extraction function."""
    pdf = Path(pdf_path)
    if not pdf.exists():
        print(f"ERROR: File not found → {pdf}")
        return

    # ── Output folder = PDF stem (paper name) ──
    output_dir = pdf.parent / pdf.stem
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {output_dir}\n")

    # ── Configure Docling pipeline (memory-optimized) ──
    pipeline_options = PdfPipelineOptions()
    pipeline_options.generate_picture_images = True
    pipeline_options.images_scale = 1.0
    pipeline_options.do_ocr = False
    pipeline_options.do_table_structure = False

    converter = DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_options
            )
        }
    )

    # ── Convert the PDF ──
    print(f"Processing: {pdf.name} ...")
    result = converter.convert(str(pdf))
    doc = result.document
    print("Conversion complete.\n")

    # ── Step 1: Collect all figure captions, grouped by page ──
    # A caption is a TextItem whose text matches CAPTION_RE
    captions_by_page = {}  # page_no -> list of (vertical_center, caption_text)

    for item, _level in doc.iterate_items():
        if not isinstance(item, TextItem):
            continue
        text = item.text.strip() if hasattr(item, 'text') else ""
        if CAPTION_RE.match(text):
            page = get_page_no(item)
            vert = get_vertical_center(item)
            if page not in captions_by_page:
                captions_by_page[page] = []
            captions_by_page[page].append((vert, text))

    total_captions = sum(len(v) for v in captions_by_page.values())
    print(f"Found {total_captions} figure caption(s) across {len(captions_by_page)} page(s).\n")

    if total_captions == 0:
        print("WARNING: No figure captions found!")
        print("  The caption patterns tried:")
        for p in CAPTION_PATTERNS:
            print(f"    {p}")
        print("\n  You may need to add your paper's caption pattern to CAPTION_PATTERNS.")

        # Print sample text items to help the user identify the pattern
        print("\n  Sample text items from the document:")
        sample_count = 0
        for item, _ in doc.iterate_items():
            if isinstance(item, TextItem):
                text = item.text.strip() if hasattr(item, 'text') else ""
                if len(text) > 20:
                    print(f"    Page {get_page_no(item)}: {text[:120]}")
                    sample_count += 1
                    if sample_count >= 20:
                        break
        return

    # ── Step 2: Collect all "real" pictures (filtered by size) ──
    pictures_by_page = {}  # page_no -> list of (vertical_center, PictureItem)

    for item, _level in doc.iterate_items():
        if not isinstance(item, PictureItem):
            continue
        if not item.image:
            continue
        w = item.image.size.width if item.image.size else 0
        h = item.image.size.height if item.image.size else 0
        if w < MIN_WIDTH or h < MIN_HEIGHT:
            continue  # skip tiny icons/logos

        page = get_page_no(item)
        vert = get_vertical_center(item)
        if page not in pictures_by_page:
            pictures_by_page[page] = []
        pictures_by_page[page].append((vert, item))

    total_pics = sum(len(v) for v in pictures_by_page.values())
    print(f"Found {total_pics} figure-sized image(s) (>={MIN_WIDTH}x{MIN_HEIGHT}px).\n")

    # ── Step 3: Match each caption to its nearest picture on the same page ──
    # Captions usually appear just below the figure, so we look for the
    # closest picture that is ABOVE the caption (higher vertical position).
    saved_count = 0
    unmatched_captions = []
    used_names = set()

    for page_no, caption_list in sorted(captions_by_page.items()):
        page_pics = pictures_by_page.get(page_no, [])
        if not page_pics:
            for _, cap_text in caption_list:
                unmatched_captions.append((page_no, cap_text))
            continue

        for cap_vert, cap_text in caption_list:
            # Find the nearest picture on this page
            best_pic = None
            best_dist = float('inf')
            for pic_vert, pic_item in page_pics:
                dist = abs(pic_vert - cap_vert)
                if dist < best_dist:
                    best_dist = dist
                    best_pic = pic_item

            if best_pic is None:
                unmatched_captions.append((page_no, cap_text))
                continue

            # Build filename from caption
            base_name = sanitize_filename(cap_text)

            final_name = base_name
            counter = 2
            while final_name.lower() in used_names:
                final_name = f"{base_name} ({counter})"
                counter += 1
            used_names.add(final_name.lower())

            filename = f"{final_name}op.png"

            # Save the image
            image = best_pic.get_image(doc)
            if image:
                save_path = output_dir / filename
                image.save(save_path, format="PNG")
                saved_count += 1
                print(f"  [Page {page_no}] Saved: {filename}")
            else:
                print(f"  [Page {page_no}] WARNING: Could not extract image for: {cap_text[:80]}")

    # ── Summary ──
    print(f"\n{'='*60}")
    print(f"Total captions found     : {total_captions}")
    print(f"Large images found       : {total_pics}")
    print(f"Successfully saved       : {saved_count}")
    if unmatched_captions:
        print(f"Unmatched captions       : {len(unmatched_captions)}")
        for pg, cap in unmatched_captions:
            print(f"  - Page {pg}: {cap[:100]}")
    print(f"Output folder            : {output_dir}")
    print(f"{'='*60}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.WARNING)
    extract_images(PDF_PATH)