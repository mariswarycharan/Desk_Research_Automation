"""
PDF Research Paper Image Extractor using Docling + Ollama Relevance Filter
---------------------------------------------------------------------------
Extracts figures from a research paper PDF, then uses an Ollama LLM to check
if each figure's caption is relevant to the research topic before saving.

Usage:
  1. Set PDF_PATH below to your research paper PDF.
  2. pip install docling ollama
  3. python pdf_image_extractor.py
"""

import logging
import re
import time
from pathlib import Path

from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
from docling.document_converter import DocumentConverter, PdfFormatOption
from docling_core.types.doc import PictureItem, TextItem
from ollama import Client


# ╔══════════════════════════════════════════════════════════════╗
# ║  CONFIGURATION                                              ║
# ╚══════════════════════════════════════════════════════════════╝

PDF_PATH = r"C:\Users\vijay\Downloads\Informe_ANUAL_SiVIRA_2023-24_20250211_translated.pdf"

# Minimum image dimensions (pixels) to filter out logos/icons
MIN_WIDTH = 200
MIN_HEIGHT = 150

# Caption patterns — text items starting with these are figure captions
CAPTION_PATTERNS = [
    r"^Figure\s+\d+",       # English: "Figure 1. ..."
    r"^Figura\s+\d+",       # Spanish: "Figura 1. ..."
    r"^Fig\.\s*\d+",        # Abbreviated: "Fig. 1 ..."
    r"^Gr[aá]fico\s+\d+",   # Spanish: "Grafico/Gráfico 1. ..."
    r"^Chart\s+\d+",        # "Chart 1. ..."
    r"^Image\s+\d+",        # "Image 1. ..."
    r"^Illustration\s+\d+", # "Illustration 1. ..."
]

CAPTION_RE = re.compile("|".join(CAPTION_PATTERNS), re.IGNORECASE)

# ── Ollama Config ──
OLLAMA_HOST = "https://ollama.com"
OLLAMA_API_KEY = "dd456319486541c5bfd8dfd001136b32.krWKOjU-1cHaHoKQr0iQDe0r"
OLLAMA_MODEL = "gemma3:12b"  # Available: gemma3:4b, gemma3:12b, gemma3:27b, gemma4:31b

# ── Research context for relevance checking ──
PRIMARY_ASK = (
    "Evaluate the expansion potential of Chinese IVD suppliers in Europe "
    "with focus on access and reimbursement, budget structures, regulatory "
    "requirements, and customer receptivity."
)

SUPPORTING_ASKS = [
    "European IVD market represents high-value opportunity with EUR 1.65T healthcare expenditure",
    "Chinese suppliers face significant regulatory barriers with IVDR transition extending certification by 6-12 months",
    "Geopolitical factors creating new challenges with EU restricting Chinese participation in contracts >EUR 5M",
    "Leading Chinese companies (Mindray, Snibe, Autobio) achieving IVDR certification and showing strong overseas growth",
    "Strategic acquisitions (Mindray/DiaSys, Mindray/HyTest) key to overcoming trust deficit and supply chain risks",
    "Cost-effectiveness drives adoption in low-mid segments, but high-end market remains dominated by Western incumbents",
]


# ╔══════════════════════════════════════════════════════════════╗
# ║  HELPER FUNCTIONS                                           ║
# ╚══════════════════════════════════════════════════════════════╝

def sanitize_filename(text: str, max_len: int = 150) -> str:
    """Clean arbitrary text for use as a filename."""
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


def check_relevance_with_ollama(client: Client, caption: str) -> tuple:
    """
    Send the image caption to Ollama to check if it's relevant to the research.
    Returns (is_relevant: bool, reason: str).
    """
    supporting_asks_text = "\n".join(f"  - {ask}" for ask in SUPPORTING_ASKS)

    prompt = f"""You are a research analyst. You are evaluating whether a figure/image from a research paper could be useful for a specific research project.

PRIMARY RESEARCH QUESTION:
{PRIMARY_ASK}

SUPPORTING RESEARCH THEMES:
{supporting_asks_text}

FIGURE DESCRIPTION:
"{caption}"

TASK: Determine if this figure could even SLIGHTLY contribute to or add context to any aspect of the research question or supporting themes above. Be INCLUSIVE - if there is any reasonable connection (e.g., healthcare data, market data, epidemiological trends that could inform market analysis, regulatory data, European health systems data, diagnostic/IVD related content), mark it as relevant.

Respond with EXACTLY this format (no extra text):
RELEVANT: YES or NO
REASON: One brief sentence explaining why.
"""

    try:
        response = client.chat(
            model=OLLAMA_MODEL,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0.1},  # low temp for consistent decisions
        )
        answer = response.message.content.strip()

        # Parse the response
        is_relevant = "RELEVANT: YES" in answer.upper()
        reason_match = re.search(r"REASON:\s*(.+)", answer, re.IGNORECASE)
        reason = reason_match.group(1).strip() if reason_match else answer

        return is_relevant, reason

    except Exception as e:
        print(f"    WARNING: Ollama API error: {e}")
        # On error, include the image to be safe
        return True, f"Included by default (API error: {str(e)[:50]})"


# ╔══════════════════════════════════════════════════════════════╗
# ║  MAIN EXTRACTION                                            ║
# ╚══════════════════════════════════════════════════════════════╝

def extract_images(pdf_path: str):
    """Main extraction function."""
    pdf = Path(pdf_path)
    if not pdf.exists():
        print(f"ERROR: File not found -> {pdf}")
        return

    # ── Output folder = PDF stem + _trial_ver ──
    output_dir = pdf.parent / (pdf.stem + "_trial_ver")
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {output_dir}\n")

    # ── Initialize Ollama client ──
    print("Connecting to Ollama...")
    ollama_client = Client(
        host=OLLAMA_HOST,
        headers={"Authorization": f"Bearer {OLLAMA_API_KEY}"}
    )
    print("Ollama client ready.\n")

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
    captions_by_page = {}

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
    pictures_by_page = {}

    for item, _level in doc.iterate_items():
        if not isinstance(item, PictureItem):
            continue
        if not item.image:
            continue
        w = item.image.size.width if item.image.size else 0
        h = item.image.size.height if item.image.size else 0
        if w < MIN_WIDTH or h < MIN_HEIGHT:
            continue

        page = get_page_no(item)
        vert = get_vertical_center(item)
        if page not in pictures_by_page:
            pictures_by_page[page] = []
        pictures_by_page[page].append((vert, item))

    total_pics = sum(len(v) for v in pictures_by_page.values())
    print(f"Found {total_pics} figure-sized image(s) (>={MIN_WIDTH}x{MIN_HEIGHT}px).\n")

    # ── Step 3: Match captions to pictures ──
    matched_figures = []  # list of (page_no, caption_text, pic_item)
    unmatched_captions = []

    for page_no, caption_list in sorted(captions_by_page.items()):
        page_pics = pictures_by_page.get(page_no, [])
        if not page_pics:
            for _, cap_text in caption_list:
                unmatched_captions.append((page_no, cap_text))
            continue

        for cap_vert, cap_text in caption_list:
            best_pic = None
            best_dist = float('inf')
            for pic_vert, pic_item in page_pics:
                dist = abs(pic_vert - cap_vert)
                if dist < best_dist:
                    best_dist = dist
                    best_pic = pic_item

            if best_pic is None:
                unmatched_captions.append((page_no, cap_text))
            else:
                matched_figures.append((page_no, cap_text, best_pic))

    print(f"Matched {len(matched_figures)} figure(s) to captions.\n")

    # ── Step 4: LLM relevance check via Ollama ──
    print("=" * 60)
    print("STEP 4: Checking relevance with Ollama LLM...")
    print(f"Model: {OLLAMA_MODEL}")
    print(f"Research focus: {PRIMARY_ASK[:80]}...")
    print("=" * 60 + "\n")

    saved_count = 0
    skipped_count = 0
    used_names = set()

    for i, (page_no, cap_text, pic_item) in enumerate(matched_figures, 1):
        print(f"  [{i}/{len(matched_figures)}] Checking: {cap_text[:90]}...")

        # Ask Ollama if this figure is relevant
        is_relevant, reason = check_relevance_with_ollama(ollama_client, cap_text)

        if not is_relevant:
            skipped_count += 1
            print(f"    -> SKIPPED (not relevant): {reason}")
            continue

        # Build filename from caption
        base_name = sanitize_filename(cap_text)
        final_name = base_name
        counter = 2
        while final_name.lower() in used_names:
            final_name = f"{base_name} ({counter})"
            counter += 1
        used_names.add(final_name.lower())

        filename = f"{final_name}.png"

        # Save the image
        image = pic_item.get_image(doc)
        if image:
            save_path = output_dir / filename
            image.save(save_path, format="PNG")
            saved_count += 1
            print(f"    -> SAVED: {filename}")
            print(f"       Reason: {reason}")
        else:
            print(f"    -> WARNING: Could not extract image data")

    # ── Summary ──
    print(f"\n{'='*60}")
    print(f"RESULTS SUMMARY")
    print(f"{'='*60}")
    print(f"Total captions found     : {total_captions}")
    print(f"Large images found       : {total_pics}")
    print(f"Caption-image matches    : {len(matched_figures)}")
    print(f"LLM: Relevant (saved)    : {saved_count}")
    print(f"LLM: Not relevant (skip) : {skipped_count}")
    if unmatched_captions:
        print(f"Unmatched captions       : {len(unmatched_captions)}")
        for pg, cap in unmatched_captions:
            print(f"  - Page {pg}: {cap[:100]}")
    print(f"Output folder            : {output_dir}")
    print(f"{'='*60}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.WARNING)
    extract_images(PDF_PATH)
