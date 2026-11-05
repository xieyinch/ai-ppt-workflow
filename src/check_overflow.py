"""Detect likely text overflow in a built PPTX before the slides are presented.

Local habit this encodes: every delivered deck is checked for text spilling past
its placeholder ("通过文本溢出检查"), not just glanced at as a PNG.

The estimate measures each paragraph with real font metrics (Windows system
fonts, CJK-aware) against the shape's inner text box, then sums wrapped line
heights. It is a screening tool: a positive hit means "open this page and look",
not proof of a visual defect.

Usage:
    python src/check_overflow.py <deck.pptx> [--limit 1.0] [--json report.json]

Exit codes: 0 = no page over the limit, 1 = at least one page over the limit.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

# Rough line-height multiplier per font size (single spacing, PowerPoint default).
LINE_FACTOR = 1.22
DEFAULT_PT = 18.0
FONT_CANDIDATES = [
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/msyhbd.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/arial.ttf",
]


def _font(size_px: int):
    from PIL import ImageFont
    for path in FONT_CANDIDATES:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size_px)
            except Exception:
                continue
    return ImageFont.load_default()


def _text_width_pt(text: str, size_pt: float) -> float:
    """Advance width of text in points, measured with a real CJK-capable font."""
    font = _font(max(8, int(round(size_pt * 4))))  # measure at 4x for precision
    try:
        return font.getlength(text) / 4.0
    except AttributeError:
        return font.getsize(text)[0] / 4.0


def _paragraph_size(paragraph, fallback: float) -> float:
    for run in paragraph.runs:
        if run.font.size is not None:
            return run.font.size.pt
    return fallback


def _shape_report(shape) -> dict | None:
    if not shape.has_text_frame:
        return None
    tf = shape.text_frame
    text = tf.text.strip()
    if not text:
        return None

    width_pt = Emu(shape.width).pt - (
        Emu(tf.margin_left).pt + Emu(tf.margin_right).pt)
    height_pt = Emu(shape.height).pt - (
        Emu(tf.margin_top).pt + Emu(tf.margin_bottom).pt)
    if width_pt <= 1 or height_pt <= 1:
        return None

    used_pt = 0.0
    for paragraph in tf.paragraphs:
        body = "".join(run.text for run in paragraph.runs) or paragraph.text
        if not body.strip():
            used_pt += _paragraph_size(paragraph, DEFAULT_PT) * LINE_FACTOR
            continue
        size = _paragraph_size(paragraph, DEFAULT_PT)
        unit = _text_width_pt("字", size) or size  # CJK glyph advance
        lines = 0
        remaining = width_pt
        for char in body:
            advance = _text_width_pt(char, size)
            if advance > remaining:
                lines += 1
                remaining = width_pt - advance
            else:
                remaining -= advance
        lines += 1
        used_pt += lines * size * LINE_FACTOR
        _ = unit

    if used_pt <= 0:
        return None
    return {
        "shape": shape.name,
        "used_pt": round(used_pt, 1),
        "box_pt": round(height_pt, 1),
        "ratio": round(used_pt / height_pt, 2),
    }


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    deck = Path(sys.argv[1])
    limit = 1.0
    report_path: Path | None = None
    args = sys.argv[2:]
    for i, arg in enumerate(args):
        if arg == "--limit" and i + 1 < len(args):
            limit = float(args[i + 1])
        if arg == "--json" and i + 1 < len(args):
            report_path = Path(args[i + 1])

    if not deck.exists():
        print(json.dumps({"status": "FAIL", "errors": [f"missing deck: {deck}"]}))
        return 1

    prs = Presentation(str(deck))
    hits = []
    checked = 0
    for number, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            info = _shape_report(shape)
            if not info:
                continue
            checked += 1
            if info["ratio"] > limit:
                info["slide"] = number
                hits.append(info)

    hits.sort(key=lambda h: h["ratio"], reverse=True)
    report = {
        "deck": str(deck),
        "text_shapes_checked": checked,
        "limit": limit,
        "overflow_pages": sorted({h["slide"] for h in hits}),
        "hits": hits,
        "status": "PASS" if not hits else "REVIEW",
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report_path:
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if not hits else 1


if __name__ == "__main__":
    raise SystemExit(main())
