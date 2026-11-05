"""Check a deck for repeated compositions before delivery.

Rule this encodes (user, on the lithography deck):
"后面 5/6/7 页用的模板完全相同，很容易让人产生审美疲劳，以后不要这么做，
可以同时使用多个 PPT 模板，但要保持风格一致。"

So: do not run the same layout back to back, and do not let one layout carry the
whole deck. Mixing several templates is allowed — but palette, fonts, title
hierarchy, header/footer and illustration language must stay one system.

Usage:
    python src/check_layout_diversity.py <deck.pptx> [--max-run 2] [--max-share 0.4]
                                         [--json report.json]

--max-run N     fail if N or more consecutive slides share one layout (2 = 3+ in a row)
--max-share F   fail if one layout covers more than F of all slides

Exit codes: 0 = diverse enough, 1 = repetition found.
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

from pptx import Presentation


def signature(slide) -> tuple[str, int, int, int]:
    """Layout name plus counts of text / picture / shape objects."""
    try:
        layout = slide.slide_layout.name
    except Exception:
        layout = "?"
    texts = pics = shapes = 0
    for shape in slide.shapes:
        shapes += 1
        if shape.has_text_frame and shape.text_frame.text.strip():
            texts += 1
        if shape.shape_type is not None and str(shape.shape_type).startswith("PICTURE"):
            pics += 1
    return (layout, texts, pics, shapes)


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    deck = Path(sys.argv[1])
    max_run = 2
    max_share = 0.4
    report_path: Path | None = None
    args = sys.argv[2:]
    for i, arg in enumerate(args):
        if arg == "--max-run" and i + 1 < len(args):
            max_run = int(args[i + 1])
        if arg == "--max-share" and i + 1 < len(args):
            max_share = float(args[i + 1])
        if arg == "--json" and i + 1 < len(args):
            report_path = Path(args[i + 1])

    if not deck.exists():
        print(json.dumps({"status": "FAIL", "errors": [f"missing deck: {deck}"]}))
        return 1

    prs = Presentation(str(deck))
    slides = list(prs.slides)
    total = len(slides)
    if total == 0:
        report = {"deck": str(deck), "slides": 0, "status": "PASS",
                  "notes": ["empty deck"]}
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0

    layouts = [signature(s)[0] for s in slides]
    shapes = [signature(s)[1:] for s in slides]

    # 1) consecutive runs of the same layout
    runs: list[dict] = []
    start = 1
    for index in range(1, total + 1):
        same = index < total and layouts[index] == layouts[index - 1]
        if not same:
            length = index - start + 1
            if length > max_run:
                runs.append({"layout": layouts[start - 1],
                             "slides": list(range(start, index + 1)),
                             "length": length})
            start = index + 1

    # 2) one layout dominating the deck
    counts = Counter(layouts)
    dominating = [
        {"layout": name, "slides": count, "share": round(count / total, 2)}
        for name, count in counts.items()
        if count / total > max_share
    ]

    # 3) neighbouring pages that are the same layout AND the same object mix
    clones = [
        {"slides": [i, i + 1], "layout": layouts[i - 1]}
        for i in range(1, total)
        if layouts[i - 1] == layouts[i] and shapes[i - 1] == shapes[i]
    ]

    ok = not runs and not dominating
    report = {
        "deck": str(deck),
        "slides": total,
        "distinct_layouts": len(counts),
        "max_consecutive_repeats": max_run + 1,
        "consecutive_runs_over_limit": runs,
        "dominant_layouts": dominating,
        "identical_neighbour_pairs": clones,
        "status": "PASS" if ok else "REVIEW",
        "rule": "same layout may not run 3+ pages, or cover >40% of the deck; "
                "mixing templates is fine only with one consistent visual system",
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report_path:
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
