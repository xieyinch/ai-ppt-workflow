"""Write per-slide speaker notes from a markdown file into a PPTX.

Local habit this encodes: the deck that gets delivered carries its script in the
PowerPoint notes pane, not in a separate text file the presenter has to hunt for.

Notes markdown format (produced by scripts/new_deck.py):

    # 讲稿 / Speaker notes
    ## Slide 1 — 封面
    第一页要讲的话……
    ## Slide 2 — 章节页
    ……

Usage:
    python src/write_notes.py <deck.pptx> <notes.md> [--check]

Exit codes: 0 = notes written (or verified), 1 = mismatch/parse failure.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from pptx import Presentation


def parse_notes(path: Path) -> dict[int, str]:
    """Map 1-based slide number -> note body."""
    text = path.read_text(encoding="utf-8", errors="replace")
    blocks: dict[int, list[str]] = {}
    current: int | None = None
    for line in text.splitlines():
        match = re.match(r"^#{1,3}\s*(?:Slide|幻灯片|第)?\s*(\d+)\b", line.strip(), re.I)
        if match:
            current = int(match.group(1))
            blocks.setdefault(current, [])
            continue
        if line.startswith("# ") and current is None:
            continue
        if current is not None:
            blocks[current].append(line)
    return {n: "\n".join(lines).strip() for n, lines in blocks.items()}


def write(pptx: Path, notes: dict[int, str]) -> tuple[int, list[str]]:
    prs = Presentation(str(pptx))
    total = len(prs.slides)
    written = 0
    missing: list[str] = []
    for number, slide in enumerate(prs.slides, start=1):
        body = notes.get(number, "").strip()
        if not body:
            missing.append(f"slide {number}")
            continue
        slide.notes_slide.notes_text_frame.text = body
        written += 1
    prs.save(str(pptx))
    return written, [m for m in missing if int(m.split()[1]) <= total]


def verify(pptx: Path, notes: dict[int, str]) -> list[str]:
    """Reopen and read back every note we wrote."""
    prs = Presentation(str(pptx))
    problems: list[str] = []
    for number, slide in enumerate(prs.slides, start=1):
        expected = notes.get(number, "").strip()
        if not expected:
            continue
        try:
            actual = slide.notes_slide.notes_text_frame.text.strip()
        except AttributeError:
            actual = ""
        if expected.splitlines()[0] not in actual:
            problems.append(f"slide {number}: notes did not persist")
    return problems


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    pptx = Path(sys.argv[1])
    notes_path = Path(sys.argv[2])
    check_only = "--check" in sys.argv[3:]
    if not pptx.exists():
        print(f"FAIL missing deck: {pptx}")
        return 1
    if not notes_path.exists():
        print(f"FAIL missing notes: {notes_path}")
        return 1

    notes = parse_notes(notes_path)
    if not notes:
        print(f"FAIL no '## Slide N' headings found in {notes_path}")
        return 1

    if check_only:
        problems = verify(pptx, notes)
        print({"notes_slides": len(notes), "problems": problems,
               "status": "PASS" if not problems else "FAIL"}.__repr__())
        return 0 if not problems else 1

    written, missing = write(pptx, notes)
    problems = verify(pptx, notes)
    print(f"notes_written={written}")
    if missing:
        print("slides_without_notes=" + ",".join(missing))
    if problems:
        for problem in problems:
            print("FAIL " + problem)
        return 1
    print(f"notes_verified={written}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
