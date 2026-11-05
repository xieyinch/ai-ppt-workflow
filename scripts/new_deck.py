"""Scaffold a new deck from a bundled template.

Usage:
    python scripts/new_deck.py mydeck
    python scripts/new_deck.py mydeck --template inner-chapter-navy
    python scripts/new_deck.py mydeck --template inner-chapter --blank
    python scripts/new_deck.py --list

Creates decks/<name>/ containing:
    <template>.pptx   copy of the bundled deck to edit
    content.json      one entry per slide, pre-filled with the current text
    deck.config.json  ready to pass to src/run_workflow.py
    notes.md          speaker notes stub
    sources.md        source list stub

Then edit content.json (map: original text -> new text) and run:
    python src/run_workflow.py --config decks/<name>/deck.config.json
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
TEMPLATES = ROOT / "templates"
DECKS = ROOT / "decks"

SKIP_TYPES = {"PICTURE", "SLIDE_NUMBER", "DATE", "FOOTER", "MEDIA_CLIP"}


def list_templates() -> int:
    manifest_path = TEMPLATES / "manifest.json"
    if not manifest_path.exists():
        print("templates/manifest.json is missing")
        return 1
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    print(f"{'slug':26} {'name':34} {'layouts':>7}  license")
    for item in manifest["templates"]:
        print(f"{item['slug']:26} {item['name'][:34]:34} {item.get('layouts', '-'):>7}  {item.get('license', '-')}")
    return 0


def source_deck(slug: str, blank: bool) -> Path:
    """Resolve the .pptx to start from.

    A slug with a starter deck (starter.pptx) hands over a deck that already has
    content in every layout; --blank picks the empty 59-layout master instead.
    """
    base = TEMPLATES / slug
    if not base.is_dir():
        raise SystemExit(f"unknown template: {slug} (run --list)")
    starter = base / "starter.pptx"
    master = base / "template.pptx"
    if starter.exists() and not blank:
        return starter
    if master.exists():
        return master
    raise SystemExit(f"no .pptx found under templates/{slug}")


def slide_text(prs: Presentation) -> tuple[int, dict[int, dict[str, str]]]:
    """Return slide count and per-slide {placeholder_text: placeholder_text}."""
    slides: dict[int, dict[str, str]] = {}
    for number, slide in enumerate(prs.slides, start=1):
        texts: dict[str, str] = {}
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            try:
                if str(shape.placeholder_format.type) .startswith(tuple(SKIP_TYPES)):
                    continue
                if shape.placeholder_format.type is not None and shape.placeholder_format.type.name in SKIP_TYPES:
                    continue
            except Exception:
                pass
            body = "\n".join(p.text for p in shape.text_frame.paragraphs).strip()
            if not body:
                continue
            key = body.split("\n")[0][:80]
            if key in texts:
                continue
            texts[key] = body
        if texts:
            slides[number] = texts
    return len(prs.slides._sldIdLst), slides


def notes_text(texts: dict[int, dict[str, str]]) -> str:
    """Speaker-notes stub: one heading per slide, seeded with the current text."""
    lines = ["# 讲稿 / Speaker notes", "", "每页一段：这页要讲什么、证据是什么、结论是什么。", ""]
    for number in sorted(texts):
        head = next(iter(texts[number].values())).split("\n")[0][:40]
        lines.append(f"## Slide {number} — {head}")
        lines.append("")
        lines.append("TODO: 讲稿要点")
        lines.append("")
    if len(texts) < 1:
        lines += ["## Slide 1", "", "TODO: 讲稿要点", ""]
    return "\n".join(lines)


def sources_text(slug: str) -> str:
    """Source list stub, seeded with the real template attribution URL."""
    manifest_path = TEMPLATES / "manifest.json"
    origin = ""
    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        for item in manifest["templates"]:
            if item["slug"] == slug:
                origin = item.get("origin", "")
                break
    lines = ["# 来源 / Sources", ""]
    url = re.search(r"https?://\S+", origin)
    if url:
        lines.append(f"模板来源（{slug}）：{url.group(0).rstrip('),')}")
        lines.append("")
    lines += ["## 内容来源（待补）", "", "TODO: 每条结论/数据都要有可查链接", ""]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Scaffold a deck from a bundled template")
    parser.add_argument("name", nargs="?", help="deck folder name")
    parser.add_argument("--template", default="starter-cn", help="bundled template slug")
    parser.add_argument("--blank", action="store_true", help="start from the empty master/layout set")
    parser.add_argument("--list", action="store_true", help="list bundled templates")
    args = parser.parse_args()

    if args.list or not args.name:
        return list_templates()

    asset = source_deck(args.template, args.blank)
    target = DECKS / args.name
    if target.exists():
        raise SystemExit(f"{target} already exists")

    (target).mkdir(parents=True)
    deck_path = target / asset.name
    shutil.copy2(asset, deck_path)

    prs = Presentation(deck_path)
    count, texts = slide_text(prs)

    content = {"slides": [{"src": n, "map": texts[n]} for n in sorted(texts)]}
    if not content["slides"]:
        content["slides"] = [{"src": 1, "map": {}}]
    content_path = target / "content.json"
    content_path.write_text(json.dumps(content, ensure_ascii=False, indent=2), encoding="utf-8")

    rel = lambda p: p.relative_to(ROOT).as_posix()
    config = {
        "topic": args.name,
        "template": rel(deck_path),
        "content": rel(content_path),
        "output": rel(target / "output" / f"{args.name}.pptx"),
        "render_dir": rel(target / "render"),
        "notes": rel(target / "notes.md"),
        "sources": rel(target / "sources.md"),
        "quality": {"min_slides": 1, "min_text_chars": 20, "required_terms": []},
    }
    (target / "deck.config.json").write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    (target / "notes.md").write_text(notes_text(texts), encoding="utf-8")
    (target / "sources.md").write_text(sources_text(args.template), encoding="utf-8")

    print(f"deck:     {target}")
    print(f"template: {asset.relative_to(ROOT).as_posix()}  ({count} slides)")
    print(f"editable text slots: {sum(len(v) for v in texts.values())} across {len(texts)} slides")
    print()
    print("next:")
    print(f"  1. edit {rel(content_path)}  (map: original text -> your text)")
    print(f"  2. python src/run_workflow.py --config {rel(target / 'deck.config.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
