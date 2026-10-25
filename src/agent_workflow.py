"""Unified agent workflow entry point.

Routes one request through:
1) native PowerPoint template/native-object build
2) structural and content QA
3) rendered slide contact sheet for visual review
4) optional browser preview HTML (never confused with native PPTX)

The existing build.py remains the authoritative PowerPoint/COM mutation engine.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run(command: list[str]) -> int:
    result = subprocess.run(command, cwd=ROOT)
    return result.returncode


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the complete agent PPT workflow")
    parser.add_argument("--config", default=str(ROOT / "deck.config.json"))
    parser.add_argument("--html-preview", action="store_true", help="also export a browser preview")
    parser.add_argument("--no-contact-sheet", action="store_true")
    args = parser.parse_args()

    config = Path(args.config).resolve()
    import json
    cfg = json.loads(config.read_text(encoding="utf-8"))
    output = ROOT / cfg["output"]
    render_dir = ROOT / cfg["render_dir"]
    contact = output.with_name(output.stem + "-contact.png")

    # Plan semantic page jobs before touching PowerPoint. This is the bridge
    # between content, design_system.json, and the template-fill executor.
    plan_path = output.with_name(output.stem + "-layout-plan.json")
    if run([sys.executable, str(ROOT / "src" / "plan_layouts.py"),
            "--config", str(config), "--out", str(plan_path)]) != 0:
        return 1

    if run([sys.executable, str(ROOT / "src" / "build.py"), "--config", str(config)]) != 0:
        return 1
    if run([sys.executable, str(ROOT / "src" / "qa.py"), str(config)]) != 0:
        return 1
    if not args.no_contact_sheet:
        if run([sys.executable, str(ROOT / "src" / "make_contact_sheet.py"), str(render_dir), str(contact)]) != 0:
            return 1
    if args.html_preview:
        html_path = output.with_suffix(".html")
        if run([sys.executable, str(ROOT / "src" / "export_html.py"), str(output), str(render_dir), str(html_path)]) != 0:
            return 1
        print(f"html_preview={html_path}")
    print(f"pptx={output}")
    print(f"contact_sheet={contact}")
    print(f"layout_plan={plan_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
