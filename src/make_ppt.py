"""Unified multi-backend PPT agent dispatcher.

The dispatcher selects one artifact route; it does not pretend HTML, native
PPTX, and image-PPTX have the same editability. Each backend reports its own
artifact contract in the manifest.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = ROOT.parent
ROUTES = json.loads((PROJECT_ROOT / "workflow_routes.json").read_text(encoding="utf-8"))


def run(cmd: list[str], cwd: Path) -> int:
    print("$", " ".join(str(x) for x in cmd))
    return subprocess.run(cmd, cwd=cwd).returncode


def main() -> int:
    p = argparse.ArgumentParser(description="Route a PPT request to the correct backend")
    p.add_argument("--route", choices=["auto", "native", "guizang", "huashu", "image"], default="auto")
    p.add_argument("--config", default=str(ROOT / "deck.config.json"))
    p.add_argument("--html-preview", action="store_true")
    p.add_argument("--slides-dir", help="directory of constrained HTML slides for Huashu export")
    p.add_argument("--out", help="Huashu/image route output path")
    args = p.parse_args()

    route = args.route
    if route == "auto":
        cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
        design = cfg.get("design", {})
        route = design.get("route") or ("guizang" if design.get("preset") in {"editorial", "swiss"} else "native")

    if route == "native":
        cmd = [sys.executable, str(ROOT / "agent_workflow.py"), "--config", str(Path(args.config).resolve())]
        if args.html_preview:
            cmd.append("--html-preview")
        code = run(cmd, PROJECT_ROOT)
        artifact = "native PPTX"
    elif route == "guizang":
        script = Path(os.environ.get("GUIZANG_PPT_SCRIPT", ""))
        if not script:
            raise SystemExit("guizang route requires GUIZANG_PPT_SCRIPT pointing to the backend script")
        code = run([sys.executable, str(script)], script.parent.parent)
        artifact = "Guizang HTML deck"
    elif route == "huashu":
        if not args.slides_dir or not args.out:
            p.error("--route huashu requires --slides-dir and --out")
        script = Path(os.environ.get("HUASHU_PPT_SCRIPT", ""))
        if not script:
            raise SystemExit("huashu route requires HUASHU_PPT_SCRIPT pointing to the backend script")
        code = run(["node", str(script), "--slides", str(Path(args.slides_dir).resolve()), "--out", str(Path(args.out).resolve())], script.parent.parent)
        artifact = "Huashu-converted PPTX (editability depends on HTML constraints)"
    else:
        if not args.out:
            p.error("--route image requires --out")
        script = Path(args.out).resolve().parent / "export_fallback.mjs"
        code = run(["node", str(script)], script.parent)
        artifact = "image-based PPTX"

    manifest = PROJECT_ROOT / "output" / "last-route.json"
    manifest.parent.mkdir(exist_ok=True)
    manifest.write_text(json.dumps({
        "route": route,
        "artifact": artifact,
        "status": "PASS" if code == 0 else "FAIL",
        "created_at": datetime.now(timezone.utc).isoformat(),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"route={route}")
    print(f"artifact={artifact}")
    print(f"manifest={manifest}")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
