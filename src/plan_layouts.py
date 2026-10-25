"""Create a semantic slide plan before rendering.

This is the missing bridge between content and template filling: it records the
page job, visual type, candidate layouts, density rules, and selected backend.
The native builder still owns the actual PowerPoint mutation.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DESIGN = json.loads((ROOT / "src" / "design_system.json").read_text(encoding="utf-8"))

ROLE_RULES = [
    ("cover", ("title", "NE   N", "Your Project", "AI AGENT")),
    ("agenda", ("Table Of Contents", "Contents", "目录")),
    ("process", ("Our Process", "Process", "Timeline", "workflow", "机制")),
    ("comparison", ("Comparison", "Before", "After", "对比")),
    ("quote", ("quote", "statement", "inspiring quote", "价值")),
    ("closing", ("Thank you", "closing", "结束", "总结")),
]


def infer_role(slide: dict, index: int, total: int) -> tuple[str, str]:
    if slide.get("role"):
        return str(slide["role"]), "explicit"
    text = " ".join(str(k) for k in slide.get("map", {}).keys()).lower()
    if index == 0:
        return "cover", "position"
    if index == total - 1:
        return "closing", "position"
    for role, needles in ROLE_RULES:
        if any(str(n).lower() in text for n in needles):
            return role, "content-key"
    if slide.get("charts"):
        return "data", "chart"
    if slide.get("images"):
        return "image-text", "image"
    return "content", "fallback"


def visual_type(role: str) -> str:
    return {
        "cover": "hero-statement",
        "agenda": "agenda",
        "process": "process-4-phase",
        "comparison": "comparison-2col",
        "quote": "quote",
        "closing": "closing",
        "data": "data-chart",
        "image-text": "image-text",
        "content": "content-2col",
    }.get(role, "content-2col")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    parser.add_argument("--out")
    args = parser.parse_args()
    cfg_path = Path(args.config).resolve()
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    content_path = ROOT / cfg["content"]
    content = json.loads(content_path.read_text(encoding="utf-8"))
    design_cfg = cfg.get("design", {})
    preset_name = design_cfg.get("preset", "corporate")
    preset = DESIGN["presets"].get(preset_name, DESIGN["presets"]["corporate"])
    catalog = {}
    catalog_path = design_cfg.get("template_catalog")
    if catalog_path:
        cp = ROOT / catalog_path
        if cp.exists():
            catalog = json.loads(cp.read_text(encoding="utf-8"))
    role_sources = {k: list(v) for k, v in catalog.get("role_sources", {}).items()}
    slides = content.get("slides", [])
    planned = []
    for i, slide in enumerate(slides):
        role, reason = infer_role(slide, i, len(slides))
        candidates = list(preset.get("layouts", []))
        source_candidates = role_sources.get(role, [])
        template_src = source_candidates.pop(0) if source_candidates else slide.get("src")
        role_sources[role] = source_candidates
        planned.append({
            "index": i + 1,
            "source_slide": slide.get("src"),
            "template_src": template_src,
            "role": role,
            "inference": reason,
            "visual_type": visual_type(role),
            "layout_candidates": candidates[:3],
            "page_job": f"{role}: communicate one primary claim",
            "density": preset.get("density"),
            "rules": preset.get("rules", []),
            "explicit_role": bool(slide.get("role")),
        })
    out = Path(args.out).resolve() if args.out else ROOT / cfg["output"].replace(".pptx", "-layout-plan.json")
    plan = {
        "version": 1,
        "topic": cfg.get("topic"),
        "preset": preset_name,
        "backend": "native" if design_cfg.get("editable", True) else "html",
        "source_content": str(content_path),
        "open_source_reference": "template profiling → semantic visual type → layout plan",
        "slides": planned,
        "warnings": ["Add explicit role fields to content JSON to remove heuristic inference."] if any(not x["explicit_role"] for x in planned) else [],
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"plan": str(out), "slides": len(planned), "preset": preset_name}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
