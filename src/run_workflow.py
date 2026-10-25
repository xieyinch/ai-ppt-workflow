"""Stable public entry point for the complete agent PPT workflow.

Usage:
    python src/run_workflow.py
    python src/run_workflow.py --config deck.config.json

The mutation is still performed by build.py through PowerPoint COM. The public
entry additionally runs QA and creates a visual contact sheet; use
--html-preview for the optional browser preview route.
"""
import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description="Build and verify a template-driven PPTX")
parser.add_argument("--config", default=str(ROOT / "deck.config.json"), help="workflow JSON config")
parser.add_argument("--html-preview", action="store_true", help="export an optional browser preview")
parser.add_argument("--no-contact-sheet", action="store_true", help="skip the visual contact sheet")
args = parser.parse_args()

cmd = [sys.executable, str(ROOT / "src" / "agent_workflow.py"), "--config", args.config]
if args.html_preview:
    cmd.append("--html-preview")
if args.no_contact_sheet:
    cmd.append("--no-contact-sheet")
raise SystemExit(subprocess.call(cmd, cwd=str(ROOT)))
