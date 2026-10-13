"""Stable public entry point for the template-first PPT workflow.

Usage:
    python src/run_workflow.py
    python src/run_workflow.py --config deck.config.json

The actual mutation is performed by build.py through PowerPoint COM so that
PowerPoint itself writes and reopens the resulting package.
"""
import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description="Build and verify a template-driven PPTX")
parser.add_argument("--config", default=str(ROOT / "deck.config.json"), help="workflow JSON config")
args = parser.parse_args()

cmd = [sys.executable, str(ROOT / "src" / "build.py"), "--config", args.config]
raise SystemExit(subprocess.call(cmd, cwd=str(ROOT)))
