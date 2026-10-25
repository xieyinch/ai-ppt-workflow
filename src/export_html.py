"""Export an optional browser preview from a rendered PPTX workflow.

This is deliberately a preview route, not a claim that HTML is a native PPTX.
It uses the already-rendered slide PNGs and extracted slide text, so review can
happen in a browser without changing the editable PowerPoint artifact.
"""
from __future__ import annotations

import argparse
import html
import re
import zipfile
from pathlib import Path



def slide_texts(pptx: Path) -> list[str]:
    results: list[str] = []
    with zipfile.ZipFile(pptx) as archive:
        names = sorted(
            [n for n in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
            key=lambda n: int(re.search(r"(\d+)", n).group(1)),
        )
        for name in names:
            # OOXML slide parts are local package data. Regex extraction avoids
            # enabling an XML parser on a file that may contain external entities.
            raw = archive.read(name).decode("utf-8", errors="replace")
            texts = [html.unescape(x) for x in re.findall(r"<a:t[^>]*>(.*?)</a:t>", raw)]
            results.append("\n".join(t for t in texts if t.strip()))
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx")
    parser.add_argument("render_dir")
    parser.add_argument("output")
    args = parser.parse_args()

    pptx = Path(args.pptx)
    render_dir = Path(args.render_dir)
    output = Path(args.output)
    texts = slide_texts(pptx)
    slides = []
    for index, text in enumerate(texts, 1):
        image = render_dir / f"slide-{index:02d}.png"
        if not image.exists():
            raise SystemExit(f"missing rendered image: {image}")
        paragraphs = "".join(
            f"<p>{html.escape(line)}</p>" for line in text.splitlines() if line.strip()
        )
        slides.append(
            f'<section class="slide"><img src="{html.escape(image.name)}" alt="slide {index}">'
            f'<div class="text">{paragraphs}</div></section>'
        )

    page = """<!doctype html>
<html lang="zh-CN"><meta charset="utf-8"><title>PPT workflow preview</title>
<style>
:root{color-scheme:dark;font-family:Inter,Arial,sans-serif;background:#101114;color:#f4f4f5}
body{margin:0;padding:32px;display:grid;gap:32px;justify-items:center}
.slide{width:min(1100px,94vw);background:#1b1d22;border:1px solid #343741;border-radius:18px;padding:16px;box-shadow:0 20px 70px #0008}
.slide img{width:100%;display:block;border-radius:10px;background:#000}
.text{display:none;color:#c8cad0;line-height:1.55;padding:12px 4px 0}.text p{margin:4px 0}
.slide:focus-within .text{display:block}
.slide:hover{border-color:#747bff}
</style>
<body>
<h1>Rendered PPTX preview</h1>
<p>Images are rendered from the editable PPTX. Text extraction is available on focus for review.</p>
""" + "\n".join(slides) + """
</body></html>"""
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(page, encoding="utf-8")
    for image in render_dir.glob("slide-*.png"):
        target = output.parent / image.name
        target.write_bytes(image.read_bytes())
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
