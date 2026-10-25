"""Create a review contact sheet from rendered slide PNGs."""
from __future__ import annotations

import argparse
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("render_dir")
    parser.add_argument("output")
    parser.add_argument("--columns", type=int, default=2)
    args = parser.parse_args()

    try:
        from PIL import Image, ImageDraw
    except ImportError as exc:
        raise SystemExit("Pillow is required for contact sheets: " + str(exc))

    files = sorted(Path(args.render_dir).glob("slide-*.png"))
    if not files:
        raise SystemExit(f"no rendered slides found in {args.render_dir}")

    thumb_w, thumb_h = 320, 180
    cell_w, cell_h = 340, 215
    rows = (len(files) + args.columns - 1) // args.columns
    sheet = Image.new("RGB", (args.columns * cell_w, rows * cell_h), "#d8d8d8")
    for i, path in enumerate(files):
        image = Image.open(path).convert("RGB")
        image.thumbnail((thumb_w, thumb_h))
        cell = Image.new("RGB", (cell_w, cell_h), "white")
        cell.paste(image, ((cell_w - image.width) // 2, 18))
        ImageDraw.Draw(cell).text((10, 194), path.stem, fill="black")
        sheet.paste(cell, ((i % args.columns) * cell_w, (i // args.columns) * cell_h))

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
