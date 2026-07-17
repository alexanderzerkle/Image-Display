#!/usr/bin/env python3
"""Download a public Google Sheet CSV and render it as an 800x480 monochrome PNG."""

from __future__ import annotations

import csv
import io
import os
import textwrap
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SHEET_CSV_URL = os.environ.get(
    "SHEET_CSV_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQDVv0Nt9doPsH3SQgnFUHbvcs8i_s_iCw727ZP2sBX1Ty5RNjWRUCswVSxTa8qVLD5XAziLgN4IC60/pub?output=csv",
)
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "dist"))
WIDTH = 800
HEIGHT = 480
MARGIN = 18
MAX_COLUMNS = 5
MAX_DATA_ROWS = 9


def download_csv(url: str) -> list[list[str]]:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "github-pages-sheet-renderer/1.0"},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        raw = response.read()

    text = raw.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(text)))
    rows = [[cell.strip() for cell in row] for row in rows]
    rows = [row for row in rows if any(cell for cell in row)]
    if not rows:
        raise RuntimeError("The spreadsheet CSV did not contain any non-empty rows.")
    return rows


def load_font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    candidates = (
        [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        ]
        if bold
        else [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ]
    )
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def fit_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> str:
    text = text.replace("\n", " ").strip()
    if not text:
        return ""
    if draw.textlength(text, font=font) <= width:
        return text

    ellipsis = "..."
    shortened = text
    while shortened and draw.textlength(shortened + ellipsis, font=font) > width:
        shortened = shortened[:-1]
    return shortened.rstrip() + ellipsis if shortened else ellipsis


def normalize_rows(rows: list[list[str]]) -> list[list[str]]:
    column_count = min(MAX_COLUMNS, max(len(row) for row in rows))
    selected = rows[: MAX_DATA_ROWS + 1]
    return [(row + [""] * column_count)[:column_count] for row in selected]


def render_png(rows: list[list[str]], output_path: Path) -> None:
    rows = normalize_rows(rows)
    column_count = len(rows[0])
    row_count = len(rows)

    image = Image.new("1", (WIDTH, HEIGHT), 1)  # 1-bit image: white background
    draw = ImageDraw.Draw(image)
    header_font = load_font(18, bold=True)
    body_font = load_font(17)
    footer_font = load_font(12)

    table_top = MARGIN
    table_bottom = HEIGHT - 34
    table_width = WIDTH - (2 * MARGIN)
    table_height = table_bottom - table_top
    col_width = table_width / column_count
    row_height = table_height / row_count

    # Border and grid.
    draw.rectangle((MARGIN, table_top, WIDTH - MARGIN, table_bottom), outline=0, width=2)
    for col in range(1, column_count):
        x = round(MARGIN + col * col_width)
        draw.line((x, table_top, x, table_bottom), fill=0, width=1)
    for row in range(1, row_count):
        y = round(table_top + row * row_height)
        draw.line((MARGIN, y, WIDTH - MARGIN, y), fill=0, width=1)

    # Header cells are black with white text for strong monochrome contrast.
    header_bottom = round(table_top + row_height)
    draw.rectangle((MARGIN + 1, table_top + 1, WIDTH - MARGIN - 1, header_bottom - 1), fill=0)

    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            left = round(MARGIN + col_index * col_width)
            right = round(MARGIN + (col_index + 1) * col_width)
            top = round(table_top + row_index * row_height)
            bottom = round(table_top + (row_index + 1) * row_height)
            available_width = max(1, right - left - 12)
            font = header_font if row_index == 0 else body_font
            fitted = fit_text(draw, value, font, available_width)
            bbox = draw.textbbox((0, 0), fitted, font=font)
            text_height = bbox[3] - bbox[1]
            y = top + max(4, ((bottom - top) - text_height) // 2 - bbox[1])
            draw.text((left + 6, y), fitted, font=font, fill=1 if row_index == 0 else 0)

    footer = "Generated automatically from the public Google Sheet"
    draw.text((MARGIN, HEIGHT - 25), footer, font=footer_font, fill=0)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path, format="PNG", optimize=True)


def write_index(output_dir: Path) -> None:
    html = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Google Sheet PNG</title>
  <style>
    html, body { min-height: 100%; }
    body { margin: 0; display: grid; place-items: center; background: #eee; font-family: sans-serif; }
    main { max-width: 840px; padding: 20px; text-align: center; }
    img { display: block; width: min(800px, 100%); height: auto; border: 1px solid #000; image-rendering: auto; }
    p { margin-bottom: 0; }
  </style>
</head>
<body>
  <main>
    <img src="sheet.png" width="800" height="480" alt="Table generated from a public Google Sheet">
    <p><a href="sheet.png">Open the PNG directly</a></p>
  </main>
</body>
</html>
"""
    (output_dir / "index.html").write_text(textwrap.dedent(html), encoding="utf-8")
    (output_dir / ".nojekyll").write_text("", encoding="utf-8")


def main() -> None:
    rows = download_csv(SHEET_CSV_URL)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    render_png(rows, OUTPUT_DIR / "sheet.png")
    write_index(OUTPUT_DIR)
    print(f"Created {OUTPUT_DIR / 'sheet.png'} ({WIDTH}x{HEIGHT}, 1-bit black and white)")


if __name__ == "__main__":
    main()
