#!/usr/bin/env python3
"""Download a published Google Sheet HTML page and render it as an 800x480 monochrome PNG."""

from __future__ import annotations

import html as html_module
import os
import re
import textwrap
import urllib.request
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SHEET_HTML_URL = os.environ.get(
    "SHEET_HTML_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQDVv0Nt9doPsH3SQgnFUHbvcs8i_s_iCw727ZP2sBX1Ty5RNjWRUCswVSxTa8qVLD5XAziLgN4IC60/pubhtml",
)
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "dist"))
WIDTH = 800
HEIGHT = 480
MARGIN = 18
MAX_COLUMNS = 5
MAX_DATA_ROWS = 10


@dataclass
class Cell:
    text_parts: list[str] = field(default_factory=list)
    checkbox: bool | None = None

    @property
    def text(self) -> str:
        text = " ".join(part.strip() for part in self.text_parts if part.strip())
        return re.sub(r"\s+", " ", html_module.unescape(text)).strip()


class PublishedSheetParser(HTMLParser):
    """Extract rows, text, and checkbox state from the first spreadsheet table."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[Cell]] = []
        self._table_depth = 0
        self._inside_target_table = False
        self._inside_row = False
        self._cell: Cell | None = None
        self._row: list[Cell] | None = None

    @staticmethod
    def _attrs(attrs: list[tuple[str, str | None]]) -> dict[str, str]:
        return {name.lower(): (value or "") for name, value in attrs}

    @staticmethod
    def _looks_like_sheet_table(attrs: dict[str, str]) -> bool:
        classes = attrs.get("class", "").lower().split()
        return "waffle" in classes or attrs.get("id", "").lower().startswith("sheets-viewport")

    @staticmethod
    def _checkbox_state(tag: str, attrs: dict[str, str]) -> bool | None:
        classes = attrs.get("class", "").lower()
        role = attrs.get("role", "").lower()
        input_type = attrs.get("type", "").lower()
        aria_checked = attrs.get("aria-checked", "").lower()
        checked_attr = "checked" in attrs

        is_checkbox = (
            input_type == "checkbox"
            or role == "checkbox"
            or "checkbox" in classes
            or "checkbox" in attrs.get("aria-label", "").lower()
        )
        if not is_checkbox:
            return None
        if aria_checked in {"true", "1", "yes"}:
            return True
        if aria_checked in {"false", "0", "no"}:
            return False
        if checked_attr:
            return True
        # Google sometimes marks checked boxes through a class name.
        if any(token in classes for token in ("checked", "is-checked", "waffle-checkbox-checked")):
            return True
        return False

    def handle_starttag(self, tag: str, attrs_list: list[tuple[str, str | None]]) -> None:
        attrs = self._attrs(attrs_list)

        if tag == "table":
            if not self._inside_target_table and self._looks_like_sheet_table(attrs):
                self._inside_target_table = True
                self._table_depth = 1
                return
            if self._inside_target_table:
                self._table_depth += 1

        if not self._inside_target_table:
            return

        if tag == "tr" and self._table_depth == 1:
            self._inside_row = True
            self._row = []
        elif tag in {"td", "th"} and self._inside_row:
            self._cell = Cell()
            state = self._checkbox_state(tag, attrs)
            if state is not None:
                self._cell.checkbox = state
        elif self._cell is not None:
            state = self._checkbox_state(tag, attrs)
            if state is not None:
                self._cell.checkbox = state
            if tag == "img":
                alt = attrs.get("alt", "").strip()
                title = attrs.get("title", "").strip()
                marker = f"{alt} {title}".lower()
                if "checkbox" in marker:
                    self._cell.checkbox = any(word in marker for word in ("checked", "true", "selected"))
                elif alt:
                    self._cell.text_parts.append(alt)
            elif tag == "br":
                self._cell.text_parts.append(" ")

    def handle_endtag(self, tag: str) -> None:
        if not self._inside_target_table:
            return

        if tag in {"td", "th"} and self._cell is not None and self._row is not None:
            self._row.append(self._cell)
            self._cell = None
        elif tag == "tr" and self._inside_row:
            if self._row and any(cell.text or cell.checkbox is not None for cell in self._row):
                self.rows.append(self._row)
            self._row = None
            self._inside_row = False
        elif tag == "table":
            self._table_depth -= 1
            if self._table_depth <= 0:
                self._inside_target_table = False
                self._table_depth = 0

    def handle_data(self, data: str) -> None:
        if self._cell is not None and data.strip():
            self._cell.text_parts.append(data)


def download_sheet_html(url: str) -> list[list[Cell]]:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "github-pages-sheet-renderer/2.0"},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        page = response.read().decode("utf-8", errors="replace")

    parser = PublishedSheetParser()
    parser.feed(page)
    if not parser.rows:
        raise RuntimeError("No spreadsheet table was found in the published Google Sheet HTML page.")
    return parser.rows


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


def normalize_rows(rows: list[list[Cell]]) -> list[list[Cell]]:
    column_count = min(MAX_COLUMNS, max(len(row) for row in rows))
    selected = rows[: MAX_DATA_ROWS + 1]
    return [(row + [Cell()] * column_count)[:column_count] for row in selected]


def draw_checkbox(
    draw: ImageDraw.ImageDraw,
    left: int,
    top: int,
    right: int,
    bottom: int,
    checked: bool,
) -> None:
    size = max(12, min(22, bottom - top - 12, right - left - 12))
    x = left + max(6, (right - left - size) // 2)
    y = top + max(5, (bottom - top - size) // 2)
    draw.rectangle((x, y, x + size, y + size), outline=0, width=2)
    if checked:
        draw.line(
            (
                x + round(size * 0.20),
                y + round(size * 0.53),
                x + round(size * 0.42),
                y + round(size * 0.75),
                x + round(size * 0.82),
                y + round(size * 0.25),
            ),
            fill=0,
            width=max(2, size // 7),
            joint="curve",
        )


def render_png(rows: list[list[Cell]], output_path: Path) -> None:
    rows = normalize_rows(rows)
    column_count = len(rows[0])
    row_count = len(rows)

    image = Image.new("1", (WIDTH, HEIGHT), 1)
    draw = ImageDraw.Draw(image)
    header_font = load_font(18, bold=True)
    body_font = load_font(17)

    # With no footer, the table uses all available height between the margins.
    table_top = MARGIN
    table_bottom = HEIGHT - MARGIN
    table_width = WIDTH - (2 * MARGIN)
    table_height = table_bottom - table_top
    col_width = table_width / column_count
    row_height = table_height / row_count

    draw.rectangle((MARGIN, table_top, WIDTH - MARGIN, table_bottom), outline=0, width=2)
    for col in range(1, column_count):
        x = round(MARGIN + col * col_width)
        draw.line((x, table_top, x, table_bottom), fill=0, width=1)
    for row in range(1, row_count):
        y = round(table_top + row * row_height)
        draw.line((MARGIN, y, WIDTH - MARGIN, y), fill=0, width=1)

    header_bottom = round(table_top + row_height)
    draw.rectangle((MARGIN + 1, table_top + 1, WIDTH - MARGIN - 1, header_bottom - 1), fill=0)

    for row_index, row in enumerate(rows):
        for col_index, cell in enumerate(row):
            left = round(MARGIN + col_index * col_width)
            right = round(MARGIN + (col_index + 1) * col_width)
            top = round(table_top + row_index * row_height)
            bottom = round(table_top + (row_index + 1) * row_height)

            if cell.checkbox is not None:
                # The HTML checkbox is redrawn as a crisp monochrome checkbox.
                if row_index == 0:
                    # A checkbox in a dark header needs inverted colors.
                    box_size = max(12, min(22, bottom - top - 12, right - left - 12))
                    x = left + max(6, (right - left - box_size) // 2)
                    y = top + max(5, (bottom - top - box_size) // 2)
                    draw.rectangle((x, y, x + box_size, y + box_size), outline=1, width=2)
                    if cell.checkbox:
                        draw.line(
                            (
                                x + round(box_size * 0.20), y + round(box_size * 0.53),
                                x + round(box_size * 0.42), y + round(box_size * 0.75),
                                x + round(box_size * 0.82), y + round(box_size * 0.25),
                            ),
                            fill=1,
                            width=max(2, box_size // 7),
                            joint="curve",
                        )
                else:
                    draw_checkbox(draw, left, top, right, bottom, cell.checkbox)
                continue

            available_width = max(1, right - left - 12)
            font = header_font if row_index == 0 else body_font
            fitted = fit_text(draw, cell.text, font, available_width)
            bbox = draw.textbbox((0, 0), fitted, font=font)
            text_height = bbox[3] - bbox[1]
            y = top + max(4, ((bottom - top) - text_height) // 2 - bbox[1])
            draw.text((left + 6, y), fitted, font=font, fill=1 if row_index == 0 else 0)

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
    rows = download_sheet_html(SHEET_HTML_URL)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    render_png(rows, OUTPUT_DIR / "sheet.png")
    write_index(OUTPUT_DIR)
    print(f"Created {OUTPUT_DIR / 'sheet.png'} ({WIDTH}x{HEIGHT}, 1-bit black and white)")


if __name__ == "__main__":
    main()
