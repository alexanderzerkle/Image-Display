#!/usr/bin/env python3
"""Download a published Google Sheet as XLSX and render its first sheet to PNG."""

from __future__ import annotations

import io
import os
import textwrap
import urllib.request
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.styles import Border, Color
from openpyxl.utils import get_column_letter
from PIL import Image, ImageColor, ImageDraw, ImageFont

SHEET_XLSX_URL = os.environ.get(
    "SHEET_XLSX_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQDVv0Nt9doPsH3SQgnFUHbvcs8i_s_iCw727ZP2sBX1Ty5RNjWRUCswVSxTa8qVLD5XAziLgN4IC60/pub?output=xlsx",
)
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "dist"))
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "").strip()
WIDTH = 800
HEIGHT = 480
MARGIN = 12

# Safety limits prevent accidentally rendering a workbook with an enormous styled range.
MAX_ROWS = int(os.environ.get("MAX_ROWS", "40"))
MAX_COLUMNS = int(os.environ.get("MAX_COLUMNS", "12"))


# Approximate Excel indexed colors. Unrecognized theme colors fall back safely.
INDEXED_COLORS = {
    0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00",
    4: "#0000FF", 5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF",
    8: "#000000", 9: "#FFFFFF", 10: "#FF0000", 11: "#00FF00",
    12: "#0000FF", 13: "#FFFF00", 14: "#FF00FF", 15: "#00FFFF",
    16: "#800000", 17: "#008000", 18: "#000080", 19: "#808000",
    20: "#800080", 21: "#008080", 22: "#C0C0C0", 23: "#808080",
}


def download_xlsx(url: str) -> bytes:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "github-pages-sheet-renderer/2.0"},
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        data = response.read()
        content_type = response.headers.get("Content-Type", "")

    if len(data) < 4 or data[:2] != b"PK":
        preview = data[:200].decode("utf-8", errors="replace")
        raise RuntimeError(
            "The XLSX URL did not return an Excel workbook. "
            f"Content-Type was {content_type!r}; response began with {preview!r}."
        )
    return data


def load_font(size: int, bold: bool = False, italic: bool = False) -> ImageFont.ImageFont:
    if bold and italic:
        names = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-BoldItalic.ttf",
        ]
    elif bold:
        names = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        ]
    elif italic:
        names = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf",
        ]
    else:
        names = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        ]
    for name in names:
        if Path(name).exists():
            return ImageFont.truetype(name, size=max(7, size))
    return ImageFont.load_default()


def color_to_hex(color: Color | None, default: str) -> str:
    if color is None:
        return default
    try:
        if color.type == "rgb" and color.rgb:
            value = color.rgb[-6:]
            return f"#{value}"
        if color.type == "indexed" and color.indexed is not None:
            return INDEXED_COLORS.get(int(color.indexed), default)
    except (TypeError, ValueError):
        pass
    return default


def safe_color(value: str, default: str) -> str:
    try:
        ImageColor.getrgb(value)
        return value
    except ValueError:
        return default


def fill_color(cell: Cell) -> str:
    fill = cell.fill
    if fill and fill.fill_type in {"solid", "gray125"}:
        return safe_color(color_to_hex(fill.fgColor, "#FFFFFF"), "#FFFFFF")
    return "#FFFFFF"


def font_color(cell: Cell) -> str:
    return safe_color(color_to_hex(cell.font.color, "#000000"), "#000000")


def border_width(style: str | None) -> int:
    if style in {"medium", "mediumDashed", "mediumDashDot", "mediumDashDotDot"}:
        return 2
    if style in {"thick", "double"}:
        return 3
    return 1 if style else 0


def meaningful_cells(ws) -> tuple[int, int]:
    max_row = 1
    max_col = 1
    for row in ws.iter_rows(
        min_row=1,
        max_row=min(ws.max_row, MAX_ROWS),
        min_col=1,
        max_col=min(ws.max_column, MAX_COLUMNS),
    ):
        for cell in row:
            if cell.value not in (None, "") or cell.has_style:
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
    return max_row, max_col


def checkbox_cells(ws) -> set[str]:
    """Return cells with TRUE/FALSE-style validation, when preserved in the XLSX."""
    result: set[str] = set()
    validations = getattr(ws.data_validations, "dataValidation", [])
    for validation in validations:
        formula = str(validation.formula1 or "").upper().replace(" ", "")
        if "TRUE" not in formula or "FALSE" not in formula:
            continue
        for cell_range in validation.ranges.ranges:
            for row in ws.iter_rows(
                min_row=cell_range.min_row,
                max_row=cell_range.max_row,
                min_col=cell_range.min_col,
                max_col=cell_range.max_col,
            ):
                result.update(cell.coordinate for cell in row)
    return result


def excel_column_width(ws, column: int) -> float:
    dimension = ws.column_dimensions[get_column_letter(column)]
    return float(dimension.width or 8.43)


def excel_row_height(ws, row: int) -> float:
    dimension = ws.row_dimensions[row]
    return float(dimension.height or 15.0)


def scaled_boundaries(values: Iterable[float], start: int, end: int) -> list[int]:
    values = list(values)
    total = sum(values) or 1.0
    available = end - start
    boundaries = [start]
    running = 0.0
    for value in values:
        running += value
        boundaries.append(round(start + available * running / total))
    boundaries[-1] = end
    return boundaries


def merged_anchor_map(ws, max_row: int, max_col: int) -> tuple[dict[str, str], dict[str, tuple[int, int, int, int]]]:
    covered: dict[str, str] = {}
    anchors: dict[str, tuple[int, int, int, int]] = {}
    for merged in ws.merged_cells.ranges:
        if merged.min_row > max_row or merged.min_col > max_col:
            continue
        anchor = f"{get_column_letter(merged.min_col)}{merged.min_row}"
        bounds = (
            merged.min_row,
            min(merged.max_row, max_row),
            merged.min_col,
            min(merged.max_col, max_col),
        )
        anchors[anchor] = bounds
        for row in range(bounds[0], bounds[1] + 1):
            for col in range(bounds[2], bounds[3] + 1):
                covered[f"{get_column_letter(col)}{row}"] = anchor
    return covered, anchors


def fit_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> str:
    text = text.replace("\r", "").strip()
    if not text:
        return ""
    if draw.textlength(text, font=font) <= width:
        return text
    ellipsis = "..."
    shortened = text
    while shortened and draw.textlength(shortened + ellipsis, font=font) > width:
        shortened = shortened[:-1]
    return shortened.rstrip() + ellipsis if shortened else ellipsis


def draw_checkbox(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], checked: bool, color: str) -> None:
    left, top, right, bottom = box
    size = max(8, min(20, right - left - 8, bottom - top - 8))
    x = left + (right - left - size) // 2
    y = top + (bottom - top - size) // 2
    draw.rectangle((x, y, x + size, y + size), outline=color, width=max(1, size // 8))
    if checked:
        stroke = max(2, size // 7)
        draw.line(
            (
                x + round(size * 0.20), y + round(size * 0.53),
                x + round(size * 0.43), y + round(size * 0.76),
                x + round(size * 0.82), y + round(size * 0.25),
            ),
            fill=color,
            width=stroke,
            joint="curve",
        )


def cell_text(cell: Cell) -> str:
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def draw_cell_borders(
    draw: ImageDraw.ImageDraw,
    border: Border,
    left: int,
    top: int,
    right: int,
    bottom: int,
) -> None:
    sides = [
        (border.left, (left, top, left, bottom)),
        (border.top, (left, top, right, top)),
        (border.right, (right, top, right, bottom)),
        (border.bottom, (left, bottom, right, bottom)),
    ]
    for side, line in sides:
        width = border_width(side.style)
        if width:
            draw.line(line, fill=color_to_hex(side.color, "#000000"), width=width)


def render_png(workbook_bytes: bytes, output_path: Path) -> None:
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    if WORKSHEET_NAME:
        if WORKSHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Worksheet {WORKSHEET_NAME!r} was not found. Available sheets: {workbook.sheetnames}"
            )
        ws = workbook[WORKSHEET_NAME]
    else:
        ws = workbook.active

    max_row, max_col = meaningful_cells(ws)
    checkbox_coordinates = checkbox_cells(ws)
    covered, merged_anchors = merged_anchor_map(ws, max_row, max_col)

    col_bounds = scaled_boundaries(
        [excel_column_width(ws, col) for col in range(1, max_col + 1)],
        MARGIN,
        WIDTH - MARGIN,
    )
    row_bounds = scaled_boundaries(
        [excel_row_height(ws, row) for row in range(1, max_row + 1)],
        MARGIN,
        HEIGHT - MARGIN,
    )

    image = Image.new("RGB", (WIDTH, HEIGHT), "white")
    draw = ImageDraw.Draw(image)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            coordinate = f"{get_column_letter(col)}{row}"
            anchor = covered.get(coordinate, coordinate)
            if coordinate != anchor:
                continue

            cell = ws[anchor]
            if anchor in merged_anchors:
                min_row, merged_max_row, min_col, merged_max_col = merged_anchors[anchor]
            else:
                min_row = merged_max_row = row
                min_col = merged_max_col = col

            left = col_bounds[min_col - 1]
            right = col_bounds[merged_max_col]
            top = row_bounds[min_row - 1]
            bottom = row_bounds[merged_max_row]

            draw.rectangle((left, top, right, bottom), fill=fill_color(cell))
            draw_cell_borders(draw, cell.border, left, top, right, bottom)

            value = cell.value
            is_checkbox = coordinate in checkbox_coordinates or isinstance(value, bool)
            if is_checkbox and (isinstance(value, bool) or str(value).upper() in {"TRUE", "FALSE"}):
                checked = value is True or str(value).upper() == "TRUE"
                draw_checkbox(draw, (left, top, right, bottom), checked, font_color(cell))
                continue

            text = cell_text(cell)
            if not text:
                continue

            # Convert Excel points to a practical screen size, then scale down when needed.
            requested_size = int(round(float(cell.font.sz or 11) * 1.15))
            cell_height = max(1, bottom - top)
            font_size = min(requested_size, max(7, cell_height - 6))
            font = load_font(font_size, bool(cell.font.bold), bool(cell.font.italic))
            padding = 5
            available_width = max(1, right - left - 2 * padding)
            fitted = fit_text(draw, text, font, available_width)
            bbox = draw.textbbox((0, 0), fitted, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            horizontal = cell.alignment.horizontal or "general"
            vertical = cell.alignment.vertical or "bottom"
            if horizontal in {"center", "centerContinuous"}:
                x = left + (right - left - text_width) // 2
            elif horizontal == "right" or (horizontal == "general" and isinstance(value, (int, float))):
                x = right - padding - text_width
            else:
                x = left + padding

            if vertical == "top":
                y = top + padding - bbox[1]
            elif vertical in {"center", "distributed", "justify"}:
                y = top + (bottom - top - text_height) // 2 - bbox[1]
            else:
                y = bottom - padding - text_height - bbox[1]

            draw.text((x, y), fitted, font=font, fill=font_color(cell))

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
    img { display: block; width: min(800px, 100%); height: auto; border: 1px solid #000; }
    p { margin-bottom: 0; }
  </style>
</head>
<body>
  <main>
    <img src="sheet.png" width="800" height="480" alt="Spreadsheet rendered as an image">
    <p><a href="sheet.png">Open the PNG directly</a></p>
  </main>
</body>
</html>
"""
    (output_dir / "index.html").write_text(textwrap.dedent(html), encoding="utf-8")
    (output_dir / ".nojekyll").write_text("", encoding="utf-8")


def main() -> None:
    workbook_bytes = download_xlsx(SHEET_XLSX_URL)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    render_png(workbook_bytes, OUTPUT_DIR / "sheet.png")
    write_index(OUTPUT_DIR)
    print(f"Created {OUTPUT_DIR / 'sheet.png'} ({WIDTH}x{HEIGHT})")


if __name__ == "__main__":
    main()
