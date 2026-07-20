#!/usr/bin/env python3
"""Render the first 40 x 24 cells of a published Google Sheet XLSX as an 800 x 480 PNG.

The renderer treats the workbook as a page layout: every worksheet cell occupies an
exact 20 x 20 pixel square, so spreadsheet coordinates map directly to image pixels.
"""

from __future__ import annotations

import io
import os
import textwrap
import urllib.request
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Border, Color, PatternFill
from PIL import Image, ImageDraw, ImageFont

SHEET_XLSX_URL = os.environ.get(
    "SHEET_XLSX_URL",
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vQDVv0Nt9doPsH3SQgnFUHbvcs8i_s_iCw727ZP2sBX1Ty5RNjWRUCswVSxTa8qVLD5XAziLgN4IC60/"
    "pub?output=xlsx",
)
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "dist"))
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "").strip()

ROWS = 24
COLUMNS = 40
CELL_WIDTH = 20
CELL_HEIGHT = 20
WIDTH = COLUMNS * CELL_WIDTH
HEIGHT = ROWS * CELL_HEIGHT

MIN_FONT_SIZE = 6
MAX_FONT_SIZE = 32
DEFAULT_FONT_SIZE = 11
TEXT_PADDING_X = 3


FONT_ROOTS = (
    Path("/usr/share/fonts/truetype/liberation2"),
    Path("/usr/share/fonts/truetype/dejavu"),
    Path("/usr/share/fonts/truetype/crosextra"),
)

FONT_FAMILIES = {
    "arial": ("LiberationSans", "DejaVuSans"),
    "helvetica": ("LiberationSans", "DejaVuSans"),
    "calibri": ("Carlito", "LiberationSans", "DejaVuSans"),
    "aptos": ("Carlito", "LiberationSans", "DejaVuSans"),
    "times new roman": ("LiberationSerif", "DejaVuSerif"),
    "times": ("LiberationSerif", "DejaVuSerif"),
    "georgia": ("LiberationSerif", "DejaVuSerif"),
    "courier new": ("LiberationMono", "DejaVuSansMono"),
    "consolas": ("LiberationMono", "DejaVuSansMono"),
}


def download_xlsx(url: str) -> bytes:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "github-pages-sheet-renderer/4.1-no-ellipsis"},
    )
    with urllib.request.urlopen(request, timeout=45) as response:
        data = response.read()
        content_type = response.headers.get("Content-Type", "")

    if not data.startswith(b"PK"):
        preview = data[:200].decode("utf-8", errors="replace")
        raise RuntimeError(
            "The spreadsheet URL did not return an XLSX file. "
            f"Content-Type was {content_type!r}. Response began with {preview!r}."
        )
    return data


def _font_filename(base: str, bold: bool, italic: bool) -> tuple[str, ...]:
    if base == "Carlito":
        suffix = "BoldItalic" if bold and italic else "Bold" if bold else "Italic" if italic else "Regular"
        return (f"Carlito-{suffix}.ttf",)
    if base.startswith("Liberation"):
        suffix = "-BoldItalic" if bold and italic else "-Bold" if bold else "-Italic" if italic else "-Regular"
        return (f"{base}{suffix}.ttf",)
    if base.startswith("DejaVu"):
        suffix = "-BoldOblique" if bold and italic else "-Bold" if bold else "-Oblique" if italic else ""
        return (f"{base}{suffix}.ttf",)
    return ()


def load_font(family: str | None, size: float, bold: bool, italic: bool) -> ImageFont.ImageFont:
    pixel_size = max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, int(round(size))))
    requested = (family or "Arial").strip().lower()
    bases = FONT_FAMILIES.get(requested, ("LiberationSans", "DejaVuSans"))

    for base in bases:
        for filename in _font_filename(base, bold, italic):
            for root in FONT_ROOTS:
                path = root / filename
                if path.exists():
                    return ImageFont.truetype(str(path), size=pixel_size)
    return ImageFont.load_default()


def color_to_rgb(color: Color | None, fallback: tuple[int, int, int]) -> tuple[int, int, int]:
    if color is None:
        return fallback
    if color.type == "rgb" and color.rgb:
        value = color.rgb[-6:]
        try:
            return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))
        except ValueError:
            return fallback
    if color.type == "indexed" and color.indexed is not None:
        palette = {
            0: (0, 0, 0), 1: (255, 255, 255), 2: (255, 0, 0),
            3: (0, 255, 0), 4: (0, 0, 255), 5: (255, 255, 0),
            6: (255, 0, 255), 7: (0, 255, 255), 8: (0, 0, 0),
            9: (255, 255, 255),
        }
        return palette.get(int(color.indexed), fallback)
    return fallback


def fill_rgb(fill: PatternFill | None) -> tuple[int, int, int]:
    if fill is None or fill.fill_type != "solid":
        return (255, 255, 255)
    return color_to_rgb(fill.fgColor, (255, 255, 255))


def luminance(rgb: tuple[int, int, int]) -> float:
    r, g, b = rgb
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def text_rgb(cell, background: tuple[int, int, int]) -> tuple[int, int, int]:
    sentinel = (-1, -1, -1)
    explicit = color_to_rgb(getattr(cell.font, "color", None), sentinel)
    if explicit != sentinel:
        return explicit
    return (255, 255, 255) if luminance(background) < 110 else (0, 0, 0)


def cell_box(row: int, column: int, end_row: int | None = None, end_column: int | None = None) -> tuple[int, int, int, int]:
    end_row = end_row or row
    end_column = end_column or column
    return (
        (column - 1) * CELL_WIDTH,
        (row - 1) * CELL_HEIGHT,
        end_column * CELL_WIDTH,
        end_row * CELL_HEIGHT,
    )


def merged_maps(sheet):
    covered: set[tuple[int, int]] = set()
    spans: dict[tuple[int, int], tuple[int, int, int, int]] = {}
    for merged in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged.bounds
        if min_row > ROWS or min_col > COLUMNS:
            continue
        max_row = min(max_row, ROWS)
        max_col = min(max_col, COLUMNS)
        spans[(min_row, min_col)] = (min_row, min_col, max_row, max_col)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (row, col) != (min_row, min_col):
                    covered.add((row, col))
    return covered, spans


def side_visible(side) -> bool:
    return side is not None and getattr(side, "style", None) is not None


def draw_border(draw: ImageDraw.ImageDraw, border: Border | None, box: tuple[int, int, int, int]) -> None:
    if border is None:
        return
    left, top, right, bottom = box
    sides = {
        "left": getattr(border, "left", None),
        "right": getattr(border, "right", None),
        "top": getattr(border, "top", None),
        "bottom": getattr(border, "bottom", None),
    }
    if side_visible(sides["left"]):
        draw.line((left, top, left, bottom - 1), fill=(0, 0, 0), width=1)
    if side_visible(sides["right"]):
        draw.line((right - 1, top, right - 1, bottom - 1), fill=(0, 0, 0), width=1)
    if side_visible(sides["top"]):
        draw.line((left, top, right - 1, top), fill=(0, 0, 0), width=1)
    if side_visible(sides["bottom"]):
        draw.line((left, bottom - 1, right - 1, bottom - 1), fill=(0, 0, 0), width=1)


def cell_blocks_overflow(cell) -> bool:
    if cell.value not in (None, ""):
        return True
    if fill_rgb(getattr(cell, "fill", None)) != (255, 255, 255):
        return True
    border = getattr(cell, "border", None)
    if border is not None and any(
        side_visible(getattr(border, name, None)) for name in ("left", "right", "top", "bottom")
    ):
        return True
    return False


def overflow_box(sheet, row: int, col: int, base_box: tuple[int, int, int, int], covered: set[tuple[int, int]]) -> tuple[int, int, int, int]:
    left, top, right, bottom = base_box
    for next_col in range(col + 1, COLUMNS + 1):
        if (row, next_col) in covered or cell_blocks_overflow(sheet.cell(row, next_col)):
            break
        right = next_col * CELL_WIDTH
    return left, top, right, bottom


def fit_single_line(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> str:
    """Return the full single-line text without truncation or ellipses."""
    del draw, font, width
    return text.replace("\r", " ").replace("\n", " ").strip()


def wrap_lines(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int, max_lines: int) -> list[str]:
    paragraphs = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines: list[str] = []
    for paragraph in paragraphs:
        words = paragraph.split()
        if not words:
            lines.append("")
            continue
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if draw.textlength(candidate, font=font) <= width:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
    # Do not truncate wrapped text or add ellipses. Pillow does not clip text to
    # the source cell, so extra lines may overlap content below, as requested.
    del max_lines
    return lines


def draw_checkbox(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], checked: bool) -> None:
    left, top, right, bottom = box
    size = 15
    x = left + (right - left - size) // 2
    y = top + (bottom - top - size) // 2
    outline = (120, 120, 120)
    draw.rounded_rectangle((x, y, x + size - 1, y + size - 1), radius=2, outline=outline, width=2)
    if checked:
        draw.rectangle((x, y, x + size - 1, y + size - 1), fill=(128, 128, 128))
        draw.line((x + 3, y + 8, x + 6, y + 11, x + 12, y + 4), fill=(255, 255, 255), width=2, joint="curve")


def draw_underline(draw: ImageDraw.ImageDraw, font: ImageFont.ImageFont, style: str | None, x: float, y: float, width: float, color: tuple[int, int, int]) -> None:
    if not style:
        return
    ascent, _ = font.getmetrics()
    line_y = round(y + ascent + 1)
    draw.line((round(x), line_y, round(x + width), line_y), fill=color, width=1)
    if str(style) in {"double", "doubleAccounting"}:
        draw.line((round(x), line_y + 2, round(x + width), line_y + 2), fill=color, width=1)


def draw_rotated_text(
    image: Image.Image,
    text: str,
    font: ImageFont.ImageFont,
    color: tuple[int, int, int],
    x: int,
    y: int,
    angle: int,
    underline: str | None,
) -> None:
    """Draw full text at an Excel rotation angle without clipping."""
    scratch = Image.new("RGBA", (max(WIDTH, 1600), max(HEIGHT, 800)), (255, 255, 255, 0))
    scratch_draw = ImageDraw.Draw(scratch)
    scratch_draw.text((2, 2), text, font=font, fill=(*color, 255))
    if underline:
        text_width = scratch_draw.textlength(text, font=font)
        ascent, _ = font.getmetrics()
        underline_y = 2 + ascent + 1
        scratch_draw.line((2, underline_y, 2 + text_width, underline_y), fill=(*color, 255), width=1)
        if str(underline) in {"double", "doubleAccounting"}:
            scratch_draw.line((2, underline_y + 2, 2 + text_width, underline_y + 2), fill=(*color, 255), width=1)
    bbox = scratch.getbbox()
    if bbox is None:
        return
    glyph = scratch.crop(bbox).rotate(angle, expand=True, resample=Image.Resampling.BICUBIC)
    image.alpha_composite(glyph, (x, y))


def draw_cell_text(image: Image.Image, draw: ImageDraw.ImageDraw, sheet, cell, box: tuple[int, int, int, int], covered: set[tuple[int, int]], merged: bool) -> None:
    value = cell.value
    if value in (None, ""):
        return
    if isinstance(value, bool):
        draw_checkbox(draw, box, value)
        return

    background = fill_rgb(cell.fill)
    color = text_rgb(cell, background)
    font = load_font(cell.font.name, cell.font.sz or DEFAULT_FONT_SIZE, bool(cell.font.bold), bool(cell.font.italic))
    alignment = cell.alignment
    wrap = bool(alignment.wrap_text) or "\n" in str(value) or "\r" in str(value)

    text_box = box
    if not wrap and not merged:
        horizontal_mode = alignment.horizontal or "general"
        if horizontal_mode not in {"right", "center", "centerContinuous"}:
            # Let left-aligned text continue all the way to the image edge.
            # It is intentionally allowed to overlap cells, borders, and text.
            text_box = (box[0], box[1], WIDTH, box[3])
        elif horizontal_mode == "right":
            # Likewise, allow right-aligned text to extend toward the left edge.
            text_box = (0, box[1], box[2], box[3])

    left, top, right, bottom = text_box
    available_width = max(1, right - left - 2 * TEXT_PADDING_X)
    bbox = draw.textbbox((0, 0), "Ag", font=font)
    line_height = max(1, bbox[3] - bbox[1] + 2)
    max_lines = max(1, (bottom - top - 2) // line_height)

    if wrap:
        lines = wrap_lines(draw, str(value), font, available_width, max_lines)
    else:
        lines = [fit_single_line(draw, str(value), font, available_width)]

    block_height = len(lines) * line_height
    vertical = alignment.vertical or "center"
    if vertical == "top":
        y = top + 1
    elif vertical == "bottom":
        y = bottom - block_height - 1
    else:
        y = top + (bottom - top - block_height) // 2

    horizontal = alignment.horizontal or "left"
    underline = getattr(cell.font, "underline", None)
    rotation = int(getattr(alignment, "textRotation", 0) or 0)

    # Excel stores ordinary vertical text as 90 degrees. Render the complete
    # value rather than forcing it through the horizontal cell-width logic.
    if rotation not in (0, 255):
        angle = rotation if rotation <= 90 else rotation - 180
        x = box[0] + 2
        y = box[1] + 1
        rgba = image.convert("RGBA")
        draw_rotated_text(rgba, str(value), font, color, x, y, angle, underline)
        image.paste(rgba.convert("RGB"))
        return

    for line in lines:
        line_width = draw.textlength(line, font=font)
        if horizontal in {"center", "centerContinuous"}:
            x = left + (right - left - line_width) / 2
        elif horizontal == "right":
            x = right - TEXT_PADDING_X - line_width
        else:
            x = left + TEXT_PADDING_X
        draw.text((round(x), round(y)), line, font=font, fill=color)
        draw_underline(draw, font, underline, x, y, line_width, color)
        y += line_height


def render_png(workbook_bytes: bytes, output_path: Path) -> None:
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=True)
    if WORKSHEET_NAME:
        if WORKSHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(f"Worksheet {WORKSHEET_NAME!r} not found; available sheets: {workbook.sheetnames}")
        sheet = workbook[WORKSHEET_NAME]
    else:
        sheet = workbook.active

    covered, spans = merged_maps(sheet)
    image = Image.new("RGB", (WIDTH, HEIGHT), (255, 255, 255))
    draw = ImageDraw.Draw(image)

    # Paint cell backgrounds first so overflowing text is never erased later.
    for row in range(1, ROWS + 1):
        for col in range(1, COLUMNS + 1):
            if (row, col) in covered:
                continue
            cell = sheet.cell(row, col)
            if isinstance(cell, MergedCell):
                continue
            start_row, start_col, end_row, end_col = spans.get((row, col), (row, col, row, col))
            box = cell_box(start_row, start_col, end_row, end_col)
            draw.rectangle(box, fill=fill_rgb(cell.fill))

    # Render text and checkbox content at exact worksheet coordinates.
    for row in range(1, ROWS + 1):
        for col in range(1, COLUMNS + 1):
            if (row, col) in covered:
                continue
            cell = sheet.cell(row, col)
            if isinstance(cell, MergedCell):
                continue
            start_row, start_col, end_row, end_col = spans.get((row, col), (row, col, row, col))
            box = cell_box(start_row, start_col, end_row, end_col)
            draw_cell_text(image, draw, sheet, cell, box, covered, (row, col) in spans)

    # Explicit workbook borders are drawn last. No artificial grid is added.
    for row in range(1, ROWS + 1):
        for col in range(1, COLUMNS + 1):
            if (row, col) in covered:
                continue
            cell = sheet.cell(row, col)
            if isinstance(cell, MergedCell):
                continue
            start_row, start_col, end_row, end_col = spans.get((row, col), (row, col, row, col))
            draw_border(draw, cell.border, cell_box(start_row, start_col, end_row, end_col))

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
