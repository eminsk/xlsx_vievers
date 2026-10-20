"""
High-Quality Vector Icon Engine for Excel Viewer Pro.
Generates modern, crisp, anti-aliased Microsoft Excel-style icons using PIL vector drawing
with 4x supersampling and Lanczos downsampling.
"""

from __future__ import annotations

import math
from typing import Callable, Dict, Tuple
from PIL import Image, ImageDraw, ImageTk


# Cache: (name, size) -> ImageTk.PhotoImage
_photo_cache: Dict[Tuple[str, int], ImageTk.PhotoImage] = {}
# Cache: (name, size) -> Image.Image
_pil_cache: Dict[Tuple[str, int], Image.Image] = {}


def _supersample(draw_fn: Callable[[ImageDraw.ImageDraw, int], None], size: int = 16, factor: int = 4) -> Image.Image:
    s = size * factor
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    draw_fn(d, s)
    return img.resize((size, size), Image.Resampling.LANCZOS)


# =============================================================================
# Vector Drawing Routines
# =============================================================================

# --- File & History ---
def _draw_save(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.12, s*0.1, s*0.88, s*0.9], radius=s*0.08, fill=(33, 115, 70), outline=(16, 75, 40), width=int(s*0.04))
    d.rectangle([s*0.28, s*0.1, s*0.72, s*0.38], fill=(235, 240, 245), outline=(160, 170, 180), width=int(s*0.03))
    d.rectangle([s*0.36, s*0.14, s*0.48, s*0.34], fill=(70, 80, 90))
    d.rounded_rectangle([s*0.24, s*0.5, s*0.76, s*0.9], radius=s*0.04, fill=(255, 255, 255), outline=(180, 190, 200), width=int(s*0.03))
    d.line([s*0.32, s*0.62, s*0.68, s*0.62], fill=(33, 115, 70), width=int(s*0.04))
    d.line([s*0.32, s*0.74, s*0.68, s*0.74], fill=(33, 115, 70), width=int(s*0.04))

def _draw_undo(d: ImageDraw.ImageDraw, s: int) -> None:
    d.arc([s*0.2, s*0.25, s*0.85, s*0.85], start=180, end=360, fill=(24, 90, 189), width=int(s*0.14))
    d.polygon([(s*0.12, s*0.55), (s*0.38, s*0.3), (s*0.38, s*0.8)], fill=(24, 90, 189))

def _draw_redo(d: ImageDraw.ImageDraw, s: int) -> None:
    d.arc([s*0.15, s*0.25, s*0.8, s*0.85], start=180, end=360, fill=(24, 90, 189), width=int(s*0.14))
    d.polygon([(s*0.88, s*0.55), (s*0.62, s*0.3), (s*0.62, s*0.8)], fill=(24, 90, 189))

# --- Clipboard ---
def _draw_paste(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.18, s*0.15, s*0.82, s*0.92], radius=s*0.06, fill=(217, 160, 102), outline=(160, 110, 60), width=int(s*0.04))
    d.rectangle([s*0.28, s*0.25, s*0.72, s*0.85], fill=(255, 255, 255), outline=(200, 205, 215), width=int(s*0.03))
    d.rounded_rectangle([s*0.35, s*0.08, s*0.65, s*0.22], radius=s*0.04, fill=(190, 195, 205), outline=(100, 110, 120), width=int(s*0.03))
    d.line([s*0.36, s*0.4, s*0.64, s*0.4], fill=(70, 130, 180), width=int(s*0.04))
    d.line([s*0.36, s*0.52, s*0.64, s*0.52], fill=(70, 130, 180), width=int(s*0.04))
    d.line([s*0.36, s*0.64, s*0.55, s*0.64], fill=(70, 130, 180), width=int(s*0.04))

def _draw_cut(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.2, s*0.2), (s*0.8, s*0.8)], fill=(120, 130, 140), width=int(s*0.08))
    d.line([(s*0.8, s*0.2), (s*0.2, s*0.8)], fill=(120, 130, 140), width=int(s*0.08))
    d.ellipse([s*0.12, s*0.7, s*0.38, s*0.95], outline=(24, 90, 189), width=int(s*0.08))
    d.ellipse([s*0.62, s*0.7, s*0.88, s*0.95], outline=(24, 90, 189), width=int(s*0.08))

def _draw_copy(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.25, s*0.12, s*0.82, s*0.72], radius=s*0.04, fill=(240, 244, 250), outline=(140, 160, 180), width=int(s*0.04))
    d.rounded_rectangle([s*0.14, s*0.25, s*0.71, s*0.88], radius=s*0.04, fill=(255, 255, 255), outline=(24, 90, 189), width=int(s*0.04))
    d.line([s*0.24, s*0.42, s*0.61, s*0.42], fill=(160, 175, 190), width=int(s*0.04))
    d.line([s*0.24, s*0.55, s*0.61, s*0.55], fill=(160, 175, 190), width=int(s*0.04))
    d.line([s*0.24, s*0.68, s*0.48, s*0.68], fill=(160, 175, 190), width=int(s*0.04))

def _draw_format_painter(d: ImageDraw.ImageDraw, s: int) -> None:
    d.polygon([(s*0.65, s*0.12), (s*0.88, s*0.35), (s*0.65, s*0.58), (s*0.42, s*0.35)], fill=(185, 122, 60))
    d.polygon([(s*0.42, s*0.35), (s*0.65, s*0.58), (s*0.5, s*0.72), (s*0.28, s*0.5)], fill=(180, 185, 195))
    d.polygon([(s*0.28, s*0.5), (s*0.5, s*0.72), (s*0.32, s*0.88), (s*0.12, s*0.68)], fill=(33, 115, 70))
    d.ellipse([s*0.1, s*0.78, s*0.25, s*0.92], fill=(255, 192, 0))

# --- Font & Styles ---
def _draw_font_inc(d: ImageDraw.ImageDraw, s: int) -> None:
    # Letter A
    d.line([(s*0.2, s*0.8), (s*0.42, s*0.25)], fill=(40, 45, 50), width=int(s*0.07))
    d.line([(s*0.42, s*0.25), (s*0.64, s*0.8)], fill=(40, 45, 50), width=int(s*0.07))
    d.line([(s*0.28, s*0.62), (s*0.56, s*0.62)], fill=(40, 45, 50), width=int(s*0.06))
    # Green up arrow
    d.polygon([(s*0.8, s*0.25), (s*0.68, s*0.45), (s*0.92, s*0.45)], fill=(33, 115, 70))
    d.rectangle([s*0.76, s*0.45, s*0.84, s*0.65], fill=(33, 115, 70))

def _draw_font_dec(d: ImageDraw.ImageDraw, s: int) -> None:
    # Letter A
    d.line([(s*0.2, s*0.8), (s*0.42, s*0.25)], fill=(40, 45, 50), width=int(s*0.07))
    d.line([(s*0.42, s*0.25), (s*0.64, s*0.8)], fill=(40, 45, 50), width=int(s*0.07))
    d.line([(s*0.28, s*0.62), (s*0.56, s*0.62)], fill=(40, 45, 50), width=int(s*0.06))
    # Red down arrow
    d.polygon([(s*0.8, s*0.65), (s*0.68, s*0.45), (s*0.92, s*0.45)], fill=(215, 45, 30))
    d.rectangle([s*0.76, s*0.25, s*0.84, s*0.45], fill=(215, 45, 30))

def _draw_bold(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.25, s*0.18), (s*0.25, s*0.82)], fill=(30, 35, 40), width=int(s*0.14))
    d.arc([s*0.25, s*0.18, s*0.7, s*0.52], start=270, end=90, fill=(30, 35, 40), width=int(s*0.12))
    d.arc([s*0.25, s*0.48, s*0.78, s*0.82], start=270, end=90, fill=(30, 35, 40), width=int(s*0.12))
    d.line([(s*0.25, s*0.18), (s*0.48, s*0.18)], fill=(30, 35, 40), width=int(s*0.12))
    d.line([(s*0.25, s*0.5), (s*0.52, s*0.5)], fill=(30, 35, 40), width=int(s*0.12))
    d.line([(s*0.25, s*0.82), (s*0.52, s*0.82)], fill=(30, 35, 40), width=int(s*0.12))

def _draw_italic(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.35, s*0.2), (s*0.72, s*0.2)], fill=(30, 35, 40), width=int(s*0.09))
    d.line([(s*0.28, s*0.8), (s*0.65, s*0.8)], fill=(30, 35, 40), width=int(s*0.09))
    d.line([(s*0.58, s*0.2), (s*0.42, s*0.8)], fill=(30, 35, 40), width=int(s*0.1))

def _draw_underline(d: ImageDraw.ImageDraw, s: int) -> None:
    d.arc([s*0.25, s*0.35, s*0.75, s*0.7], start=0, end=180, fill=(30, 35, 40), width=int(s*0.09))
    d.line([(s*0.25, s*0.2), (s*0.25, s*0.52)], fill=(30, 35, 40), width=int(s*0.09))
    d.line([(s*0.75, s*0.2), (s*0.75, s*0.52)], fill=(30, 35, 40), width=int(s*0.09))
    d.line([(s*0.2, s*0.84), (s*0.8, s*0.84)], fill=(30, 35, 40), width=int(s*0.09))

def _draw_strike(d: ImageDraw.ImageDraw, s: int) -> None:
    d.arc([s*0.28, s*0.18, s*0.72, s*0.52], start=90, end=270, fill=(30, 35, 40), width=int(s*0.09))
    d.arc([s*0.28, s*0.48, s*0.72, s*0.82], start=270, end=90, fill=(30, 35, 40), width=int(s*0.09))
    d.line([(s*0.15, s*0.5), (s*0.85, s*0.5)], fill=(215, 45, 30), width=int(s*0.08))

def _draw_borders(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.15, s*0.85, s*0.85], outline=(40, 50, 60), width=int(s*0.06), fill=(255, 255, 255))
    d.line([(s*0.5, s*0.15), (s*0.5, s*0.85)], fill=(120, 130, 140), width=int(s*0.05))
    d.line([(s*0.15, s*0.5), (s*0.85, s*0.5)], fill=(120, 130, 140), width=int(s*0.05))

def _draw_fill_color(d: ImageDraw.ImageDraw, s: int) -> None:
    # Paint bucket
    d.polygon([(s*0.35, s*0.2), (s*0.75, s*0.5), (s*0.55, s*0.8), (s*0.15, s*0.5)], fill=(160, 170, 180), outline=(90, 100, 110), width=int(s*0.03))
    # Tipped paint pouring
    d.ellipse([s*0.62, s*0.62, s*0.82, s*0.85], fill=(255, 215, 0))
    # Colored bottom bar
    d.rectangle([s*0.15, s*0.85, s*0.85, s*0.95], fill=(255, 215, 0))

def _draw_font_color(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.22, s*0.7), (s*0.5, s*0.15)], fill=(40, 45, 50), width=int(s*0.08))
    d.line([(s*0.5, s*0.15), (s*0.78, s*0.7)], fill=(40, 45, 50), width=int(s*0.08))
    d.line([(s*0.32, s*0.52), (s*0.68, s*0.52)], fill=(40, 45, 50), width=int(s*0.07))
    d.rectangle([s*0.15, s*0.82, s*0.85, s*0.94], fill=(220, 30, 30))

# --- Alignment ---
def _draw_align_top(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.15, s*0.2), (s*0.85, s*0.2)], fill=(24, 90, 189), width=int(s*0.08))
    d.line([(s*0.25, s*0.42), (s*0.75, s*0.42)], fill=(90, 100, 110), width=int(s*0.06))
    d.line([(s*0.35, s*0.62), (s*0.65, s*0.62)], fill=(90, 100, 110), width=int(s*0.06))

def _draw_align_mid(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.2, s*0.32), (s*0.8, s*0.32)], fill=(90, 100, 110), width=int(s*0.06))
    d.line([(s*0.15, s*0.5), (s*0.85, s*0.5)], fill=(24, 90, 189), width=int(s*0.08))
    d.line([(s*0.2, s*0.68), (s*0.8, s*0.68)], fill=(90, 100, 110), width=int(s*0.06))

def _draw_align_bot(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.35, s*0.38), (s*0.65, s*0.38)], fill=(90, 100, 110), width=int(s*0.06))
    d.line([(s*0.25, s*0.58), (s*0.75, s*0.58)], fill=(90, 100, 110), width=int(s*0.06))
    d.line([(s*0.15, s*0.8), (s*0.85, s*0.8)], fill=(24, 90, 189), width=int(s*0.08))

def _draw_align_left(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.15, s*0.25), (s*0.85, s*0.25)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.15, s*0.45), (s*0.6, s*0.45)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.15, s*0.65), (s*0.85, s*0.65)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.15, s*0.85), (s*0.5, s*0.85)], fill=(50, 60, 70), width=int(s*0.06))

def _draw_align_center(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.15, s*0.25), (s*0.85, s*0.25)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.25, s*0.45), (s*0.75, s*0.45)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.15, s*0.65), (s*0.85, s*0.65)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.3, s*0.85), (s*0.7, s*0.85)], fill=(50, 60, 70), width=int(s*0.06))

def _draw_align_right(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.15, s*0.25), (s*0.85, s*0.25)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.4, s*0.45), (s*0.85, s*0.45)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.15, s*0.65), (s*0.85, s*0.65)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.5, s*0.85), (s*0.85, s*0.85)], fill=(50, 60, 70), width=int(s*0.06))

def _draw_wrap_text(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.15, s*0.28), (s*0.85, s*0.28)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.15, s*0.5), (s*0.6, s*0.5)], fill=(50, 60, 70), width=int(s*0.06))
    d.arc([s*0.45, s*0.45, s*0.85, s*0.82], start=270, end=90, fill=(24, 90, 189), width=int(s*0.07))
    d.line([(s*0.3, s*0.82), (s*0.65, s*0.82)], fill=(24, 90, 189), width=int(s*0.07))
    d.polygon([(s*0.2, s*0.82), (s*0.35, s*0.72), (s*0.35, s*0.92)], fill=(24, 90, 189))

def _draw_merge_center(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.1, s*0.25, s*0.9, s*0.75], outline=(40, 50, 60), width=int(s*0.05), fill=(255, 255, 255))
    d.line([(s*0.5, s*0.25), (s*0.5, s*0.4)], fill=(180, 190, 200), width=int(s*0.04))
    d.line([(s*0.5, s*0.6), (s*0.5, s*0.75)], fill=(180, 190, 200), width=int(s*0.04))
    # inward arrows
    d.polygon([(s*0.38, s*0.5), (s*0.28, s*0.4), (s*0.28, s*0.6)], fill=(24, 90, 189))
    d.polygon([(s*0.62, s*0.5), (s*0.72, s*0.4), (s*0.72, s*0.6)], fill=(24, 90, 189))

# --- Number Formatting ---
def _draw_currency_rub(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.3, s*0.18), (s*0.3, s*0.82)], fill=(33, 115, 70), width=int(s*0.09))
    d.arc([s*0.3, s*0.18, s*0.75, s*0.55], start=270, end=90, fill=(33, 115, 70), width=int(s*0.09))
    d.line([(s*0.3, s*0.18), (s*0.52, s*0.18)], fill=(33, 115, 70), width=int(s*0.09))
    d.line([(s*0.3, s*0.55), (s*0.52, s*0.55)], fill=(33, 115, 70), width=int(s*0.09))
    d.line([(s*0.18, s*0.65), (s*0.62, s*0.65)], fill=(33, 115, 70), width=int(s*0.08))

def _draw_currency_usd(d: ImageDraw.ImageDraw, s: int) -> None:
    d.arc([s*0.25, s*0.2, s*0.75, s*0.52], start=90, end=270, fill=(33, 115, 70), width=int(s*0.09))
    d.arc([s*0.25, s*0.48, s*0.75, s*0.8], start=270, end=90, fill=(33, 115, 70), width=int(s*0.09))
    d.line([(s*0.5, s*0.12), (s*0.5, s*0.88)], fill=(33, 115, 70), width=int(s*0.08))

def _draw_percent(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.78, s*0.2), (s*0.22, s*0.8)], fill=(24, 90, 189), width=int(s*0.09))
    d.ellipse([s*0.22, s*0.22, s*0.42, s*0.42], outline=(24, 90, 189), width=int(s*0.06))
    d.ellipse([s*0.58, s*0.58, s*0.78, s*0.78], outline=(24, 90, 189), width=int(s*0.06))

def _draw_comma(d: ImageDraw.ImageDraw, s: int) -> None:
    d.ellipse([s*0.38, s*0.32, s*0.62, s*0.56], fill=(50, 60, 70))
    d.polygon([(s*0.58, s*0.45), (s*0.52, s*0.78), (s*0.38, s*0.52)], fill=(50, 60, 70))

def _draw_decimal_inc(d: ImageDraw.ImageDraw, s: int) -> None:
    d.ellipse([s*0.2, s*0.35, s*0.42, s*0.65], outline=(50, 60, 70), width=int(s*0.06))
    d.polygon([(s*0.82, s*0.5), (s*0.58, s*0.35), (s*0.58, s*0.65)], fill=(24, 90, 189))
    d.line([(s*0.5, s*0.5), (s*0.75, s*0.5)], fill=(24, 90, 189), width=int(s*0.07))

def _draw_decimal_dec(d: ImageDraw.ImageDraw, s: int) -> None:
    d.ellipse([s*0.58, s*0.35, s*0.8, s*0.65], outline=(50, 60, 70), width=int(s*0.06))
    d.polygon([(s*0.18, s*0.5), (s*0.42, s*0.35), (s*0.42, s*0.65)], fill=(24, 90, 189))
    d.line([(s*0.25, s*0.5), (s*0.5, s*0.5)], fill=(24, 90, 189), width=int(s*0.07))

# --- Styles & Cells ---
def _draw_cond_format(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.12, s*0.15, s*0.88, s*0.85], outline=(40, 50, 60), width=int(s*0.04), fill=(255, 255, 255))
    d.rectangle([s*0.14, s*0.17, s*0.5, s*0.5], fill=(144, 238, 144))
    d.rectangle([s*0.5, s*0.17, s*0.86, s*0.5], fill=(255, 255, 150))
    d.rectangle([s*0.14, s*0.5, s*0.5, s*0.83], fill=(255, 160, 150))
    d.rectangle([s*0.5, s*0.5, s*0.86, s*0.83], fill=(144, 238, 144))

def _draw_format_cells(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.12, s*0.15, s*0.88, s*0.85], outline=(33, 115, 70), width=int(s*0.05), fill=(255, 255, 255))
    d.rectangle([s*0.12, s*0.15, s*0.88, s*0.38], fill=(33, 115, 70))
    d.line([(s*0.5, s*0.38), (s*0.5, s*0.85)], fill=(180, 190, 200), width=int(s*0.04))
    d.line([(s*0.12, s*0.62), (s*0.88, s*0.62)], fill=(180, 190, 200), width=int(s*0.04))

def _draw_insert(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.1, s*0.15, s*0.9, s*0.85], radius=s*0.06, outline=(140, 150, 160), width=int(s*0.04), fill=(255, 255, 255))
    d.line([(s*0.1, s*0.4), (s*0.9, s*0.4)], fill=(140, 150, 160), width=int(s*0.04))
    # Green plus
    d.ellipse([s*0.48, s*0.45, s*0.92, s*0.89], fill=(33, 115, 70))
    d.line([(s*0.7, s*0.55), (s*0.7, s*0.79)], fill=(255, 255, 255), width=int(s*0.07))
    d.line([(s*0.58, s*0.67), (s*0.82, s*0.67)], fill=(255, 255, 255), width=int(s*0.07))

def _draw_delete(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.1, s*0.15, s*0.9, s*0.85], radius=s*0.06, outline=(140, 150, 160), width=int(s*0.04), fill=(255, 255, 255))
    d.line([(s*0.1, s*0.4), (s*0.9, s*0.4)], fill=(140, 150, 160), width=int(s*0.04))
    # Red minus / cross
    d.ellipse([s*0.48, s*0.45, s*0.92, s*0.89], fill=(215, 45, 30))
    d.line([(s*0.58, s*0.67), (s*0.82, s*0.67)], fill=(255, 255, 255), width=int(s*0.07))

def _draw_autofit(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.2, s*0.15), (s*0.2, s*0.85)], fill=(40, 50, 60), width=int(s*0.06))
    d.line([(s*0.8, s*0.15), (s*0.8, s*0.85)], fill=(40, 50, 60), width=int(s*0.06))
    d.line([(s*0.25, s*0.5), (s*0.75, s*0.5)], fill=(24, 90, 189), width=int(s*0.06))
    d.polygon([(s*0.22, s*0.5), (s*0.35, s*0.4), (s*0.35, s*0.6)], fill=(24, 90, 189))
    d.polygon([(s*0.78, s*0.5), (s*0.65, s*0.4), (s*0.65, s*0.6)], fill=(24, 90, 189))

# --- Editing ---
def _draw_autosum(d: ImageDraw.ImageDraw, s: int) -> None:
    # Sigma symbol Σ
    d.line([(s*0.75, s*0.2), (s*0.25, s*0.2)], fill=(24, 90, 189), width=int(s*0.09))
    d.line([(s*0.25, s*0.2), (s*0.55, s*0.5)], fill=(24, 90, 189), width=int(s*0.09))
    d.line([(s*0.55, s*0.5), (s*0.25, s*0.8)], fill=(24, 90, 189), width=int(s*0.09))
    d.line([(s*0.25, s*0.8), (s*0.75, s*0.8)], fill=(24, 90, 189), width=int(s*0.09))

def _draw_clear(d: ImageDraw.ImageDraw, s: int) -> None:
    # Eraser
    d.polygon([(s*0.35, s*0.18), (s*0.78, s*0.61), (s*0.58, s*0.81), (s*0.15, s*0.38)], fill=(255, 175, 190), outline=(200, 100, 120), width=int(s*0.04))
    d.polygon([(s*0.48, s*0.31), (s*0.78, s*0.61), (s*0.58, s*0.81), (s*0.28, s*0.51)], fill=(24, 90, 189))

def _draw_sort_az(d: ImageDraw.ImageDraw, s: int) -> None:
    # 'A' and 'Z' text hint with down arrow
    d.line([(s*0.15, s*0.42), (s*0.28, s*0.18)], fill=(30, 40, 50), width=int(s*0.06))
    d.line([(s*0.28, s*0.18), (s*0.41, s*0.42)], fill=(30, 40, 50), width=int(s*0.06))
    d.line([(s*0.2, s*0.33), (s*0.36, s*0.33)], fill=(30, 40, 50), width=int(s*0.05))
    # Down arrow
    d.line([(s*0.72, s*0.2), (s*0.72, s*0.8)], fill=(24, 90, 189), width=int(s*0.08))
    d.polygon([(s*0.72, s*0.88), (s*0.58, s*0.68), (s*0.86, s*0.68)], fill=(24, 90, 189))

def _draw_sort_za(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([(s*0.15, s*0.18), (s*0.4, s*0.18)], fill=(30, 40, 50), width=int(s*0.06))
    d.line([(s*0.4, s*0.18), (s*0.15, s*0.42)], fill=(30, 40, 50), width=int(s*0.06))
    d.line([(s*0.15, s*0.42), (s*0.4, s*0.42)], fill=(30, 40, 50), width=int(s*0.06))
    # Down arrow
    d.line([(s*0.72, s*0.2), (s*0.72, s*0.8)], fill=(24, 90, 189), width=int(s*0.08))
    d.polygon([(s*0.72, s*0.88), (s*0.58, s*0.68), (s*0.86, s*0.68)], fill=(24, 90, 189))

def _draw_sort_custom(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.2, s*0.85, s*0.34], fill=(24, 90, 189))
    d.rectangle([s*0.15, s*0.44, s*0.65, s*0.58], fill=(33, 115, 70))
    d.rectangle([s*0.15, s*0.68, s*0.45, s*0.82], fill=(230, 120, 30))

def _draw_filter(d: ImageDraw.ImageDraw, s: int) -> None:
    d.polygon([(s*0.15, s*0.18), (s*0.85, s*0.18), (s*0.58, s*0.52), (s*0.58, s*0.82), (s*0.42, s*0.72), (s*0.42, s*0.52)], fill=(24, 90, 189), outline=(15, 60, 130), width=int(s*0.03))

def _draw_clear_filter(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_filter(d, s)
    d.line([(s*0.58, s*0.48), (s*0.92, s*0.82)], fill=(215, 45, 30), width=int(s*0.09))
    d.line([(s*0.92, s*0.48), (s*0.58, s*0.82)], fill=(215, 45, 30), width=int(s*0.09))

def _draw_find(d: ImageDraw.ImageDraw, s: int) -> None:
    d.ellipse([s*0.18, s*0.18, s*0.62, s*0.62], outline=(24, 90, 189), width=int(s*0.08), fill=(235, 245, 255))
    d.line([(s*0.55, s*0.55), (s*0.85, s*0.85)], fill=(50, 60, 70), width=int(s*0.12))

def _draw_replace(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_find(d, s)
    d.polygon([(s*0.75, s*0.35), (s*0.88, s*0.25), (s*0.88, s*0.45)], fill=(33, 115, 70))

def _draw_goto(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.2, s*0.85, s*0.8], outline=(150, 160, 170), width=int(s*0.04), fill=(255, 255, 255))
    d.polygon([(s*0.4, s*0.6), (s*0.65, s*0.4), (s*0.4, s*0.25)], fill=(24, 90, 189))

# --- Insert Tab ---
def _draw_demo_table(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.12, s*0.15, s*0.88, s*0.85], radius=s*0.05, fill=(255, 255, 255), outline=(33, 115, 70), width=int(s*0.05))
    d.rectangle([s*0.12, s*0.15, s*0.88, s*0.38], fill=(33, 115, 70))
    d.line([(s*0.12, s*0.62), (s*0.88, s*0.62)], fill=(220, 230, 225), width=int(s*0.04))
    d.line([(s*0.5, s*0.38), (s*0.5, s*0.85)], fill=(220, 230, 225), width=int(s*0.04))

def _draw_new_sheet(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.2, s*0.12, s*0.8, s*0.88], radius=s*0.05, fill=(255, 255, 255), outline=(150, 160, 170), width=int(s*0.04))
    d.ellipse([s*0.48, s*0.48, s*0.9, s*0.9], fill=(33, 115, 70))
    d.line([(s*0.69, s*0.56), (s*0.69, s*0.82)], fill=(255, 255, 255), width=int(s*0.06))
    d.line([(s*0.56, s*0.69), (s*0.82, s*0.69)], fill=(255, 255, 255), width=int(s*0.06))

def _draw_chart_wizard(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_chart_col(d, s)
    # Magic star / spark
    d.polygon([(s*0.8, s*0.12), (s*0.84, s*0.22), (s*0.94, s*0.24), (s*0.86, s*0.31), (s*0.89, s*0.41), (s*0.8, s*0.34), (s*0.71, s*0.41), (s*0.74, s*0.31), (s*0.66, s*0.24), (s*0.76, s*0.22)], fill=(255, 200, 0))

def _draw_chart_col(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.45, s*0.35, s*0.85], fill=(24, 90, 189), outline=(15, 60, 130), width=int(s*0.03))
    d.rectangle([s*0.4, s*0.25, s*0.6, s*0.85], fill=(33, 115, 70), outline=(20, 80, 45), width=int(s*0.03))
    d.rectangle([s*0.65, s*0.55, s*0.85, s*0.85], fill=(230, 120, 30), outline=(160, 80, 20), width=int(s*0.03))
    d.line([s*0.1, s*0.85, s*0.9, s*0.85], fill=(100, 110, 120), width=int(s*0.04))

def _draw_chart_line(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([s*0.15, s*0.85, s*0.85, s*0.85], fill=(100, 110, 120), width=int(s*0.04))
    d.line([s*0.15, s*0.15, s*0.15, s*0.85], fill=(100, 110, 120), width=int(s*0.04))
    points = [(s*0.2, s*0.65), (s*0.4, s*0.4), (s*0.6, s*0.55), (s*0.8, s*0.25)]
    d.line(points, fill=(24, 90, 189), width=int(s*0.07))
    for p in points:
        d.ellipse([p[0]-s*0.06, p[1]-s*0.06, p[0]+s*0.06, p[1]+s*0.06], fill=(230, 60, 40))

def _draw_chart_pie(d: ImageDraw.ImageDraw, s: int) -> None:
    bbox = [s*0.12, s*0.12, s*0.88, s*0.88]
    d.pieslice(bbox, start=0, end=110, fill=(24, 90, 189), outline=(255, 255, 255), width=int(s*0.03))
    d.pieslice(bbox, start=110, end=230, fill=(33, 115, 70), outline=(255, 255, 255), width=int(s*0.03))
    d.pieslice(bbox, start=230, end=360, fill=(230, 120, 30), outline=(255, 255, 255), width=int(s*0.03))

def _draw_chart_bar(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([s*0.15, s*0.15, s*0.15, s*0.85], fill=(100, 110, 120), width=int(s*0.04))
    d.rectangle([s*0.15, s*0.2, s*0.75, s*0.35], fill=(24, 90, 189))
    d.rectangle([s*0.15, s*0.45, s*0.88, s*0.6], fill=(33, 115, 70))
    d.rectangle([s*0.15, s*0.7, s*0.55, s*0.85], fill=(230, 120, 30))

def _draw_chart_area(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([s*0.15, s*0.85, s*0.85, s*0.85], fill=(100, 110, 120), width=int(s*0.04))
    pts = [(s*0.15, s*0.85), (s*0.15, s*0.6), (s*0.35, s*0.35), (s*0.6, s*0.5), (s*0.85, s*0.2), (s*0.85, s*0.85)]
    d.polygon(pts, fill=(180, 210, 255), outline=(24, 90, 189), width=int(s*0.04))

def _draw_chart_scatter(d: ImageDraw.ImageDraw, s: int) -> None:
    d.line([s*0.15, s*0.85, s*0.85, s*0.85], fill=(100, 110, 120), width=int(s*0.04))
    d.line([s*0.15, s*0.15, s*0.15, s*0.85], fill=(100, 110, 120), width=int(s*0.04))
    dots = [(s*0.3, s*0.7), (s*0.4, s*0.45), (s*0.55, s*0.58), (s*0.7, s*0.3), (s*0.8, s*0.2)]
    for p in dots:
        d.ellipse([p[0]-s*0.06, p[1]-s*0.06, p[0]+s*0.06, p[1]+s*0.06], fill=(24, 90, 189))

def _draw_comment_add(d: ImageDraw.ImageDraw, s: int) -> None:
    d.polygon([(s*0.15, s*0.2), (s*0.85, s*0.2), (s*0.85, s*0.7), (s*0.45, s*0.7), (s*0.25, s*0.88), (s*0.28, s*0.7), (s*0.15, s*0.7)], fill=(255, 230, 140), outline=(210, 170, 70), width=int(s*0.04))
    d.line([(s*0.28, s*0.38), (s*0.72, s*0.38)], fill=(120, 100, 40), width=int(s*0.04))
    d.line([(s*0.28, s*0.52), (s*0.58, s*0.52)], fill=(120, 100, 40), width=int(s*0.04))

def _draw_comment_del(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_comment_add(d, s)
    d.line([(s*0.62, s*0.45), (s*0.88, s*0.75)], fill=(215, 45, 30), width=int(s*0.08))
    d.line([(s*0.88, s*0.45), (s*0.62, s*0.75)], fill=(215, 45, 30), width=int(s*0.08))

def _draw_fx(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.1, s*0.1, s*0.9, s*0.9], radius=s*0.15, fill=(33, 115, 70))
    d.line([(s*0.44, s*0.26), (s*0.35, s*0.26)], fill=(255, 255, 255), width=int(s*0.07))
    d.line([(s*0.35, s*0.26), (s*0.35, s*0.74)], fill=(255, 255, 255), width=int(s*0.07))
    d.line([(s*0.25, s*0.45), (s*0.45, s*0.45)], fill=(255, 255, 255), width=int(s*0.07))
    d.line([(s*0.52, s*0.45), (s*0.72, s*0.74)], fill=(255, 255, 255), width=int(s*0.06))
    d.line([(s*0.72, s*0.45), (s*0.52, s*0.74)], fill=(255, 255, 255), width=int(s*0.06))

def _draw_formula_ref(d: ImageDraw.ImageDraw, s: int) -> None:
    # Book
    d.rounded_rectangle([s*0.15, s*0.18, s*0.85, s*0.82], radius=s*0.05, fill=(245, 248, 252), outline=(24, 90, 189), width=int(s*0.04))
    d.line([(s*0.5, s*0.18), (s*0.5, s*0.82)], fill=(24, 90, 189), width=int(s*0.05))
    d.line([(s*0.24, s*0.35), (s*0.42, s*0.35)], fill=(120, 140, 160), width=int(s*0.04))
    d.line([(s*0.24, s*0.5), (s*0.42, s*0.5)], fill=(120, 140, 160), width=int(s*0.04))
    d.line([(s*0.58, s*0.35), (s*0.76, s*0.35)], fill=(120, 140, 160), width=int(s*0.04))
    d.line([(s*0.58, s*0.5), (s*0.76, s*0.5)], fill=(120, 140, 160), width=int(s*0.04))

def _draw_import_csv(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.2, s*0.12, s*0.8, s*0.88], radius=s*0.04, fill=(255, 255, 255), outline=(150, 160, 170), width=int(s*0.04))
    # Green arrow into document
    d.polygon([(s*0.5, s*0.72), (s*0.32, s*0.5), (s*0.68, s*0.5)], fill=(33, 115, 70))
    d.rectangle([s*0.44, s*0.3, s*0.56, s*0.5], fill=(33, 115, 70))

def _draw_export_csv(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rounded_rectangle([s*0.2, s*0.12, s*0.8, s*0.88], radius=s*0.04, fill=(255, 255, 255), outline=(150, 160, 170), width=int(s*0.04))
    # Blue arrow out of document
    d.polygon([(s*0.5, s*0.28), (s*0.32, s*0.5), (s*0.68, s*0.5)], fill=(24, 90, 189))
    d.rectangle([s*0.44, s*0.5, s*0.56, s*0.7], fill=(24, 90, 189))

# --- Data Tab ---
def _draw_refresh(d: ImageDraw.ImageDraw, s: int) -> None:
    d.arc([s*0.18, s*0.18, s*0.82, s*0.82], start=45, end=225, fill=(24, 90, 189), width=int(s*0.08))
    d.arc([s*0.18, s*0.18, s*0.82, s*0.82], start=225, end=45, fill=(33, 115, 70), width=int(s*0.08))
    d.polygon([(s*0.75, s*0.15), (s*0.88, s*0.28), (s*0.65, s*0.32)], fill=(33, 115, 70))
    d.polygon([(s*0.25, s*0.85), (s*0.12, s*0.72), (s*0.35, s*0.68)], fill=(24, 90, 189))

def _draw_text_to_columns(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.2, s*0.4, s*0.8], fill=(235, 240, 248), outline=(24, 90, 189), width=int(s*0.04))
    d.rectangle([s*0.6, s*0.2, s*0.85, s*0.8], fill=(235, 240, 248), outline=(24, 90, 189), width=int(s*0.04))
    d.line([(s*0.42, s*0.5), (s*0.56, s*0.5)], fill=(33, 115, 70), width=int(s*0.06))

def _draw_remove_duplicates(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.18, s*0.85, s*0.42], fill=(245, 248, 252), outline=(150, 160, 170), width=int(s*0.03))
    d.rectangle([s*0.15, s*0.52, s*0.85, s*0.76], fill=(255, 235, 235), outline=(215, 45, 30), width=int(s*0.03))
    d.line([(s*0.7, s*0.55), (s*0.85, s*0.73)], fill=(215, 45, 30), width=int(s*0.06))
    d.line([(s*0.85, s*0.55), (s*0.7, s*0.73)], fill=(215, 45, 30), width=int(s*0.06))

def _draw_goal_seek(d: ImageDraw.ImageDraw, s: int) -> None:
    # Target bullseye
    d.ellipse([s*0.15, s*0.15, s*0.85, s*0.85], outline=(215, 45, 30), width=int(s*0.06), fill=(255, 255, 255))
    d.ellipse([s*0.3, s*0.3, s*0.7, s*0.7], outline=(24, 90, 189), width=int(s*0.05), fill=(255, 255, 255))
    d.ellipse([s*0.42, s*0.42, s*0.58, s*0.58], fill=(215, 45, 30))

def _draw_freeze_panes(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.15, s*0.85, s*0.85], outline=(150, 160, 170), width=int(s*0.04), fill=(255, 255, 255))
    d.rectangle([s*0.15, s*0.15, s*0.85, s*0.38], fill=(180, 220, 255))
    d.rectangle([s*0.15, s*0.15, s*0.38, s*0.85], fill=(180, 220, 255))
    # Snowflake hint
    d.line([(s*0.58, s*0.58), (s*0.78, s*0.78)], fill=(24, 90, 189), width=int(s*0.05))
    d.line([(s*0.78, s*0.58), (s*0.58, s*0.78)], fill=(24, 90, 189), width=int(s*0.05))

# --- Formulas Tab ---
def _draw_fx_financial(d: ImageDraw.ImageDraw, s: int) -> None:
    # Gold coin
    d.ellipse([s*0.15, s*0.2, s*0.7, s*0.75], fill=(255, 205, 0), outline=(210, 160, 0), width=int(s*0.04))
    d.ellipse([s*0.25, s*0.3, s*0.6, s*0.65], outline=(230, 180, 0), width=int(s*0.03))
    _draw_currency_usd(d, s)

def _draw_fx_logical(d: ImageDraw.ImageDraw, s: int) -> None:
    # Decision diamond
    d.polygon([(s*0.5, s*0.15), (s*0.85, s*0.5), (s*0.5, s*0.85), (s*0.15, s*0.5)], fill=(240, 245, 255), outline=(24, 90, 189), width=int(s*0.05))
    d.line([(s*0.35, s*0.5), (s*0.45, s*0.62), (s*0.65, s*0.38)], fill=(33, 115, 70), width=int(s*0.06))

def _draw_fx_text(d: ImageDraw.ImageDraw, s: int) -> None:
    # Text 'ABC'
    d.rounded_rectangle([s*0.1, s*0.2, s*0.9, s*0.8], radius=s*0.06, fill=(245, 248, 252), outline=(150, 160, 170), width=int(s*0.03))
    d.line([(s*0.22, s*0.7), (s*0.35, s*0.3)], fill=(50, 60, 70), width=int(s*0.05))
    d.line([(s*0.35, s*0.3), (s*0.48, s*0.7)], fill=(50, 60, 70), width=int(s*0.05))
    d.line([(s*0.27, s*0.55), (s*0.43, s*0.55)], fill=(50, 60, 70), width=int(s*0.04))
    # 'B'
    d.line([(s*0.58, s*0.3), (s*0.58, s*0.7)], fill=(24, 90, 189), width=int(s*0.05))
    d.arc([s*0.58, s*0.3, s*0.82, s*0.52], start=270, end=90, fill=(24, 90, 189), width=int(s*0.05))
    d.arc([s*0.58, s*0.48, s*0.85, s*0.7], start=270, end=90, fill=(24, 90, 189), width=int(s*0.05))

def _draw_fx_date(d: ImageDraw.ImageDraw, s: int) -> None:
    # Calendar
    d.rounded_rectangle([s*0.15, s*0.2, s*0.85, s*0.85], radius=s*0.06, fill=(255, 255, 255), outline=(150, 160, 170), width=int(s*0.04))
    d.rectangle([s*0.15, s*0.2, s*0.85, s*0.4], fill=(215, 45, 30))
    d.line([(s*0.32, s*0.12), (s*0.32, s*0.25)], fill=(50, 60, 70), width=int(s*0.06))
    d.line([(s*0.68, s*0.12), (s*0.68, s*0.25)], fill=(50, 60, 70), width=int(s*0.06))
    d.rectangle([s*0.32, s*0.52, s*0.45, s*0.65], fill=(24, 90, 189))
    d.rectangle([s*0.55, s*0.52, s*0.68, s*0.65], fill=(180, 190, 200))

def _draw_fx_lookup(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.15, s*0.85, s*0.85], outline=(150, 160, 170), width=int(s*0.04), fill=(255, 255, 255))
    d.line([(s*0.15, s*0.45), (s*0.85, s*0.45)], fill=(180, 190, 200), width=int(s*0.03))
    _draw_find(d, s)

def _draw_fx_math(d: ImageDraw.ImageDraw, s: int) -> None:
    # + and ×
    d.line([(s*0.32, s*0.2), (s*0.32, s*0.45)], fill=(33, 115, 70), width=int(s*0.06))
    d.line([(s*0.2, s*0.32), (s*0.45, s*0.32)], fill=(33, 115, 70), width=int(s*0.06))
    d.line([(s*0.2, s*0.7), (s*0.45, s*0.7)], fill=(24, 90, 189), width=int(s*0.06))
    d.line([(s*0.6, s*0.22), (s*0.8, s*0.42)], fill=(230, 120, 30), width=int(s*0.06))
    d.line([(s*0.8, s*0.22), (s*0.6, s*0.42)], fill=(230, 120, 30), width=int(s*0.06))

def _draw_fx_stats(d: ImageDraw.ImageDraw, s: int) -> None:
    # Bell curve
    d.line([s*0.15, s*0.85, s*0.85, s*0.85], fill=(120, 130, 140), width=int(s*0.04))
    curve = [(s*0.15, s*0.85), (s*0.35, s*0.7), (s*0.5, s*0.25), (s*0.65, s*0.7), (s*0.85, s*0.85)]
    d.line(curve, fill=(24, 90, 189), width=int(s*0.06))

def _draw_show_formulas(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.1, s*0.25, s*0.9, s*0.75], outline=(33, 115, 70), width=int(s*0.05), fill=(245, 250, 245))
    d.line([(s*0.2, s*0.45), (s*0.35, s*0.45)], fill=(33, 115, 70), width=int(s*0.05))
    d.line([(s*0.2, s*0.55), (s*0.35, s*0.55)], fill=(33, 115, 70), width=int(s*0.05))
    d.line([(s*0.45, s*0.5), (s*0.8, s*0.5)], fill=(50, 60, 70), width=int(s*0.06))

def _draw_auditor(d: ImageDraw.ImageDraw, s: int) -> None:
    # Nodes tree
    d.ellipse([s*0.15, s*0.2, s*0.38, s*0.43], fill=(33, 115, 70))
    d.ellipse([s*0.15, s*0.57, s*0.38, s*0.8], fill=(33, 115, 70))
    d.ellipse([s*0.62, s*0.38, s*0.85, s*0.62], fill=(24, 90, 189))
    d.line([(s*0.38, s*0.31), (s*0.62, s*0.5)], fill=(120, 130, 140), width=int(s*0.04))
    d.line([(s*0.38, s*0.69), (s*0.62, s*0.5)], fill=(120, 130, 140), width=int(s*0.04))

def _draw_calculate(d: ImageDraw.ImageDraw, s: int) -> None:
    # Calculator
    d.rounded_rectangle([s*0.2, s*0.15, s*0.8, s*0.85], radius=s*0.06, fill=(245, 248, 252), outline=(24, 90, 189), width=int(s*0.04))
    d.rectangle([s*0.3, s*0.25, s*0.7, s*0.4], fill=(200, 230, 210))
    d.ellipse([s*0.32, s*0.5, s*0.44, s*0.62], fill=(80, 90, 100))
    d.ellipse([s*0.56, s*0.5, s*0.68, s*0.62], fill=(80, 90, 100))
    d.ellipse([s*0.32, s*0.68, s*0.44, s*0.8], fill=(80, 90, 100))
    d.ellipse([s*0.56, s*0.68, s*0.68, s*0.8], fill=(24, 90, 189))

def _draw_simd(d: ImageDraw.ImageDraw, s: int) -> None:
    # Microchip CPU
    d.rounded_rectangle([s*0.25, s*0.25, s*0.75, s*0.75], radius=s*0.04, fill=(40, 45, 55), outline=(24, 90, 189), width=int(s*0.04))
    for i in range(3):
        pos = s * (0.35 + i * 0.15)
        d.line([(pos, s*0.12), (pos, s*0.25)], fill=(255, 200, 0), width=int(s*0.04))
        d.line([(pos, s*0.75), (pos, s*0.88)], fill=(255, 200, 0), width=int(s*0.04))
        d.line([(s*0.12, pos), (s*0.25, pos)], fill=(255, 200, 0), width=int(s*0.04))
        d.line([(s*0.75, pos), (s*0.88, pos)], fill=(255, 200, 0), width=int(s*0.04))

def _draw_benchmark(d: ImageDraw.ImageDraw, s: int) -> None:
    # Speedometer / gauge
    d.arc([s*0.15, s*0.15, s*0.85, s*0.85], start=150, end=390, fill=(24, 90, 189), width=int(s*0.08))
    d.line([(s*0.5, s*0.5), (s*0.75, s*0.32)], fill=(215, 45, 30), width=int(s*0.07))
    d.ellipse([s*0.42, s*0.42, s*0.58, s*0.58], fill=(40, 50, 60))

# --- View Tab ---
def _draw_view_normal(d: ImageDraw.ImageDraw, s: int) -> None:
    d.rectangle([s*0.15, s*0.18, s*0.85, s*0.82], outline=(33, 115, 70), width=int(s*0.05), fill=(255, 255, 255))
    d.line([(s*0.5, s*0.18), (s*0.5, s*0.82)], fill=(180, 190, 200), width=int(s*0.03))
    d.line([(s*0.15, s*0.5), (s*0.85, s*0.5)], fill=(180, 190, 200), width=int(s*0.03))

def _draw_zoom_in(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_find(d, s)
    d.line([(s*0.32, s*0.4), (s*0.48, s*0.4)], fill=(33, 115, 70), width=int(s*0.06))
    d.line([(s*0.4, s*0.32), (s*0.4, s*0.48)], fill=(33, 115, 70), width=int(s*0.06))

def _draw_zoom_out(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_find(d, s)
    d.line([(s*0.32, s*0.4), (s*0.48, s*0.4)], fill=(215, 45, 30), width=int(s*0.06))

def _draw_zoom_100(d: ImageDraw.ImageDraw, s: int) -> None:
    _draw_find(d, s)
    d.line([(s*0.35, s*0.32), (s*0.35, s*0.48)], fill=(24, 90, 189), width=int(s*0.05))

def _draw_theme(d: ImageDraw.ImageDraw, s: int) -> None:
    # Paint palette
    d.ellipse([s*0.15, s*0.18, s*0.85, s*0.82], fill=(240, 230, 210), outline=(180, 160, 130), width=int(s*0.04))
    d.ellipse([s*0.28, s*0.3, s*0.4, s*0.42], fill=(215, 45, 30))
    d.ellipse([s*0.48, s*0.25, s*0.6, s*0.37], fill=(24, 90, 189))
    d.ellipse([s*0.68, s*0.4, s*0.8, s*0.52], fill=(33, 115, 70))
    d.ellipse([s*0.55, s*0.58, s*0.67, s*0.7], fill=(255, 200, 0))

def _draw_shortcuts(d: ImageDraw.ImageDraw, s: int) -> None:
    # Keyboard
    d.rounded_rectangle([s*0.12, s*0.25, s*0.88, s*0.75], radius=s*0.06, fill=(245, 248, 252), outline=(100, 110, 120), width=int(s*0.04))
    for r in range(2):
        y = s * (0.38 + r * 0.18)
        for c in range(4):
            x = s * (0.2 + c * 0.16)
            d.rectangle([x, y, x + s*0.1, y + s*0.1], fill=(180, 190, 205))


# Map icon names to drawing functions
_ICON_DRAWERS: Dict[str, Callable[[ImageDraw.ImageDraw, int], None]] = {
    "save": _draw_save,
    "undo": _draw_undo,
    "redo": _draw_redo,
    "paste": _draw_paste,
    "cut": _draw_cut,
    "copy": _draw_copy,
    "format_painter": _draw_format_painter,
    "font_inc": _draw_font_inc,
    "font_dec": _draw_font_dec,
    "bold": _draw_bold,
    "italic": _draw_italic,
    "underline": _draw_underline,
    "strike": _draw_strike,
    "borders": _draw_borders,
    "fill_color": _draw_fill_color,
    "font_color": _draw_font_color,
    "align_top": _draw_align_top,
    "align_mid": _draw_align_mid,
    "align_bot": _draw_align_bot,
    "align_left": _draw_align_left,
    "align_center": _draw_align_center,
    "align_right": _draw_align_right,
    "wrap_text": _draw_wrap_text,
    "merge_center": _draw_merge_center,
    "currency_rub": _draw_currency_rub,
    "currency_usd": _draw_currency_usd,
    "percent": _draw_percent,
    "comma": _draw_comma,
    "decimal_inc": _draw_decimal_inc,
    "decimal_dec": _draw_decimal_dec,
    "cond_format": _draw_cond_format,
    "format_cells": _draw_format_cells,
    "insert": _draw_insert,
    "delete": _draw_delete,
    "autofit": _draw_autofit,
    "autosum": _draw_autosum,
    "clear": _draw_clear,
    "sort_az": _draw_sort_az,
    "sort_za": _draw_sort_za,
    "sort_custom": _draw_sort_custom,
    "filter": _draw_filter,
    "clear_filter": _draw_clear_filter,
    "find": _draw_find,
    "replace": _draw_replace,
    "goto": _draw_goto,
    "demo_table": _draw_demo_table,
    "new_sheet": _draw_new_sheet,
    "chart_wizard": _draw_chart_wizard,
    "chart_col": _draw_chart_col,
    "chart_line": _draw_chart_line,
    "chart_pie": _draw_chart_pie,
    "chart_bar": _draw_chart_bar,
    "chart_area": _draw_chart_area,
    "chart_scatter": _draw_chart_scatter,
    "comment_add": _draw_comment_add,
    "comment_del": _draw_comment_del,
    "fx": _draw_fx,
    "formula_ref": _draw_formula_ref,
    "import_csv": _draw_import_csv,
    "export_csv": _draw_export_csv,
    "refresh": _draw_refresh,
    "text_to_columns": _draw_text_to_columns,
    "remove_duplicates": _draw_remove_duplicates,
    "goal_seek": _draw_goal_seek,
    "freeze_panes": _draw_freeze_panes,
    "fx_financial": _draw_fx_financial,
    "fx_logical": _draw_fx_logical,
    "fx_text": _draw_fx_text,
    "fx_date": _draw_fx_date,
    "fx_lookup": _draw_fx_lookup,
    "fx_math": _draw_fx_math,
    "fx_stats": _draw_fx_stats,
    "show_formulas": _draw_show_formulas,
    "auditor": _draw_auditor,
    "calculate": _draw_calculate,
    "simd": _draw_simd,
    "benchmark": _draw_benchmark,
    "view_normal": _draw_view_normal,
    "zoom_in": _draw_zoom_in,
    "zoom_out": _draw_zoom_out,
    "zoom_100": _draw_zoom_100,
    "theme": _draw_theme,
    "shortcuts": _draw_shortcuts,
}


def get_pil_image(name: str, size: int = 16) -> Image.Image:
    """Get or generate a PIL RGBA Image for the requested icon."""
    key = (name, size)
    if key in _pil_cache:
        return _pil_cache[key]

    drawer = _ICON_DRAWERS.get(name)
    if drawer is None:
        # Fallback placeholder icon
        def fallback(d: ImageDraw.ImageDraw, s: int) -> None:
            d.rectangle([s*0.2, s*0.2, s*0.8, s*0.8], outline=(150, 160, 170), width=int(s*0.05), fill=(240, 245, 250))
            d.line([(s*0.2, s*0.2), (s*0.8, s*0.8)], fill=(180, 190, 200), width=int(s*0.04))
        drawer = fallback

    img = _supersample(drawer, size)
    _pil_cache[key] = img
    return img


def get_icon(name: str, size: int = 16) -> ImageTk.PhotoImage:
    """Get or generate a cached Tkinter PhotoImage for the requested icon."""
    key = (name, size)
    if key in _photo_cache:
        return _photo_cache[key]

    pil_img = get_pil_image(name, size)
    photo = ImageTk.PhotoImage(pil_img)
    _photo_cache[key] = photo
    return photo
