"""
High-Quality Multi-Resolution Icon Generator for Excel Viewer Pro.
Generates modern Excel-style 32-bit RGBA icon (.ico and .png) with multi-resolution support:
256x256, 128x128, 64x64, 48x48, 32x32, 24x24, 16x16.
"""

from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter


def draw_rounded_rect(draw: ImageDraw.ImageDraw, bounds: tuple[float, float, float, float], radius: float, fill, outline=None, width: int = 1):
    x0, y0, x1, y1 = bounds
    if x1 <= x0 or y1 <= y0:
        return
    radius = min(radius, (x1 - x0) / 2, (y1 - y0) / 2)
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill, outline=outline, width=width)


def create_excel_icon_image(size: int) -> Image.Image:
    """Render a modern Excel spreadsheet icon at high resolution."""
    # Always render on a fixed 1024x1024 canvas and downsample with Lanczos
    canvas_size = 1024
    img = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))

    w = canvas_size
    h = canvas_size

    # 1. Back Sheet (Workbook Page)
    sheet_left = int(w * 0.36)
    sheet_top = int(h * 0.08)
    sheet_right = int(w * 0.94)
    sheet_bottom = int(h * 0.92)
    sheet_radius = 48

    # Soft drop shadow for sheet
    shadow_img = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    s_draw = ImageDraw.Draw(shadow_img)
    draw_rounded_rect(s_draw, (sheet_left + 16, sheet_top + 24, sheet_right + 16, sheet_bottom + 24),
                      radius=sheet_radius, fill=(0, 0, 0, 80))
    shadow_img = shadow_img.filter(ImageFilter.GaussianBlur(radius=24))
    img.alpha_composite(shadow_img)

    # Sheet background - clean light gradient
    sheet_img = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    sh_draw = ImageDraw.Draw(sheet_img)
    draw_rounded_rect(sh_draw, (sheet_left, sheet_top, sheet_right, sheet_bottom),
                      radius=sheet_radius, fill=(250, 252, 255, 255), outline=(203, 213, 225, 255), width=8)

    # Header bar on sheet (Excel dark green #107C41)
    header_h = int((sheet_bottom - sheet_top) * 0.22)
    # Clip header to top rounded corners
    header_mask = Image.new("L", (canvas_size, canvas_size), 0)
    hm_draw = ImageDraw.Draw(header_mask)
    draw_rounded_rect(hm_draw, (sheet_left, sheet_top, sheet_right, sheet_bottom), radius=sheet_radius, fill=255)
    
    header_layer = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    hl_draw = ImageDraw.Draw(header_layer)
    hl_draw.rectangle([sheet_left, sheet_top, sheet_right, sheet_top + header_h], fill=(16, 124, 65, 255))
    header_layer.putalpha(header_mask)
    sheet_img.alpha_composite(header_layer)

    # Gridlines on sheet (horizontal and vertical cells)
    grid_y_start = sheet_top + header_h
    num_rows = 4
    row_h = (sheet_bottom - grid_y_start) / num_rows
    for r in range(1, num_rows):
        gy = int(grid_y_start + r * row_h)
        sh_draw.line([(sheet_left + 16, gy), (sheet_right - 16, gy)], fill=(226, 232, 240, 255), width=8)

    num_cols = 3
    col_w = (sheet_right - sheet_left) / num_cols
    for c in range(1, num_cols):
        gx = int(sheet_left + c * col_w)
        sh_draw.line([(gx, grid_y_start), (gx, sheet_bottom - 16)], fill=(226, 232, 240, 255), width=8)

    # Selected active cell accent
    sh_draw.rectangle([int(sheet_left + col_w) + 8, int(grid_y_start + row_h) + 8,
                       int(sheet_left + 2 * col_w) - 8, int(grid_y_start + 2 * row_h) - 8],
                      fill=(204, 224, 248, 255)) # Soft blue #CCE0F8

    img.alpha_composite(sheet_img)

    # 2. Front Tile (The Excel Green Badge with 'X')
    badge_w = int(w * 0.52)
    badge_h = int(h * 0.68)
    badge_left = int(w * 0.06)
    badge_top = int((h - badge_h) / 2)
    badge_right = badge_left + badge_w
    badge_bottom = badge_top + badge_h
    badge_radius = 64

    # Badge Drop Shadow
    b_shadow = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    bs_draw = ImageDraw.Draw(b_shadow)
    draw_rounded_rect(bs_draw, (badge_left + 20, badge_top + 28, badge_right + 20, badge_bottom + 28),
                      radius=badge_radius, fill=(0, 0, 0, 130))
    b_shadow = b_shadow.filter(ImageFilter.GaussianBlur(radius=32))
    img.alpha_composite(b_shadow)

    # Badge Body: Rich Green Gradient (top-left #1E8E4D, bottom-right #0D5C2E)
    badge_img = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    b_draw = ImageDraw.Draw(badge_img)

    for y in range(badge_top, badge_bottom):
        t = (y - badge_top) / max(1, badge_bottom - badge_top)
        r = int(32 * (1 - t) + 10 * t)
        g = int(148 * (1 - t) + 90 * t)
        b = int(72 * (1 - t) + 40 * t)
        b_draw.line([(badge_left, y), (badge_right, y)], fill=(r, g, b, 255))

    # Mask to rounded rectangle
    mask = Image.new("L", (canvas_size, canvas_size), 0)
    m_draw = ImageDraw.Draw(mask)
    draw_rounded_rect(m_draw, (badge_left, badge_top, badge_right, badge_bottom), radius=badge_radius, fill=255)
    badge_img.putalpha(mask)

    # Badge highlight border (subtle top highlight)
    b_border = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    bb_draw = ImageDraw.Draw(b_border)
    draw_rounded_rect(bb_draw, (badge_left, badge_top, badge_right, badge_bottom),
                      radius=badge_radius, fill=None, outline=(255, 255, 255, 90), width=8)
    badge_img.alpha_composite(b_border)

    # 3. Bold White Letter 'X' on Badge
    cx = (badge_left + badge_right) / 2
    cy = (badge_top + badge_bottom) / 2
    x_span = int(badge_w * 0.30)
    x_thickness = int(badge_w * 0.16)

    # Draw sleek, perfectly proportioned stylized 'X'
    x_layer = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    x_draw = ImageDraw.Draw(x_layer)

    p1 = (cx - x_span, cy - x_span * 1.15)
    p2 = (cx + x_span, cy + x_span * 1.15)
    p3 = (cx + x_span, cy - x_span * 1.15)
    p4 = (cx - x_span, cy + x_span * 1.15)

    # Soft shadow under X
    x_draw.line([p1, p2], fill=(0, 0, 0, 70), width=x_thickness + 10)
    x_draw.line([p3, p4], fill=(0, 0, 0, 70), width=x_thickness + 10)
    x_layer = x_layer.filter(ImageFilter.GaussianBlur(radius=8))

    # Crisp white X
    x_fg = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    xf_draw = ImageDraw.Draw(x_fg)
    xf_draw.line([p1, p2], fill=(255, 255, 255, 255), width=x_thickness)
    xf_draw.line([p3, p4], fill=(255, 255, 255, 255), width=x_thickness)

    # Composite badge and X
    badge_img.alpha_composite(x_layer)
    badge_img.alpha_composite(x_fg)
    img.alpha_composite(badge_img)

    # Downsample with high-quality Lanczos filter
    final_img = img.resize((size, size), Image.Resampling.LANCZOS)
    return final_img


def generate_all_icons():
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = {}
    for s in sizes:
        images[s] = create_excel_icon_image(s)

    out_dir = Path(__file__).parent.resolve()

    # Save PNG (256x256)
    png_path = out_dir / "app_icon.png"
    images[256].save(png_path, "PNG", optimize=True)

    # Save multi-resolution ICO (256 down to 16)
    ico_path = out_dir / "app_icon.ico"
    img_256 = images[256]
    other_imgs = [images[s] for s in [128, 64, 48, 32, 24, 16]]
    img_256.save(ico_path, format="ICO", append_images=other_imgs, sizes=[(s, s) for s in sizes])

    # Also copy to xlsx_viewer package directory
    pkg_dir = out_dir / "xlsx_viewer"
    if pkg_dir.exists():
        images[256].save(pkg_dir / "app_icon.png", "PNG", optimize=True)
        img_256.save(pkg_dir / "app_icon.ico", format="ICO", append_images=other_imgs, sizes=[(s, s) for s in sizes])

    # Also copy to asm directory
    asm_dir = out_dir / "asm"
    if asm_dir.exists():
        images[256].save(asm_dir / "app_icon.png", "PNG", optimize=True)
        img_256.save(asm_dir / "app_icon.ico", format="ICO", append_images=other_imgs, sizes=[(s, s) for s in sizes])

    print(f"[OK] Generated {ico_path} ({ico_path.stat().st_size:,} bytes) and {png_path}")


if __name__ == "__main__":
    generate_all_icons()
