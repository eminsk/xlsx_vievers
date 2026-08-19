"""
Генератор иконки для Excel Viewer Pro.
Создает красивую иконку в стиле Excel (зеленый лист с сеткой).

Запуск: python generate_icon.py
Требуется: pip install Pillow
"""

import sys

try:
    from PIL import Image, ImageDraw

    def create_excel_icon():
        """Создает иконку Excel в различных размерах."""
        sizes = [16, 32, 48, 64, 128, 256]
        ico_images = []

        for size in sizes:
            # Создаем изображение с прозрачным фоном
            img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)

            # Фон - зеленый прямоугольник (Excel зеленый #107C41)
            margin = max(2, size // 16)
            draw.rectangle(
                [margin, margin, size - margin, size - margin],
                fill=(16, 124, 65, 255)
            )

            # Белый лист внутри
            sheet_margin = max(4, size // 8)
            draw.rectangle(
                [sheet_margin, sheet_margin, size - sheet_margin, size - sheet_margin],
                fill=(255, 255, 255, 255)
            )

            # Сетка на листе (горизонтальные линии)
            line_spacing = max(3, size // 8)
            line_width = max(1, size // 64)
            for i in range(2, 6):
                y = sheet_margin + (i * line_spacing)
                if y < size - sheet_margin - line_width:
                    draw.line(
                        [sheet_margin + line_width * 2, y, size - sheet_margin - line_width * 2, y],
                        fill=(220, 220, 220, 255), width=line_width
                    )

            # Сетка на листе (вертикальные линии)
            for i in range(2, 5):
                x = sheet_margin + int(i * line_spacing * 1.5)
                if x < size - sheet_margin - line_width:
                    draw.line(
                        [x, sheet_margin + line_width * 2, x, size - sheet_margin - line_width * 2],
                        fill=(220, 220, 220, 255), width=line_width
                    )

            ico_images.append(img)

        # Сохраняем PNG версию (256x256)
        ico_images[-1].save('app_icon.png', 'PNG')
        # Сохраняем ICO файл с несколькими размерами
        ico_images[0].save('app_icon.ico', format='ICO', sizes=[(s, s) for s in sizes])

        print("[OK] Icons created: app_icon.png and app_icon.ico")

    if __name__ == "__main__":
        create_excel_icon()

except ImportError:
    print("ERROR: Pillow is not installed. Run: pip install Pillow")
