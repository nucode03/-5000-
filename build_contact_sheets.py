from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "build" / "pdf_pages"
DEST = ROOT / "build" / "contact_sheets"
DEST.mkdir(exist_ok=True)

pages = sorted(SOURCE.glob("page-*.png"))
for sheet_number, offset in enumerate(range(0, len(pages), 12), start=1):
    chunk = pages[offset:offset + 12]
    thumb_w, thumb_h = 300, 424
    canvas = Image.new("RGB", (thumb_w * 3, (thumb_h + 25) * 4), "#f4f6f8")
    draw = ImageDraw.Draw(canvas)
    for local_index, path in enumerate(chunk):
        image = Image.open(path).convert("RGB")
        image.thumbnail((thumb_w, thumb_h))
        x = (local_index % 3) * thumb_w + (thumb_w - image.width) // 2
        y = (local_index // 3) * (thumb_h + 25)
        canvas.paste(image, (x, y + 20))
        draw.text((x + 8, y + 3), f"Page {offset + local_index + 1}", fill="#163A5F")
    canvas.save(DEST / f"sheet-{sheet_number:02d}.png")

print(f"Created {len(list(DEST.glob('sheet-*.png')))} contact sheets")
