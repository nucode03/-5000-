"""Generate PNG icons for the Oxford 5000 PWA."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ICON_DIR = ROOT / "web_app" / "icons"


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for candidate in ("arialbd.ttf", "arial.ttf", "malgunbd.ttf", "malgun.ttf"):
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    return ImageFont.load_default()


def make_icon(size: int) -> None:
    image = Image.new("RGB", (size, size), "#163A5F")
    draw = ImageDraw.Draw(image)
    margin = round(size * 0.14)
    card = [margin, margin, size - margin, size - margin]
    draw.rounded_rectangle(card, radius=round(size * 0.07), fill="#EDF4FB")
    draw.rectangle(
        [margin, margin, size - margin, margin + round(size * 0.18)],
        fill="#2E74B5",
    )

    title_font = load_font(round(size * 0.18))
    sub_font = load_font(round(size * 0.09))
    draw.text((size * 0.5, size * 0.43), "A-Z", anchor="mm", fill="#163A5F", font=title_font)
    draw.text((size * 0.5, size * 0.60), "5000", anchor="mm", fill="#2E74B5", font=sub_font)
    draw.text((size * 0.5, size * 0.74), "KO", anchor="mm", fill="#163A5F", font=sub_font)
    image.save(ICON_DIR / f"icon-{size}.png")


def main() -> None:
    ICON_DIR.mkdir(parents=True, exist_ok=True)
    for size in (192, 512):
        make_icon(size)
    print("Generated PWA icons.")


if __name__ == "__main__":
    main()
