from __future__ import annotations

import importlib
from pathlib import Path

WIDTH = 2560
HEIGHT = 640
BACKGROUND = "#08111F"
TITLE = "outlook-mail-assistant"
SUBTITLE = "Local-first Outlook mail intelligence for Windows PCs"
TAGLINE = "Ingest mail. Extract actions. Keep runtime workspaces outside the repo."
OUTPUT = Path(__file__).with_name("cover.png")


def _pil():
    image = importlib.import_module("PIL.Image")
    image_chops = importlib.import_module("PIL.ImageChops")
    image_draw = importlib.import_module("PIL.ImageDraw")
    image_filter = importlib.import_module("PIL.ImageFilter")
    image_font = importlib.import_module("PIL.ImageFont")
    return image, image_chops, image_draw, image_filter, image_font


def _load_font(size: int, *, bold: bool = False):
    _, _, _, _, image_font = _pil()
    candidates = []
    if bold:
        candidates.extend(
            [
                "C:/Windows/Fonts/consolab.ttf",
                "C:/Windows/Fonts/Consolas.ttf",
                "C:/Windows/Fonts/lucon.ttf",
                "C:/Windows/Fonts/courbd.ttf",
            ]
        )
    else:
        candidates.extend(
            [
                "C:/Windows/Fonts/consola.ttf",
                "C:/Windows/Fonts/Consolas.ttf",
                "C:/Windows/Fonts/lucon.ttf",
                "C:/Windows/Fonts/cour.ttf",
            ]
        )
    for candidate in candidates:
        if Path(candidate).exists():
            return image_font.truetype(candidate, size=size)
    return image_font.load_default()


def _ellipse_blob(
    size: tuple[int, int],
    bbox: tuple[int, int, int, int],
    color: tuple[int, int, int, int],
    blur: int,
):
    image, _, image_draw, image_filter, _ = _pil()
    layer = image.new("RGBA", size, (0, 0, 0, 0))
    draw = image_draw.Draw(layer)
    draw.ellipse(bbox, fill=color)
    return layer.filter(image_filter.GaussianBlur(blur))


def _add_film_grain(base_image, opacity: int = 30):
    image, _, _, image_filter, _ = _pil()
    grain = image.effect_noise((WIDTH, HEIGHT), 18).convert("RGBA")
    grain.putalpha(opacity)
    grain = grain.filter(image_filter.GaussianBlur(0.3))
    return image.alpha_composite(base_image, grain)


def _draw_glow_text(
    base,
    position: tuple[int, int],
    text: str,
    font,
    *,
    fill: str,
    glow: tuple[int, int, int],
    blur: int = 18,
    anchor: str | None = None,
) -> None:
    image, _, image_draw, image_filter, _ = _pil()
    glow_layer = image.new("RGBA", base.size, (0, 0, 0, 0))
    glow_draw = image_draw.Draw(glow_layer)
    glow_draw.text(position, text, font=font, fill=(*glow, 230), anchor=anchor)
    glow_layer = glow_layer.filter(image_filter.GaussianBlur(blur))
    base.alpha_composite(glow_layer)
    draw = image_draw.Draw(base)
    draw.text(position, text, font=font, fill=fill, anchor=anchor)


def build_cover():
    image, image_chops, image_draw, image_filter, _ = _pil()
    canvas = image.new("RGBA", (WIDTH, HEIGHT), BACKGROUND)

    blobs = [
        ((-120, 120, 1120, 1120), (40, 106, 255, 150), 120),
        ((1180, 40, 2440, 1060), (0, 219, 222, 120), 140),
        ((760, 320, 1820, 1220), (106, 27, 154, 110), 130),
        ((1580, 640, 2520, 1320), (79, 70, 229, 95), 110),
    ]
    for bbox, color, blur in blobs:
        canvas.alpha_composite(_ellipse_blob((WIDTH, HEIGHT), bbox, color, blur))

    vignette = image.new("L", (WIDTH, HEIGHT), 0)
    vignette_draw = image_draw.Draw(vignette)
    vignette_draw.ellipse((-220, -140, WIDTH + 220, HEIGHT + 160), fill=255)
    vignette = image_chops.invert(vignette).filter(image_filter.GaussianBlur(120))
    vignette_rgba = image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    vignette_rgba.putalpha(vignette)
    canvas = image.alpha_composite(canvas, vignette_rgba)

    canvas = _add_film_grain(canvas)

    title_font = _load_font(132, bold=True)
    subtitle_font = _load_font(38)
    tagline_font = _load_font(28)

    draw = image_draw.Draw(canvas)

    center_x = WIDTH // 2
    title_y = 280
    subtitle_y = title_y + 170
    tagline_y = subtitle_y + 78

    _draw_glow_text(
        canvas,
        (center_x, title_y),
        TITLE,
        title_font,
        fill="#F8FBFF",
        glow=(82, 168, 255),
        anchor="mm",
    )

    draw.text(
        (center_x, subtitle_y),
        SUBTITLE,
        font=subtitle_font,
        fill="#D4E4FF",
        anchor="mm",
    )
    draw.text(
        (center_x, tagline_y), TAGLINE, font=tagline_font, fill="#9FB6D9", anchor="mm"
    )

    line_y = tagline_y + 64
    line_width = 520
    line_x = (WIDTH - line_width) // 2
    draw.rounded_rectangle(
        (line_x, line_y, line_x + line_width, line_y + 8), radius=4, fill="#5CC8FF"
    )

    chip_font = _load_font(26, bold=True)
    chips = ["OUTLOOK DESKTOP", "JSONL + SQLITE", "TASK REPORTS", "DRY-RUN FIRST"]
    chip_y = 138
    chip_widths = []
    for chip in chips:
        text_box = draw.textbbox((0, 0), chip, font=chip_font)
        chip_widths.append(text_box[2] - text_box[0] + 42)
    total_chip_width = sum(chip_widths) + (18 * (len(chips) - 1))
    chip_x = (WIDTH - total_chip_width) // 2
    for chip, chip_w in zip(chips, chip_widths):
        draw.rounded_rectangle(
            (chip_x, chip_y, chip_x + chip_w, chip_y + 46),
            radius=18,
            fill=(10, 20, 37, 180),
            outline="#5CC8FF",
            width=2,
        )
        draw.text((chip_x + 21, chip_y + 10), chip, font=chip_font, fill="#DDF4FF")
        chip_x += chip_w + 18

    border = image_draw.Draw(canvas)
    border.rounded_rectangle(
        (18, 18, WIDTH - 18, HEIGHT - 18),
        radius=54,
        outline=(255, 255, 255, 52),
        width=2,
    )

    mask = image.new("L", (WIDTH, HEIGHT), 0)
    image_draw.Draw(mask).rounded_rectangle((0, 0, WIDTH, HEIGHT), radius=60, fill=255)
    rounded = image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    rounded.paste(canvas, (0, 0), mask)
    return rounded


def main() -> None:
    image = build_cover()
    image.save(OUTPUT, format="PNG", optimize=True)
    print(f"wrote {OUTPUT}")


if __name__ == "__main__":
    main()
