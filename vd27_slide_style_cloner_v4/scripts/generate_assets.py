# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | 테마 색으로 주제 엠블럼/성과 배지 PNG 생성 + 교체 이미지 맞춤 자르기 함수
=============================================================================
"""

import os
import glob

from PIL import Image, ImageDraw, ImageFont, ImageOps

import cloner_common as cc

FONT_CANDIDATES = [
    "C:/Windows/Fonts/malgunbd.ttf", "C:/Windows/Fonts/malgun.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
]


def load_font(size):
    for path in FONT_CANDIDATES:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    for path in glob.glob("C:/Windows/Fonts/*.ttf"):
        if "gothic" in path.lower() or "malgun" in path.lower():
            return ImageFont.truetype(path, size)
    print("[!] 한글 글꼴을 찾지 못해 기본 글꼴로 그립니다. 한글이 네모로 보일 수 있습니다.")
    return ImageFont.load_default()


def hex_rgb(value, default):
    try:
        v = value.lstrip("#")
        return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return default


def palette(pillars):
    """테마에서 주색/강조색을 고른다. 흑백뿐이면 기본 녹색·금색."""
    theme = pillars.get("theme", {}).get("colors", {}) if pillars else {}
    used = [c for c, _ in pillars.get("style", {}).get("colors", [])] if pillars else []
    chromatic = [c for c in used if c not in ("#000000", "#FFFFFF")]
    primary = hex_rgb(theme.get("dk2") or (chromatic[0] if chromatic else ""), (27, 59, 43))
    accent = hex_rgb(theme.get("accent4") or (chromatic[1] if len(chromatic) > 1 else ""), (229, 169, 60))
    return primary, accent


def _fit_text(draw, text, max_w, start):
    size = start
    while size > 10:
        font = load_font(size)
        if draw.textlength(text, font=font) <= max_w:
            return font
        size -= 2
    return load_font(10)


def make_emblem(path, text, primary, accent, px=800):
    img = Image.new("RGBA", (px, px), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, px - 8, px - 8), fill=primary + (255,), outline=accent + (255,), width=px // 40)
    d.ellipse((px * 0.1, px * 0.1, px * 0.9, px * 0.9), outline=accent + (200,), width=px // 120)
    label = text.strip()[:6] or "LOGO"
    font = _fit_text(d, label, px * 0.66, px // 4)
    d.text((px / 2, px / 2), label, font=font, fill=(255, 255, 255, 255), anchor="mm")
    img.save(path)


def make_badge(path, text, accent, w=900, h=300):
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((4, 4, w - 4, h - 4), radius=h // 2, fill=accent + (255,))
    font = _fit_text(d, text, w * 0.8, h // 2)
    d.text((w / 2, h / 2), text, font=font, fill=(255, 255, 255, 255), anchor="mm")
    img.save(path)


def fit_replacement(new_path, target_size, fmt):
    """
    교체 이미지를 원본 이미지와 같은 픽셀 크기로 가운데 맞춰 자른다 (비율 유지 cover crop).
    PPT 안의 위치·크기·자르기 값이 원본 픽셀 비율을 전제로 하므로, 같은 크기로 맞춰야 화면이 그대로 유지된다.
    """
    import io
    img = Image.open(new_path)
    img = ImageOps.exif_transpose(img)
    fitted = ImageOps.fit(img, target_size, method=Image.LANCZOS)
    fmt = (fmt or "PNG").upper().replace("JPG", "JPEG")
    if fmt == "JPEG" and fitted.mode not in ("RGB", "L"):
        fitted = fitted.convert("RGB")
    buf = io.BytesIO()
    fitted.save(buf, format=fmt, quality=92) if fmt == "JPEG" else fitted.save(buf, format=fmt)
    return buf.getvalue()


def main():
    cc.setup_stdout()
    args = cc.base_parser("도메인 맞춤 엠블럼/배지 에셋 생성").parse_args()
    state = cc.load_state(args.output)
    pillars_path = os.path.join(args.output, cc.F_PILLARS_JSON)
    pillars = cc.read_json(pillars_path) if os.path.exists(pillars_path) else {}
    topic = args.topic or state.get("topic") or state.get("deck_title") or "PROJECT"

    out_dir = os.path.join(args.output, "assets")
    os.makedirs(out_dir, exist_ok=True)
    primary, accent = palette(pillars)
    emblem = os.path.join(out_dir, "emblem_topic.png")
    badge = os.path.join(out_dir, "badge_topic.png")
    make_emblem(emblem, topic.replace(" ", "")[:4], primary, accent)
    make_badge(badge, topic[:14], accent)
    cc.save_state(args.output, assets={"emblem": os.path.abspath(emblem), "badge": os.path.abspath(badge)})
    print(f"[+] 에셋 생성: {emblem}, {badge}")


if __name__ == "__main__":
    main()
