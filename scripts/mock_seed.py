# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | 레퍼런스가 없을 때 쓰는 16장 Mock 시드 덱 생성기
=============================================================================
"""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

PRIMARY = RGBColor(0x1B, 0x3B, 0x2B)
SECONDARY = RGBColor(0x3E, 0x7B, 0x54)
ACCENT = RGBColor(0xE5, 0xA9, 0x3C)
TEXT = RGBColor(0x22, 0x22, 0x22)
LIGHT = RGBColor(0xF8, 0xFA, 0xF8)

SECTIONS = [
    "표지", "목차", "문제 정의", "시장 기회", "해결 방안", "핵심 기술",
    "비즈니스 모델", "경쟁 분석", "추진 전략", "실행 일정", "조직 구성",
    "재무 계획", "성과 지표", "리스크 관리", "기대 효과", "마무리",
]


def _textbox(slide, x, y, w, h, text, size, color, bold=False):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def build_mock_seed(out_path, topic):
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]

    for idx, section in enumerate(SECTIONS, 1):
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = PRIMARY if idx in (1, len(SECTIONS)) else LIGHT

        if idx == 1:
            _textbox(slide, 914400, 2286000, W - 1828800, 1371600, topic, 40, LIGHT, bold=True)
            _textbox(slide, 914400, 3657600, W - 1828800, 609600, "부제목을 입력하세요", 18, ACCENT)
            continue
        if idx == len(SECTIONS):
            _textbox(slide, 914400, 2743200, W - 1828800, 1371600, "감사합니다", 40, LIGHT, bold=True)
            continue

        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Emu(800000))
        bar.fill.solid()
        bar.fill.fore_color.rgb = PRIMARY
        bar.line.fill.background()
        _textbox(slide, 457200, 160000, W - 2743200, 560000, f"{idx:02d}. {section}", 26, LIGHT, bold=True)
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, W - Emu(1600000), Emu(200000), Emu(1300000), Emu(420000))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT
        badge.line.fill.background()
        badge.text_frame.text = f"Step {idx - 1:02d}"

        card_w = (W - Emu(914400) * 2 - Emu(304800) * 2) // 3
        for c in range(3):
            x = Emu(914400) + c * (card_w + Emu(304800))
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Emu(1371600), card_w, Emu(4572000))
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            card.line.color.rgb = SECONDARY
            _textbox(slide, x + 152400, 1524000, card_w - 304800, 3962400,
                     f"{section} 핵심 {c + 1}\n내용을 입력하세요", 15, TEXT)

    prs.core_properties.title = topic
    prs.core_properties.author = "slide-style-cloner mock seed"
    prs.save(out_path)
