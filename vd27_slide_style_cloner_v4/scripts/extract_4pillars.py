# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | [Step 1] 레퍼런스 PPTX를 실제로 분석해 4대 설계도서와 pillars.json 생성
=============================================================================
"""

import os
import re
import zipfile
import posixpath
from collections import Counter

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import cloner_common as cc

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A = "{%s}" % NS["a"]

ARCHETYPE_LABELS = {
    "cover": "표지",
    "closing": "엔딩",
    "table": "표/일정",
    "process": "프로세스 흐름도",
    "gallery": "이미지 갤러리",
    "cards": "카드형 목록",
    "text": "텍스트 중심",
    "media": "동영상",
    "content": "일반 내용",
}

STOPWORDS = {"그리고", "또는", "있는", "하는", "위한", "대한", "통해", "및", "등", "the", "and", "for"}


# ---------------------------------------------------------------------------
# 테마 (색 / 글꼴)
# ---------------------------------------------------------------------------
def read_theme(pptx_path):
    with zipfile.ZipFile(pptx_path) as z:
        names = z.namelist()
        theme_name = None
        master_rels = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
        if master_rels in names:
            root = etree.fromstring(z.read(master_rels))
            for rel in root:
                if rel.get("Type", "").endswith("/theme"):
                    theme_name = posixpath.normpath(posixpath.join("ppt/slideMasters", rel.get("Target")))
        if theme_name not in names:
            themes = sorted(n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n))
            theme_name = themes[0] if themes else None
        if not theme_name:
            return {"colors": {}, "major_font": {}, "minor_font": {}, "name": ""}
        root = etree.fromstring(z.read(theme_name))

    colors = {}
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is not None:
        for el in scheme:
            key = etree.QName(el).localname
            child = el[0] if len(el) else None
            if child is None:
                continue
            val = child.get("val") if etree.QName(child).localname == "srgbClr" else child.get("lastClr")
            if val:
                colors[key] = "#" + val.upper()

    def font(tag):
        el = root.find(f".//a:fontScheme/a:{tag}", NS)
        if el is None:
            return {}
        out = {}
        for script in ("latin", "ea", "cs"):
            s = el.find(f"a:{script}", NS)
            if s is not None and s.get("typeface"):
                out[script] = s.get("typeface")
        return out

    return {
        "name": root.get("name", ""),
        "colors": colors,
        "major_font": font("majorFont"),
        "minor_font": font("minorFont"),
    }


# ---------------------------------------------------------------------------
# 슬라이드 분석
# ---------------------------------------------------------------------------
def classify(slide_no, total, info, slide_w, slide_h):
    if slide_no == 1:
        return "cover"
    if slide_no == total and total > 2:
        return "closing"
    if info["media_count"]:
        return "media"
    if info["tables"]:
        return "table"
    if info["max_group_size"] >= 6:
        return "process"
    if info["pictures"] >= 3:
        return "gallery"
    if info["card_like"] >= 3:
        return "cards"
    if info["chars"] >= 300:
        return "text"
    return "content"


def analyze_slide(slide, slide_no, total, slide_w, slide_h, repeated):
    info = {
        "no": slide_no,
        "layout": slide.slide_layout.name,
        "title": cc.pick_title(slide, repeated, slide_h, is_cover=(slide_no == 1)),
        "shapes_count": 0,
        "pictures": 0,
        "groups": 0,
        "tables": [],
        "media_count": 0,
        "chars": 0,
        "max_group_size": 0,
        "card_like": 0,
        "text_slots": [],
        "pictures_detail": [],
        "lines": [],
    }
    sizes = Counter()
    for shape, depth, parent in cc.iter_shapes(slide.shapes):
        info["shapes_count"] += 1
        st = shape.shape_type
        if st == MSO_SHAPE_TYPE.PICTURE:
            info["pictures"] += 1
            info["pictures_detail"].append({
                "id": shape.shape_id, "name": shape.name,
                "left": shape.left, "top": shape.top, "width": shape.width, "height": shape.height,
                "image": _picture_target(shape),
            })
        elif st == MSO_SHAPE_TYPE.GROUP:
            info["groups"] += 1
            info["max_group_size"] = max(info["max_group_size"], len(shape.shapes))
        elif st == MSO_SHAPE_TYPE.MEDIA:
            info["media_count"] += 1
        if getattr(shape, "has_table", False) and shape.has_table:
            info["tables"].append({"id": shape.shape_id, "name": shape.name, "rows": cc.table_text(shape)})
        text = cc.shape_text(shape)
        if text:
            info["chars"] += len(text)
            if shape.width and shape.height:
                sizes[(round(shape.width, -5), round(shape.height, -5))] += 1
            info["text_slots"].append({
                "id": cc.slot_id(slide_no, shape),
                "shape_id": shape.shape_id,
                "name": shape.name,
                "type": cc.shape_type_name(shape),
                "in_group": parent.name if parent is not None else None,
                "left": shape.left, "top": shape.top, "width": shape.width, "height": shape.height,
                "font_pt": cc.max_font_pt(shape),
                "repeated": text in repeated,
                "text": text,
            })
            info["lines"].extend(l.strip() for l in text.split("\n") if l.strip())
    info["card_like"] = max(sizes.values()) if sizes else 0
    info["archetype"] = classify(slide_no, total, info, slide_w, slide_h)
    return info


def _picture_target(shape):
    try:
        return shape.image.filename or ""
    except Exception:
        return ""


def style_stats(prs):
    """슬라이드 XML 전체에서 실제 쓰인 색, 글꼴, 글자 크기, 도형, 선, 그림자 통계."""
    colors, scheme_colors, fonts, sizes, geoms, line_w = Counter(), Counter(), Counter(), Counter(), Counter(), Counter()
    shadows = soft_edges = 0
    round_adj = Counter()
    for slide in prs.slides:
        root = slide._element
        for el in root.iter(A + "srgbClr"):
            colors["#" + el.get("val", "").upper()] += 1
        for el in root.iter(A + "schemeClr"):
            scheme_colors[el.get("val")] += 1
        for tag in ("latin", "ea"):
            for el in root.iter(A + tag):
                tf = el.get("typeface")
                if tf and not tf.startswith("+"):
                    fonts[tf] += 1
        for el in root.iter(A + "rPr"):
            if el.get("sz"):
                sizes[int(el.get("sz")) / 100] += 1
        for el in root.iter(A + "prstGeom"):
            geoms[el.get("prst")] += 1
            if el.get("prst") == "roundRect":
                gd = el.find(".//a:gd", NS)
                if gd is not None:
                    round_adj[gd.get("fmla")] += 1
        for el in root.iter(A + "ln"):
            if el.get("w"):
                line_w[round(int(el.get("w")) / 12700, 2)] += 1
        shadows += sum(1 for _ in root.iter(A + "outerShdw"))
        soft_edges += sum(1 for _ in root.iter(A + "softEdge"))
    return {
        "colors": colors.most_common(12),
        "scheme_colors": scheme_colors.most_common(10),
        "fonts": fonts.most_common(10),
        "font_sizes": sorted(sizes.items(), key=lambda kv: -kv[0]),
        "geometries": geoms.most_common(12),
        "line_widths_pt": line_w.most_common(6),
        "round_rect_adj": round_adj.most_common(3),
        "outer_shadows": shadows,
        "soft_edges": soft_edges,
    }


def chrome_elements(prs):
    """절반 이상의 슬라이드에 같은 위치/크기로 반복되는 도형 = 공통 머리글/장식 요소."""
    counter, sample = Counter(), {}
    for slide in prs.slides:
        seen = set()
        for shape in slide.shapes:
            key = (cc.shape_type_name(shape), shape.left, shape.top, shape.width, shape.height)
            if key not in seen:
                seen.add(key)
                counter[key] += 1
                sample.setdefault(key, (shape.name, cc.shape_text(shape)))
    n = len(prs.slides)
    out = []
    for key, c in counter.most_common():
        if n >= 3 and c / n >= 0.5:
            name, text = sample[key]
            out.append({"type": key[0], "left": key[1], "top": key[2], "width": key[3], "height": key[4],
                        "name": name, "text": text, "slides": c})
    return out


def keywords(slides, top=15):
    counter = Counter()
    for s in slides:
        for line in s["lines"]:
            for w in re.findall(r"[가-힣A-Za-z]{2,}", line):
                if w.lower() not in STOPWORDS:
                    counter[w] += 1
    return counter.most_common(top)


def extract(pptx_path):
    prs = Presentation(pptx_path)
    w, h = prs.slide_width, prs.slide_height
    repeated = cc.repeated_texts(prs)
    total = len(prs.slides)
    slides = [analyze_slide(s, i, total, w, h, repeated) for i, s in enumerate(prs.slides, 1)]
    usage = cc.media_usage(pptx_path)
    return {
        "source": pptx_path,
        "file_size": os.path.getsize(pptx_path),
        "slide_width": w,
        "slide_height": h,
        "slides_count": total,
        "core": cc.core_props(prs),
        "theme": read_theme(pptx_path),
        "style": style_stats(prs),
        "repeated_texts": sorted(repeated),
        "chrome": chrome_elements(prs),
        "keywords": keywords(slides),
        "media": {k: v for k, v in sorted(usage.items(), key=lambda kv: -kv[1]["size"])},
        "slides": slides,
    }


# ---------------------------------------------------------------------------
# 문서 작성
# ---------------------------------------------------------------------------
def _short(text, n=80):
    text = " ".join(text.split())
    return text if len(text) <= n else text[: n - 1] + "…"


def _ratio(w, h):
    r = w / h
    for label, val in (("16:9", 16 / 9), ("4:3", 4 / 3), ("16:10", 1.6), ("A4 가로", 297 / 210)):
        if abs(r - val) < 0.02:
            return label
    return f"{r:.2f}:1"


def deck_title(data):
    cover = data["slides"][0] if data["slides"] else None
    if cover and cover["title"]:
        return cover["title"]
    return data["core"]["title"] or os.path.splitext(os.path.basename(data["source"]))[0]


def write_business_dna(data, path):
    s = data["slides"]
    arche = Counter(x["archetype"] for x in s)
    with open(path, "w", encoding="utf-8") as f:
        f.write("# [001] BUSINESS_DNA — 비즈니스 내러티브 분석서\n\n")
        f.write("> 레퍼런스 파일을 자동 분석한 결과입니다. 청중·목적처럼 파일만으로 알 수 없는 항목은 비워 두었으니 직접 채워 주세요.\n\n")
        f.write("## 1. 개요\n")
        f.write(f"- **원본 파일**: `{os.path.basename(data['source'])}`\n")
        f.write(f"- **덱 제목(표지)**: {deck_title(data)}\n")
        f.write(f"- **파일 속성 제목**: {data['core']['title'] or '(없음)'}\n")
        f.write(f"- **슬라이드 수**: {data['slides_count']}장\n")
        if data["repeated_texts"]:
            f.write(f"- **반복 머리글/배지 문구**: {', '.join('`' + t + '`' for t in data['repeated_texts'])}\n")
        f.write("- **대상 청중**: (직접 입력)\n- **발표 목적**: (직접 입력)\n\n")
        f.write("## 2. 핵심 키워드 (본문 빈도순)\n")
        f.write(", ".join(f"{w}({c})" for w, c in data["keywords"]) + "\n\n")
        f.write("## 3. 구성 비율\n")
        for k, c in arche.most_common():
            f.write(f"- {ARCHETYPE_LABELS[k]}: {c}장\n")
        f.write("\n## 4. 슬라이드별 내러티브 흐름\n\n")
        for x in s:
            f.write(f"### Slide {x['no']:02d} — {x['title'] or '(제목 없음)'}\n")
            f.write(f"- **역할**: {ARCHETYPE_LABELS[x['archetype']]} · 글자 {x['chars']}자 · 이미지 {x['pictures']}개 · 표 {len(x['tables'])}개\n")
            body = [l for l in x["lines"] if l != x["title"] and l not in data["repeated_texts"]]
            if body:
                f.write("- **주요 메시지**:\n")
                for line in body[:3]:
                    f.write(f"  - {_short(line)}\n")
            f.write("\n")


def write_design(data, path):
    th, st = data["theme"], data["style"]
    w, h = data["slide_width"], data["slide_height"]
    with open(path, "w", encoding="utf-8") as f:
        f.write("# [002] DESIGN — 디자인 시스템 8요소 명세서\n\n")
        f.write("## 1. 캔버스\n")
        f.write(f"- **비율**: {_ratio(w, h)} ({cc.emu_to_in(w):.2f}\" × {cc.emu_to_in(h):.2f}\", {w:,} × {h:,} EMU)\n")
        lefts = [t["left"] for x in data["slides"] for t in x["text_slots"] if t["left"] and t["left"] > 0]
        if lefts:
            f.write(f"- **가장 흔한 본문 왼쪽 여백**: {cc.emu_to_in(Counter(lefts).most_common(1)[0][0]):.2f}\"\n")
        f.write("\n## 2. 타이포그래피\n")
        f.write(f"- **테마 제목 글꼴**: {th['major_font'] or '(없음)'}\n")
        f.write(f"- **테마 본문 글꼴**: {th['minor_font'] or '(없음)'}\n")
        f.write("- **실제 사용 글꼴 (사용 횟수)**:\n")
        for name, c in st["fonts"]:
            f.write(f"  - {name} ({c})\n")
        f.write("- **글자 크기 (큰 순, pt: 사용 횟수)**: ")
        f.write(", ".join(f"{int(k) if k == int(k) else k}pt:{v}" for k, v in st["font_sizes"][:12]) + "\n\n")
        f.write("## 3. 컬러 팔레트\n")
        f.write("| 구분 | HEX |\n|:---|:---|\n")
        for k, v in th["colors"].items():
            f.write(f"| 테마 {k} | `{v}` |\n")
        f.write("\n**실제로 많이 쓰인 직접 지정 색 (횟수)**: ")
        f.write(", ".join(f"`{c}`({n})" for c, n in st["colors"]) or "(없음)")
        f.write("\n\n**많이 쓰인 테마 색 슬롯**: ")
        f.write(", ".join(f"{c}({n})" for c, n in st["scheme_colors"]) or "(없음)")
        f.write("\n\n## 4. 레이아웃 아키타입\n")
        groups = {}
        for x in data["slides"]:
            groups.setdefault(x["archetype"], []).append(x["no"])
        for k, nos in groups.items():
            f.write(f"- **{ARCHETYPE_LABELS[k]}**: 슬라이드 {', '.join(map(str, nos))}\n")
        f.write("\n## 5. 컴포넌트 (도형 모양 빈도)\n")
        for g, c in st["geometries"]:
            f.write(f"- `{g}`: {c}\n")
        f.write("\n## 6. 시각 에셋\n")
        media = data["media"]
        total = sum(m["size"] for m in media.values())
        f.write(f"- 미디어 파일 {len(media)}개, 합계 {total / 1e6:.1f}MB (파일 전체 {data['file_size'] / 1e6:.1f}MB)\n")
        for name, m in list(media.items())[:5]:
            f.write(f"  - `{posixpath.basename(name)}` {m['size'] / 1e6:.1f}MB\n")
        f.write("\n## 7. 여백·간격\n")
        f.write("- 공통 반복 요소의 위치는 [003]_STYLE.md 1장 참고\n")
        f.write("\n## 8. 깊이감\n")
        f.write(f"- 바깥 그림자 {st['outer_shadows']}회, 부드러운 가장자리 {st['soft_edges']}회\n")
        if st["line_widths_pt"]:
            f.write("- 선 두께(pt: 횟수): " + ", ".join(f"{k}:{v}" for k, v in st["line_widths_pt"]) + "\n")


def write_style(data, path):
    st = data["style"]
    with open(path, "w", encoding="utf-8") as f:
        f.write("# [003] STYLE — 세부 스타일 및 컴포넌트 가이드\n\n")
        f.write("## 1. 공통 반복 요소 (머리글·배지·구분선)\n")
        if data["chrome"]:
            f.write("| 종류 | 이름 | 위치(in) | 크기(in) | 문구 | 등장 |\n|:---|:---|:---|:---|:---|:---|\n")
            for c in data["chrome"]:
                f.write(f"| {c['type']} | {c['name']} | {cc.emu_to_in(c['left']):.2f}, {cc.emu_to_in(c['top']):.2f} | "
                        f"{cc.emu_to_in(c['width']):.2f} × {cc.emu_to_in(c['height']):.2f} | {_short(c['text'], 30) or '-'} | {c['slides']}장 |\n")
        else:
            f.write("- 반복 요소 없음\n")
        f.write("\n## 2. 도형·선\n")
        if st["round_rect_adj"]:
            f.write("- 둥근 사각형 모서리 값: " + ", ".join(f"`{k}`({v})" for k, v in st["round_rect_adj"]) + "\n")
        f.write(f"- 바깥 그림자 사용: {st['outer_shadows']}회\n")
        f.write("\n## 3. 표 스타일\n")
        tables = [(x["no"], t) for x in data["slides"] for t in x["tables"]]
        if not tables:
            f.write("- 표 없음\n")
        for no, t in tables:
            rows = t["rows"]
            f.write(f"- Slide {no:02d} `{t['name']}`: {len(rows)}행 × {len(rows[0]) if rows else 0}열, 머리글: {' | '.join(rows[0]) if rows else ''}\n")
        f.write("\n## 4. 복제 시 지킬 점\n")
        f.write("- 위 반복 요소는 위치·크기를 바꾸지 말고 글자만 교체합니다.\n")
        f.write("- 이미지는 `ppt/media` 파일을 같은 이름으로 바꿔 끼워 마스크·그림자·좌표를 유지합니다.\n")


def write_template(data, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("# [004] TEMPLATE — 슬라이드별 템플릿 매핑\n\n")
        f.write("> `슬롯 ID`는 [009]_콘텐츠_매핑표.json에서 새 문구를 넣을 때 쓰는 이름입니다.\n\n")
        f.write("| 슬라이드 | 유형 | 레이아웃 | 도형 | 이미지 | 바꿀 글자 슬롯 |\n|:---|:---|:---|:---|:---|:---|\n")
        for x in data["slides"]:
            slots = [t for t in x["text_slots"] if not t["repeated"]]
            f.write(f"| {x['no']:02d} {_short(x['title'], 24)} | {ARCHETYPE_LABELS[x['archetype']]} | {x['layout']} | "
                    f"{x['shapes_count']} | {x['pictures']} | {len(slots)} |\n")
        f.write("\n## 슬롯 상세\n")
        for x in data["slides"]:
            f.write(f"\n### Slide {x['no']:02d}\n")
            for t in x["text_slots"]:
                mark = " (반복 머리글)" if t["repeated"] else ""
                f.write(f"- `{t['id']}` {t['name']}{mark}: {_short(t['text'], 60)}\n")


def main():
    cc.setup_stdout()
    args = cc.base_parser("[Step 1] 4대 설계도서 추출").parse_args()
    ref, _ = cc.resolve_reference(args)
    print(f"[*] 레퍼런스 분석: {ref}")
    data = extract(ref)
    out = args.output
    os.makedirs(out, exist_ok=True)
    write_business_dna(data, os.path.join(out, cc.F_DNA))
    write_design(data, os.path.join(out, cc.F_DESIGN))
    write_style(data, os.path.join(out, cc.F_STYLE))
    write_template(data, os.path.join(out, cc.F_TEMPLATE))
    cc.write_json(os.path.join(out, cc.F_PILLARS_JSON), data)
    cc.save_state(out, deck_title=deck_title(data), slides_count=data["slides_count"])
    print(f"[+] 4대 설계도서 생성: 슬라이드 {data['slides_count']}장, 텍스트 슬롯 "
          f"{sum(len(x['text_slots']) for x in data['slides'])}개 → {os.path.abspath(out)}")


if __name__ == "__main__":
    main()
