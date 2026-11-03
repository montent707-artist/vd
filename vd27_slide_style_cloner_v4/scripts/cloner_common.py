# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | 모든 단계 스크립트가 함께 쓰는 공통 유틸리티 (입력 탐색, 상태 파일, 도형 순회)
=============================================================================
"""

import os
import sys
import glob
import json
import hashlib
import argparse
import zipfile
import posixpath

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# 산출물 파일명 (SKILL.md 3장 번호 체계)
F_DNA = "[001]_BUSINESS_DNA.md"
F_DESIGN = "[002]_DESIGN.md"
F_STYLE = "[003]_STYLE.md"
F_TEMPLATE = "[004]_TEMPLATE.md"
F_AUDIT_MD = "[005]_1차복제_정밀검증_및_개선제안서.md"
F_CLONE = "[001]_제1차_100퍼센트_완전복제본.pptx"
F_CONTENT_MAP = "[009]_콘텐츠_매핑표.json"
F_MANIFEST = "[000]_산출물_목록.md"
F_PILLARS_JSON = "pillars.json"
F_AUDIT_PLAN = "audit_plan.json"
F_QUALITY = "audit_result.json"
F_STATE = ".cloner_state.json"


def full_deck_name(n):
    return f"[006]_제2차_고도화_완결실전본({n}장).pptx"


def exec_deck_name(n):
    return f"[007]_제2차_고도화_Executive압축본({n}장).pptx"


# 외주 템플릿 업체/워터마크로 의심되는 문구 (소문자 비교)
VENDOR_KEYWORDS = [
    "예스폼", "yesform", "망고보드", "mangoboard", "미리캔버스", "miricanvas",
    "피피티월드", "pptworld", "비즈폼", "bizforms", "슬라이드고", "slidesgo",
    "freepik", "envato", "slidemodel", "canva", "templatelab", "www.",
    "http://", "https://", "무단 배포", "무단배포", "저작권은", "all rights reserved",
]

# 템플릿에 남아 있는 안내 문구 (소문자 비교)
PLACEHOLDER_PHRASES = [
    "내용을 입력", "텍스트를 입력", "제목을 입력", "내용 입력", "텍스트 입력",
    "세부내용 기술", "여기에 입력", "lorem ipsum", "click to add", "click to edit",
    "your text here", "add your text", "enter your text",
]


def setup_stdout():
    """Windows 콘솔(cp949)에서도 한글/특수문자 출력이 깨져 멈추지 않도록 UTF-8로 맞춘다."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def base_parser(description):
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("--input", "-i", default=None, help="레퍼런스 PPTX/PDF 경로")
    parser.add_argument("--output", "-o", default="result", help="산출물 폴더 경로")
    parser.add_argument("--topic", "-t", default=None, help="신규 도메인 주제")
    return parser


def load_state(output_dir):
    path = os.path.join(output_dir, F_STATE)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(output_dir, **updates):
    state = load_state(output_dir)
    state.update({k: v for k, v in updates.items() if v is not None})
    os.makedirs(output_dir, exist_ok=True)
    with open(os.path.join(output_dir, F_STATE), "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    return state


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def find_input(explicit=None):
    """레퍼런스 파일 탐색: 지정 경로 → 현재 폴더 → upload/ → input/ 순."""
    if explicit:
        if not os.path.exists(explicit):
            sys.exit(f"[!] 입력 파일이 없습니다: {explicit}")
        return os.path.abspath(explicit)
    for pattern in ("*.pptx", "upload/*.pptx", "input/*.pptx", "*.pdf", "upload/*.pdf", "input/*.pdf"):
        found = sorted(glob.glob(pattern))
        if found:
            return os.path.abspath(found[0])
    return None


def resolve_reference(args):
    """
    이번 실행에서 쓸 레퍼런스 PPTX 경로를 정한다.
    - 상태 파일에 이미 정해져 있으면 그대로 쓴다 (단계별 스크립트가 같은 파일을 보게 함).
    - .ppt는 지원하지 않는다는 안내 후 종료.
    - .pdf는 페이지 이미지 기반 시드 PPTX로 변환한다 (pymupdf 필요).
    - 입력이 없으면 내장 Mock 시드 PPTX를 만든다.
    """
    state = load_state(args.output)
    if not args.input and state.get("reference_pptx") and os.path.exists(state["reference_pptx"]):
        return state["reference_pptx"], state

    src = find_input(args.input)
    if src and src.lower().endswith(".ppt"):
        sys.exit("[!] 예전 형식(.ppt)은 지원하지 않습니다. PowerPoint에서 .pptx로 다시 저장한 뒤 넣어 주세요.")

    seed_dir = os.path.join(args.output, "_seed")
    if src is None:
        from mock_seed import build_mock_seed
        os.makedirs(seed_dir, exist_ok=True)
        ref = os.path.abspath(os.path.join(seed_dir, "MOCK_SEED.pptx"))
        build_mock_seed(ref, args.topic or "비즈니스 사업계획서")
        print(f"[!] 레퍼런스 파일이 없어 내장 Mock 시드를 만들었습니다: {ref}")
        source_kind = "mock"
    elif src.lower().endswith(".pdf"):
        os.makedirs(seed_dir, exist_ok=True)
        ref = os.path.abspath(os.path.join(seed_dir, os.path.splitext(os.path.basename(src))[0] + ".pptx"))
        pdf_to_pptx(src, ref)
        print(f"[+] PDF를 페이지 이미지 기반 PPTX로 변환했습니다: {ref}")
        source_kind = "pdf"
    else:
        ref = src
        source_kind = "pptx"

    state = save_state(args.output, original_input=src or "MOCK_SEED", reference_pptx=ref,
                       source_kind=source_kind, topic=args.topic)
    return ref, state


def pdf_to_pptx(pdf_path, out_path, dpi=150):
    try:
        import fitz  # pymupdf
    except ImportError:
        sys.exit("[!] PDF 입력에는 pymupdf가 필요합니다. 먼저 `uv pip install pymupdf`로 설치해 주세요.")
    from pptx.util import Emu
    doc = fitz.open(pdf_path)
    first = doc[0].rect
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(int(12192000 * first.height / first.width))
    blank = prs.slide_layouts[6]
    tmp_dir = os.path.dirname(out_path)
    for i, page in enumerate(doc, 1):
        png = os.path.join(tmp_dir, f"_page_{i:03d}.png")
        page.get_pixmap(dpi=dpi).save(png)
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(png, 0, 0, prs.slide_width, prs.slide_height)
        text = page.get_text().strip()
        if text:
            slide.notes_slide.notes_text_frame.text = text
        os.remove(png)
    prs.save(out_path)


def iter_shapes(shapes, depth=0, parent=None):
    """그룹 안쪽까지 모든 도형을 (도형, 깊이, 부모그룹) 으로 순회."""
    for shape in shapes:
        yield shape, depth, parent
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes, depth + 1, shape)


def shape_text(shape):
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        return shape.text_frame.text.replace("\x0b", "\n").strip()
    return ""


def table_text(shape):
    if getattr(shape, "has_table", False) and shape.has_table:
        return [[c.text.replace("\x0b", " ").strip() for c in row.cells] for row in shape.table.rows]
    return []


def shape_type_name(shape):
    try:
        return shape.shape_type.name if shape.shape_type is not None else "UNKNOWN"
    except Exception:
        return "UNKNOWN"


def slot_id(slide_no, shape):
    """텍스트 슬롯 식별자: 슬라이드 번호 + 도형 id (이름은 중복될 수 있어 id 사용)."""
    return f"s{slide_no:02d}:{shape.shape_id}"


def title_candidates(slide):
    """
    슬라이드 제목 후보 목록: (첫 줄 텍스트, 최대 글자크기, top, 제목 플레이스홀더 여부).
    반복 헤더(모든 슬라이드에 같은 문구)를 걸러 최종 제목을 고르는 일은 pick_title()이 한다.
    """
    candidates = []
    for shape, _, _ in iter_shapes(slide.shapes):
        text = shape_text(shape)
        if not text:
            continue
        size = max_font_pt(shape) or 0
        is_title_ph = False
        if shape.is_placeholder:
            try:
                is_title_ph = shape.placeholder_format.type.name in ("TITLE", "CENTER_TITLE")
            except Exception:
                pass
        top = shape.top or 0
        candidates.append((text.split("\n")[0].strip(), size, top, is_title_ph))
    return candidates


def repeated_texts(prs, min_ratio=0.5):
    """절반 이상의 슬라이드에 똑같이 나오는 짧은 문구 (머리글, 회차 배지 등)."""
    from collections import Counter
    counter = Counter()
    for slide in prs.slides:
        seen = set()
        for shape, _, _ in iter_shapes(slide.shapes):
            t = shape_text(shape)
            if t and len(t) <= 40:
                seen.add(t)
        counter.update(seen)
    n = len(prs.slides)
    return {t for t, c in counter.items() if n >= 3 and c / n >= min_ratio}


def pick_title(slide, repeated, slide_height, is_cover=False):
    """
    제목 플레이스홀더 → 상단 40% 안에서 가장 위에 있는 글자 순으로 제목을 고른다.
    표지가 아니면 반복 머리글(덱 제목 등)은 뺀다. 12pt 미만 각주는 후보에서 제외.
    """
    cands = [c for c in title_candidates(slide)
             if c[0] and (is_cover or c[0] not in repeated) and not (c[1] and c[1] < 12)]
    if not cands:
        return ""
    ph = [c for c in cands if c[3]]
    if ph:
        return ph[0][0]
    upper = [c for c in cands if c[2] < slide_height * 0.4] or cands
    upper.sort(key=lambda c: (c[2], -(c[1] or 0)))
    return upper[0][0]


def max_font_pt(shape):
    best = None
    if not shape.has_text_frame:
        return None
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size is not None:
                pt = r.font.size.pt
                best = pt if best is None else max(best, pt)
    return best


def media_usage(pptx_path):
    """
    ZIP 수준 미디어 분석: 각 미디어 파일 크기와, 어느 슬라이드/레이아웃/마스터가 참조하는지.
    참조가 하나도 없는 미디어는 '미사용'.
    """
    usage = {}
    with zipfile.ZipFile(pptx_path) as z:
        names = z.namelist()
        for n in names:
            if n.startswith("ppt/media/"):
                usage[n] = {"size": z.getinfo(n).file_size, "used_by": []}
        for n in names:
            if not n.endswith(".rels") or "/_rels/" not in n:
                continue
            owner_dir = posixpath.dirname(posixpath.dirname(n))
            owner = posixpath.join(owner_dir, posixpath.basename(n)[:-5])
            data = z.read(n).decode("utf-8", errors="ignore")
            for target in _rel_targets(data):
                if target.startswith("http") or target.startswith("file:"):
                    continue
                full = posixpath.normpath(posixpath.join(owner_dir, target))
                if full in usage:
                    usage[full]["used_by"].append(owner)
    return usage


def external_rels(pptx_path):
    """모든 파트의 외부 링크(하이퍼링크 등) 목록."""
    out = []
    from lxml import etree
    with zipfile.ZipFile(pptx_path) as z:
        for n in z.namelist():
            if not n.endswith(".rels"):
                continue
            root = etree.fromstring(z.read(n))
            for rel in root:
                if rel.get("TargetMode") == "External":
                    out.append({"rels": n, "type": rel.get("Type", "").split("/")[-1], "target": rel.get("Target")})
    return out


def _rel_targets(xml_text):
    from lxml import etree
    root = etree.fromstring(xml_text.encode("utf-8"))
    return [rel.get("Target", "") for rel in root if rel.get("TargetMode") != "External"]


def core_props(prs):
    cp = prs.core_properties
    return {
        "title": cp.title or "",
        "author": cp.author or "",
        "last_modified_by": cp.last_modified_by or "",
        "company_hint": cp.category or "",
        "revision": cp.revision,
    }


def contains_any(text, words):
    low = (text or "").lower()
    return [w for w in words if w in low]


def emu_to_in(v):
    return (v or 0) / 914400
