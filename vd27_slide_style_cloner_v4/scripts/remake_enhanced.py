# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | [Step 4] 1차 복제본 위에서 외과수술형 리메이크 (로고/링크/메타 정리,
        |             |                 | 서식 유지 문구 교체, 제자리 이미지 교체·축소, 완결본+Executive 동시 생성)
=============================================================================
"""

import io
import os
import sys
import copy
import posixpath

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

import cloner_common as cc
from generate_assets import fit_replacement

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "{%s}" % A_NS
MAX_IMAGE_EDGE = 3840
Image.MAX_IMAGE_PIXELS = 300_000_000  # 사용자 원본의 초고해상도 PNG(수십 MB)도 열 수 있게


# ---------------------------------------------------------------------------
# 도형 찾기/삭제
# ---------------------------------------------------------------------------
def find_shape(slide, shape_id):
    for shape, _, _ in cc.iter_shapes(slide.shapes):
        if shape.shape_id == shape_id:
            return shape
    return None


def parse_id(sid):
    s, i = sid.split(":")
    return int(s.lstrip("s")), int(i)


def purge_shapes(prs, ids):
    removed = []
    for sid in sorted(ids):
        slide_no, shape_id = parse_id(sid)
        if not 1 <= slide_no <= len(prs.slides):
            print(f"[!] 슬라이드 번호가 범위를 벗어나 건너뜀: {sid}")
            continue
        shape = find_shape(prs.slides[slide_no - 1], shape_id)
        if shape is None:
            print(f"[!] 도형을 찾지 못해 건너뜀: {sid}")
            continue
        removed.append({"id": sid, "name": shape.name, "left": shape.left, "top": shape.top,
                        "width": shape.width, "height": shape.height})
        el = shape._element
        el.getparent().remove(el)
    return removed


# ---------------------------------------------------------------------------
# 외부 링크 제거
# ---------------------------------------------------------------------------
def strip_external_links(prs):
    count = 0
    for slide in prs.slides:
        part = slide.part
        external = {rId for rId, rel in part.rels.items() if rel.is_external}
        if not external:
            continue
        for tag in ("hlinkClick", "hlinkHover"):
            for el in list(slide._element.iter(A + tag)):
                if el.get("{%s}id" % R_NS) in external:
                    el.getparent().remove(el)
                    count += 1
        for rId in external:
            if part.rels[rId].reltype == RT.HYPERLINK:
                part.drop_rel(rId)
    return count


# ---------------------------------------------------------------------------
# 서식을 유지하는 글자 교체
# ---------------------------------------------------------------------------
def _clear_runs_after_first(p):
    runs = p.runs
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for br in list(p._p.findall(A + "br")):
        p._p.remove(br)


def set_paragraph_text(p, text):
    if p.runs:
        _clear_runs_after_first(p)
        p.runs[0].text = text
    else:
        for br in list(p._p.findall(A + "br")):
            p._p.remove(br)
        r = p.add_run()
        r.text = text
        end = p._p.find(A + "endParaRPr")
        if end is not None:
            # 빈 문단의 서식(endParaRPr)을 새 글자에 그대로 입힌다
            rpr = copy.deepcopy(end)
            rpr.tag = A + "rPr"
            old = r._r.find(A + "rPr")
            if old is not None:
                r._r.remove(old)
            r._r.insert(0, rpr)


def set_text_preserving(shape, new_text):
    """도형의 글자만 바꾸고 글꼴·색·크기·정렬은 각 문단 첫 글자 서식을 이어받는다."""
    tf = shape.text_frame
    lines = new_text.split("\n")
    paras = list(tf.paragraphs)
    while len(paras) < len(lines):
        clone = copy.deepcopy(paras[-1]._p)
        paras[-1]._p.addnext(clone)
        paras = list(tf.paragraphs)
    for p in paras[len(lines):]:
        p._p.getparent().remove(p._p)
    for p, line in zip(tf.paragraphs, lines):
        set_paragraph_text(p, line)


def replace_in_paragraph(p, old, new):
    full = "".join(r.text for r in p.runs)
    if old not in full:
        return 0
    n = full.count(old)
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
    rest = "".join(r.text for r in p.runs)
    if old in rest:
        # 여러 글자 조각(run)에 걸쳐 있으면 첫 조각 서식으로 합친다
        runs = p.runs
        runs[0].text = rest.replace(old, new)
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    return n


def iter_paragraphs(prs):
    for slide in prs.slides:
        for shape, _, _ in cc.iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                yield from shape.text_frame.paragraphs
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        yield from cell.text_frame.paragraphs


def global_replace(prs, mapping):
    total = 0
    pairs = sorted(((o, n) for o, n in mapping.items() if o and n), key=lambda kv: -len(kv[0]))
    for p in iter_paragraphs(prs):
        for old, new in pairs:
            total += replace_in_paragraph(p, old, new)
    return total


def clear_placeholder_phrases(prs, skip_ids):
    cleared = 0
    for no, slide in enumerate(prs.slides, 1):
        for shape, _, _ in cc.iter_shapes(slide.shapes):
            if cc.slot_id(no, shape) in skip_ids or not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                if cc.contains_any(p.text, cc.PLACEHOLDER_PHRASES):
                    set_paragraph_text(p, "")
                    cleared += 1
    return cleared


# ---------------------------------------------------------------------------
# 이미지: 제자리 교체 / 축소
# ---------------------------------------------------------------------------
def image_parts(prs):
    parts = {}
    for part in prs.part.package.iter_parts():
        name = str(part.partname)
        if name.startswith("/ppt/media/") and hasattr(part, "_blob"):
            parts[posixpath.basename(name)] = part
    return parts


def replace_media(prs, mapping, base_dir):
    parts = image_parts(prs)
    done = []
    for media_name, new_path in mapping.items():
        if not new_path:
            continue
        part = parts.get(posixpath.basename(media_name))
        path = new_path if os.path.isabs(new_path) else os.path.join(base_dir, new_path)
        if part is None:
            print(f"[!] 미디어를 찾지 못해 건너뜀: {media_name}")
            continue
        if not os.path.exists(path):
            print(f"[!] 교체 이미지가 없어 건너뜀: {path}")
            continue
        with Image.open(io.BytesIO(part._blob)) as orig:
            size, fmt = orig.size, orig.format
        part._blob = fit_replacement(path, size, fmt)
        done.append(media_name)
    return done


def compress_media(prs, names, skip):
    parts = image_parts(prs)
    saved = 0
    for name in names:
        base = posixpath.basename(name)
        if base in skip or base not in parts:
            continue
        part = parts[base]
        try:
            with Image.open(io.BytesIO(part._blob)) as img:
                fmt = img.format
                if max(img.size) <= MAX_IMAGE_EDGE or fmt not in ("PNG", "JPEG"):
                    continue
                scale = MAX_IMAGE_EDGE / max(img.size)
                resized = img.resize((round(img.width * scale), round(img.height * scale)), Image.LANCZOS)
                buf = io.BytesIO()
                if fmt == "JPEG":
                    resized.convert("RGB").save(buf, "JPEG", quality=90, optimize=True)
                else:
                    resized.save(buf, "PNG", optimize=True)
        except Exception as e:
            print(f"[!] 이미지 축소 실패, 원본 유지: {base} ({e})")
            continue
        new = buf.getvalue()
        if len(new) < len(part._blob):
            saved += len(part._blob) - len(new)
            part._blob = new
    return saved


# ---------------------------------------------------------------------------
# 슬라이드 삭제 (Executive 압축본)
# ---------------------------------------------------------------------------
def keep_only_slides(prs, keep_nos):
    sld_ids = prs.slides._sldIdLst
    for idx in range(len(sld_ids) - 1, -1, -1):
        if idx + 1 not in keep_nos:
            sld = sld_ids[idx]
            prs.part.drop_rel(sld.rId)
            sld_ids.remove(sld)


# ---------------------------------------------------------------------------
def main():
    cc.setup_stdout()
    parser = cc.base_parser("[Step 4] 제2차 고도화 리메이크")
    parser.add_argument("--content", "-c", default=None, help=f"콘텐츠 매핑표 경로 (기본: 출력 폴더의 {cc.F_CONTENT_MAP})")
    parser.add_argument("--no-compress", action="store_true", help="대용량 이미지 축소를 하지 않음")
    args = parser.parse_args()

    ref, state = cc.resolve_reference(args)
    clone = state.get("clone_pptx")
    plan_path = os.path.join(args.output, cc.F_AUDIT_PLAN)
    if not clone or not os.path.exists(clone) or not os.path.exists(plan_path):
        sys.exit("[!] 1차 복제본이나 audit_plan.json이 없습니다. Step 2, 3을 먼저 실행해 주세요.")
    plan = cc.read_json(plan_path)
    map_path = args.content or os.path.join(args.output, cc.F_CONTENT_MAP)
    cmap = cc.read_json(map_path) if os.path.exists(map_path) else {}
    map_dir = os.path.dirname(os.path.abspath(map_path))

    topic = args.topic or cmap.get("topic") or state.get("topic")
    print(f"[*] 제2차 리메이크 시작: {clone}" + (f" (주제: {topic})" if topic else ""))
    prs = Presentation(clone)
    total = len(prs.slides)

    # 1) 로고/워터마크 제거
    keep_ids = set(cmap.get("keep") or [])
    purge_ids = {p["id"] for p in plan["purge"] if p["auto"]} - keep_ids
    purge_ids |= set(cmap.get("purge") or [])
    removed = purge_shapes(prs, purge_ids)

    # 2) 외부 링크 제거
    links = strip_external_links(prs)

    # 3) 문구 교체: 슬롯 → 전체 치환 → 안내 문구 정리
    slot_done = set()
    for sid, entry in (cmap.get("slots") or {}).items():
        new = (entry or {}).get("new") if isinstance(entry, dict) else entry
        if not new:
            continue
        slide_no, shape_id = parse_id(sid)
        shape = find_shape(prs.slides[slide_no - 1], shape_id) if 1 <= slide_no <= total else None
        if shape is None or not shape.has_text_frame:
            print(f"[!] 글자 슬롯을 찾지 못해 건너뜀: {sid}")
            continue
        set_text_preserving(shape, new)
        slot_done.add(sid)
    replace_map = dict(cmap.get("replace") or {})
    deck_title = state.get("deck_title")
    if args.topic and deck_title and not replace_map.get(deck_title):
        replace_map[deck_title] = args.topic
    replaced = global_replace(prs, replace_map)
    cleared = clear_placeholder_phrases(prs, slot_done)

    # 4) 파일 속성 정리
    author = cmap.get("author") or ""
    cp = prs.core_properties
    cp.author = author
    cp.last_modified_by = author
    cp.revision = 1
    if topic:
        cp.title = topic

    # 5) 이미지 교체/축소
    media_done = replace_media(prs, cmap.get("media") or {}, map_dir)
    saved = 0
    if not args.no_compress:
        saved = compress_media(prs, [m["name"] for m in plan["compress_media"]], set(media_done))

    # 6) 표지 로고 자리에 자체 엠블럼
    emblem = (state.get("assets") or {}).get("emblem")
    cover_removed = [r for r in removed if r["id"].startswith("s01:")]
    if cmap.get("emblem_on_cover") and emblem and os.path.exists(emblem) and cover_removed:
        r = cover_removed[0]
        side = min(r["width"], r["height"])
        prs.slides[0].shapes.add_picture(emblem, r["left"], r["top"], side, side)
        print("[+] 표지 로고 자리에 자체 엠블럼 배치")

    full_path = os.path.join(args.output, cc.full_deck_name(total))
    prs.save(full_path)
    print(f"[+] 완결본 저장: {full_path} ({os.path.getsize(full_path) / 1e6:.1f}MB)")

    # 7) Executive 압축본
    keep = cmap.get("exec_slides") or plan["exec_keep"]
    keep = sorted({int(n) for n in keep if 1 <= int(n) <= total})
    exec_prs = Presentation(full_path)
    keep_only_slides(exec_prs, set(keep))
    exec_path = os.path.join(args.output, cc.exec_deck_name(len(keep)))
    exec_prs.save(exec_path)
    print(f"[+] Executive 압축본 저장: {exec_path} ({len(keep)}장, {os.path.getsize(exec_path) / 1e6:.1f}MB)")

    cc.save_state(args.output, full_pptx=os.path.abspath(full_path), exec_pptx=os.path.abspath(exec_path),
                  exec_keep=keep, purged=[r["id"] for r in removed], applied_slots=sorted(slot_done),
                  author=author, remake_stats={
                      "removed_shapes": len(removed), "removed_links": links, "slots": len(slot_done),
                      "global_replacements": replaced, "placeholders_cleared": cleared,
                      "media_replaced": media_done, "compress_saved_bytes": saved})
    print(f"[+] 로고 {len(removed)}개 제거, 링크 {links}개 제거, 슬롯 {len(slot_done)}개·전체치환 {replaced}회 반영, "
          f"안내문구 {cleared}개 정리, 이미지 교체 {len(media_done)}개, 축소로 {saved / 1e6:.1f}MB 절감")


if __name__ == "__main__":
    main()
