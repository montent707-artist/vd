# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | [Step 3] 1차 복제본 무결성 검증 + 부족/과도/추가 갭 진단 + 콘텐츠 매핑표 생성
=============================================================================
"""

import os
import sys
import posixpath

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import cloner_common as cc

LARGE_MEDIA_BYTES = 5 * 1024 * 1024
LOGO_MAX_AREA = 0.02      # 슬라이드 면적의 2% 이하
CORNER_MARGIN = 0.18      # 가장자리 18% 안쪽
DENSE_CHARS = 400


# ---------------------------------------------------------------------------
# 무결성
# ---------------------------------------------------------------------------
def texts_of(prs):
    return [[cc.shape_text(s) for s, _, _ in cc.iter_shapes(sl.shapes)] for sl in prs.slides]


def integrity(ref, clone):
    a, b = Presentation(ref), Presentation(clone)
    ua, ub = cc.media_usage(ref), cc.media_usage(clone)
    return {
        "sha256_match": cc.sha256(ref) == cc.sha256(clone),
        "slides": [len(a.slides), len(b.slides)],
        "text_match": texts_of(a) == texts_of(b),
        "shapes": [sum(1 for sl in a.slides for _ in cc.iter_shapes(sl.shapes)),
                   sum(1 for sl in b.slides for _ in cc.iter_shapes(sl.shapes))],
        "media": [len(ua), len(ub)],
        "size": [os.path.getsize(ref), os.path.getsize(clone)],
    }


# ---------------------------------------------------------------------------
# 부족한 점: 외주 로고·링크·메타데이터·안내 문구
# ---------------------------------------------------------------------------
def shape_hyperlink(shape):
    try:
        return shape.click_action.hyperlink.address or ""
    except Exception:
        return ""


def run_hyperlinks(shape):
    links = []
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                try:
                    if r.hyperlink.address:
                        links.append(r.hyperlink.address)
                except Exception:
                    pass
    return links


def in_corner(shape, w, h):
    if shape.left is None or shape.top is None:
        return False
    cx = shape.left + (shape.width or 0) / 2
    cy = shape.top + (shape.height or 0) / 2
    near_x = cx < w * CORNER_MARGIN or cx > w * (1 - CORNER_MARGIN)
    near_y = cy < h * CORNER_MARGIN or cy > h * (1 - CORNER_MARGIN)
    return near_x and near_y


def find_deficiencies(prs):
    w, h = prs.slide_width, prs.slide_height
    area = w * h
    purge, placeholders, links = [], [], []
    for no, slide in enumerate(prs.slides, 1):
        for shape, depth, parent in cc.iter_shapes(slide.shapes):
            sid = cc.slot_id(no, shape)
            text = cc.shape_text(shape)
            vendor = cc.contains_any(text, cc.VENDOR_KEYWORDS)
            href = shape_hyperlink(shape)
            href_vendor = cc.contains_any(href, cc.VENDOR_KEYWORDS) if href else []
            small = shape.width and shape.height and (shape.width * shape.height) / area <= LOGO_MAX_AREA
            st = shape.shape_type

            if vendor and len(text) <= 80:
                purge.append({"id": sid, "slide": no, "shape_id": shape.shape_id, "name": shape.name,
                              "reason": f"외주/워터마크 문구 포함: {', '.join(vendor)}", "text": text, "auto": True})
            elif href and (href_vendor or (small and st in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP))):
                purge.append({"id": sid, "slide": no, "shape_id": shape.shape_id, "name": shape.name,
                              "reason": f"외부 링크가 걸린 작은 로고형 도형: {href}", "text": text, "auto": True})
            elif (depth == 0 and small and st in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP)
                  and in_corner(shape, w, h)):
                purge.append({"id": sid, "slide": no, "shape_id": shape.shape_id, "name": shape.name,
                              "reason": "모서리에 있는 작은 로고형 이미지 (자체 로고일 수 있어 자동 삭제 안 함)",
                              "text": text, "auto": False})

            if href:
                links.append({"id": sid, "kind": "shape", "target": href})
            for target in run_hyperlinks(shape):
                links.append({"id": sid, "kind": "text", "target": target})

            ph = cc.contains_any(text, cc.PLACEHOLDER_PHRASES)
            if ph:
                placeholders.append({"id": sid, "slide": no, "phrases": ph, "text": text})
    return purge, placeholders, links


# ---------------------------------------------------------------------------
# 과도한 점: 미사용·대용량 미디어, 겹친 이미지, 빽빽한 슬라이드
# ---------------------------------------------------------------------------
def overlap_ratio(a, b):
    x1, y1 = max(a.left, b.left), max(a.top, b.top)
    x2 = min(a.left + a.width, b.left + b.width)
    y2 = min(a.top + a.height, b.top + b.height)
    if x2 <= x1 or y2 <= y1:
        return 0.0
    inter = (x2 - x1) * (y2 - y1)
    return inter / min(a.width * a.height, b.width * b.height)


def find_excesses(prs, pptx_path, pillars):
    usage = cc.media_usage(pptx_path)
    unused = [{"name": n, "size": m["size"]} for n, m in usage.items() if not m["used_by"]]
    large = [{"name": n, "size": m["size"]} for n, m in usage.items()
             if m["used_by"] and m["size"] >= LARGE_MEDIA_BYTES and not n.lower().endswith((".mp4", ".mov", ".wmv", ".avi", ".m4a", ".mp3", ".wav"))]
    large.sort(key=lambda x: -x["size"])

    stacked = []
    for no, slide in enumerate(prs.slides, 1):
        # 슬라이드 전체를 덮는 배경 그림은 겹침 계산에서 뺀다
        full = prs.slide_width * prs.slide_height * 0.9
        pics = [s for s, _, _ in cc.iter_shapes(slide.shapes)
                if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.width and s.height and s.width * s.height < full]
        count = sum(1 for i, a in enumerate(pics) for b in pics[i + 1:] if overlap_ratio(a, b) >= 0.8)
        if count:
            stacked.append({"slide": no, "pairs": count, "pictures": len(pics)})

    dense = []
    for s in pillars["slides"]:
        small = [t["font_pt"] for t in s["text_slots"] if t["font_pt"] and t["font_pt"] < 12]
        if s["chars"] >= DENSE_CHARS or small:
            dense.append({"slide": s["no"], "chars": s["chars"], "small_font_min": min(small) if small else None})
    return unused, large, stacked, dense


# ---------------------------------------------------------------------------
# 추가할 점: 제목 없는 슬라이드, Executive 선별
# ---------------------------------------------------------------------------
def select_executive(pillars, target):
    slides = pillars["slides"]
    n = len(slides)
    if n <= target:
        return [s["no"] for s in slides], []
    scores = {}
    prev_title = None
    for s in slides:
        score, notes = 1.0, []
        if s["title"]:
            score += 1
        if s["title"] and s["title"] == prev_title:
            score -= 2
            notes.append("앞 슬라이드와 제목이 같음 (연속 시리즈)")
        if s["archetype"] in ("table", "process", "cards"):
            score += 1
        if s["archetype"] == "media":
            score -= 1
            notes.append("동영상 슬라이드 (요약 발표에서 재생 부담)")
        if s["archetype"] == "gallery":
            score -= 0.5
            notes.append("이미지 위주 슬라이드")
        if s["chars"] < 30 and s["archetype"] not in ("cover", "closing"):
            score -= 0.5
            notes.append("전달 문구가 거의 없음")
        scores[s["no"]] = (score, notes)
        prev_title = s["title"] or prev_title
    forced = {slides[0]["no"], slides[-1]["no"]}
    rest = sorted((no for no in scores if no not in forced), key=lambda no: (-scores[no][0], no))
    keep = sorted(forced | set(rest[: target - len(forced)]))
    drop = [{"slide": no, "reason": "; ".join(scores[no][1]) or "점수가 낮음 (상대적으로 부가 내용)"}
            for no in sorted(set(scores) - set(keep))]
    return keep, drop


# ---------------------------------------------------------------------------
# 콘텐츠 매핑표 ([009])
# ---------------------------------------------------------------------------
def content_map_template(pillars, purge):
    deck = [t for t in pillars["repeated_texts"]]
    return {
        "_help": [
            "이 파일을 채운 뒤 run_cloner.py(또는 remake_enhanced.py)를 다시 실행하면 2차 덱에 반영됩니다.",
            "replace: 덱 전체에서 '찾을 문구' → '바꿀 문구'. 값이 빈 문자열이면 바꾸지 않습니다.",
            "slots: 슬롯별로 new에 새 문구를 넣으면 그 도형의 글자만 바뀝니다(서식 유지). 줄바꿈은 \\n.",
            "purge: 추가로 지울 도형 ID 목록 (예: \"s01:5\"). keep: 자동 삭제 대상 중 남길 도형 ID.",
            "media: ppt/media 파일명 → 새 이미지 경로. 원래 크기/비율에 맞게 잘라 같은 자리에 넣습니다.",
            "exec_slides: Executive 압축본에 넣을 슬라이드 번호 목록 (비우면 자동 선별).",
            "emblem_on_cover: true면 표지에서 지운 로고 자리에 generate_assets.py가 만든 엠블럼을 넣습니다.",
        ],
        "topic": None,
        "author": "",
        "replace": {t: "" for t in deck},
        "slots": {t["id"]: {"original": t["text"], "new": ""}
                  for s in pillars["slides"] for t in s["text_slots"] if not t["repeated"]},
        "purge": [],
        "keep": [],
        "media": {},
        "exec_slides": None,
        "emblem_on_cover": False,
        "_purge_candidates": [{"id": p["id"], "auto": p["auto"], "reason": p["reason"]} for p in purge],
    }


# ---------------------------------------------------------------------------
# 보고서
# ---------------------------------------------------------------------------
def _mb(b):
    return f"{b / 1e6:.1f}MB"


def write_report(path, plan, pillars, state):
    it = plan["integrity"]
    ok = lambda v: "✅ 일치" if v else "❌ 불일치"
    with open(path, "w", encoding="utf-8") as f:
        f.write("# [005] 제1차 복제본 정밀 검증 및 제2차 개선 제안서\n\n")
        f.write(f"- **원본**: `{os.path.basename(state.get('reference_pptx', ''))}`\n")
        f.write(f"- **1차 복제본**: `{cc.F_CLONE}`\n\n")
        f.write("## 1. 1:1 무결성 검증\n")
        f.write("| 항목 | 원본 | 복제본 | 결과 |\n|:---|:---|:---|:---|\n")
        f.write(f"| 파일 해시(SHA-256) | - | - | {ok(it['sha256_match'])} |\n")
        f.write(f"| 슬라이드 수 | {it['slides'][0]} | {it['slides'][1]} | {ok(it['slides'][0] == it['slides'][1])} |\n")
        f.write(f"| 도형 수 | {it['shapes'][0]} | {it['shapes'][1]} | {ok(it['shapes'][0] == it['shapes'][1])} |\n")
        f.write(f"| 미디어 수 | {it['media'][0]} | {it['media'][1]} | {ok(it['media'][0] == it['media'][1])} |\n")
        f.write(f"| 텍스트 | - | - | {ok(it['text_match'])} |\n")
        f.write(f"| 파일 크기 | {_mb(it['size'][0])} | {_mb(it['size'][1])} | {ok(it['size'][0] == it['size'][1])} |\n\n")

        f.write("## 2. 부족한 부분 (Deficiencies)\n")
        f.write("### 2-1. 외주 로고·워터마크 후보\n")
        if plan["purge"]:
            f.write("| 도형 ID | 이름 | 이유 | 2차에서 |\n|:---|:---|:---|:---|\n")
            for p in plan["purge"]:
                f.write(f"| `{p['id']}` | {p['name']} | {p['reason']} | {'자동 삭제' if p['auto'] else '확인 필요 (매핑표 purge에 넣으면 삭제)'} |\n")
        else:
            f.write("- 발견되지 않음\n")
        f.write("\n### 2-2. 외부 링크\n")
        f.write(f"- 도형/글자 링크 {len(plan['links'])}개" + (" → 2차에서 모두 제거\n" if plan["links"] else "\n"))
        for l in plan["links"][:10]:
            f.write(f"  - `{l['id']}` ({l['kind']}): {l['target']}\n")
        f.write("\n### 2-3. 파일 속성(작성자) 정보\n")
        core = pillars["core"]
        f.write(f"- 작성자: `{core['author'] or '(없음)'}`, 마지막 수정자: `{core['last_modified_by'] or '(없음)'}`")
        f.write(" → 2차에서 정리\n" if plan["clean_core"] else "\n")
        f.write("\n### 2-4. 템플릿 안내 문구 잔재\n")
        if plan["placeholders"]:
            for p in plan["placeholders"]:
                f.write(f"- `{p['id']}`: {', '.join(p['phrases'])} — \"{' '.join(p['text'].split())[:50]}\"\n")
            f.write("- 매핑표에 새 문구가 없으면 2차에서 해당 문구만 지웁니다.\n")
        else:
            f.write("- 없음\n")

        f.write("\n## 3. 과도한 부분 (Excesses)\n")
        f.write("### 3-1. 사용하지 않는 미디어\n")
        if plan["unused_media"]:
            total = sum(m["size"] for m in plan["unused_media"])
            f.write(f"- {len(plan['unused_media'])}개, {_mb(total)} → 2차 저장 시 자동 제거\n")
        else:
            f.write("- 없음\n")
        f.write("\n### 3-2. 대용량 이미지 (5MB 이상)\n")
        if plan["compress_media"]:
            for m in plan["compress_media"]:
                f.write(f"- `{posixpath.basename(m['name'])}` {_mb(m['size'])}\n")
            f.write("- 2차에서 긴 변 3840px(4K) 이하로 줄여 저장합니다. 화면 위치·크기는 그대로입니다.\n")
        else:
            f.write("- 없음\n")
        f.write("\n### 3-3. 같은 자리에 겹쳐 놓은 이미지\n")
        if plan["stacked"]:
            for s in plan["stacked"]:
                f.write(f"- Slide {s['slide']:02d}: 이미지 {s['pictures']}개 중 80% 이상 겹친 쌍 {s['pairs']}개 (애니메이션용인지 확인 필요)\n")
        else:
            f.write("- 없음\n")
        f.write("\n### 3-4. 글자가 빽빽하거나 작은 슬라이드\n")
        if plan["dense"]:
            for d in plan["dense"]:
                extra = f", 최소 글자 {d['small_font_min']}pt" if d["small_font_min"] else ""
                f.write(f"- Slide {d['slide']:02d}: {d['chars']}자{extra}\n")
        else:
            f.write("- 없음\n")

        f.write("\n## 4. 추가하면 좋은 요소 (Value-Add)\n")
        no_title = [s["no"] for s in pillars["slides"] if not s["title"]]
        if no_title:
            f.write(f"- 제목이 없는 슬라이드: {', '.join(map(str, no_title))} → 결론형 제목(Action Title)을 매핑표에 넣어 보세요.\n")
        f.write(f"- **Executive 압축본 ({len(plan['exec_keep'])}장)**: 슬라이드 {', '.join(map(str, plan['exec_keep']))}\n")
        for d in plan["exec_drop"]:
            f.write(f"  - 제외 Slide {d['slide']:02d}: {d['reason']}\n")
        f.write(f"- 새 문구는 `{cc.F_CONTENT_MAP}`에 넣으면 2차 덱에 반영됩니다.\n")

        f.write("\n## 5. 제2차 리메이크 실행 계획\n")
        f.write(f"1. 자동 삭제 대상 로고 {sum(1 for p in plan['purge'] if p['auto'])}개 제거, 외부 링크 {len(plan['links'])}개 제거\n")
        f.write("2. 파일 속성 정리, 안내 문구 정리\n")
        f.write("3. 매핑표의 새 문구·이미지 반영 (서식·위치 유지)\n")
        f.write("4. 미사용 미디어 제거, 대용량 이미지 축소\n")
        f.write(f"5. 완결본({pillars['slides_count']}장) + Executive 압축본({len(plan['exec_keep'])}장) 동시 생성\n")


def main():
    cc.setup_stdout()
    parser = cc.base_parser("[Step 3] 1:1 정밀 검증 및 3대 갭 분석")
    parser.add_argument("--exec-count", type=int, default=12, help="Executive 압축본 장수 (기본 12)")
    args = parser.parse_args()
    ref, state = cc.resolve_reference(args)
    clone = state.get("clone_pptx")
    pillars_path = os.path.join(args.output, cc.F_PILLARS_JSON)
    if not clone or not os.path.exists(clone) or not os.path.exists(pillars_path):
        sys.exit("[!] 1차 복제본이나 pillars.json이 없습니다. Step 1, 2를 먼저 실행해 주세요.")
    pillars = cc.read_json(pillars_path)

    print(f"[*] 1차 복제본 검증 및 갭 분석: {clone}")
    prs = Presentation(clone)
    purge, placeholders, links = find_deficiencies(prs)
    unused, large, stacked, dense = find_excesses(prs, clone, pillars)
    keep, drop = select_executive(pillars, args.exec_count)
    core = pillars["core"]
    plan = {
        "integrity": integrity(ref, clone),
        "purge": purge,
        "placeholders": placeholders,
        "links": links,
        "clean_core": bool(core["author"] or core["last_modified_by"]),
        "unused_media": unused,
        "compress_media": large,
        "stacked": stacked,
        "dense": dense,
        "exec_keep": keep,
        "exec_drop": drop,
    }
    cc.write_json(os.path.join(args.output, cc.F_AUDIT_PLAN), plan)
    write_report(os.path.join(args.output, cc.F_AUDIT_MD), plan, pillars, state)

    map_path = os.path.join(args.output, cc.F_CONTENT_MAP)
    if os.path.exists(map_path):
        print(f"[=] 기존 콘텐츠 매핑표를 그대로 둡니다: {map_path}")
    else:
        cc.write_json(map_path, content_map_template(pillars, purge))
        print(f"[+] 콘텐츠 매핑표 생성: {map_path}")

    it = plan["integrity"]
    print(f"[+] 무결성: 해시 {'일치' if it['sha256_match'] else '불일치'}, 텍스트 {'일치' if it['text_match'] else '불일치'}")
    print(f"[+] 로고 후보 {len(purge)}개(자동 {sum(1 for p in purge if p['auto'])}), 링크 {len(links)}개, "
          f"안내문구 {len(placeholders)}개, 미사용 미디어 {len(unused)}개, 대용량 이미지 {len(large)}개")
    print(f"[+] Executive 선별: {keep}")


if __name__ == "__main__":
    main()
