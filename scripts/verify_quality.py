# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | [Step 5] 산출물 3종을 실제로 열어 슬라이드 수·로고·링크·안내문구·메타 검사
=============================================================================
"""

import os
import sys

from pptx import Presentation

import cloner_common as cc


def scan_text(prs):
    """외주 문구 / 템플릿 안내 문구가 남은 도형 ID."""
    vendor, placeholder = [], []
    for no, slide in enumerate(prs.slides, 1):
        for shape, _, _ in cc.iter_shapes(slide.shapes):
            rows = cc.table_text(shape)
            text = "\n".join(" ".join(r) for r in rows) if rows else cc.shape_text(shape)
            if cc.contains_any(text, cc.VENDOR_KEYWORDS):
                vendor.append(cc.slot_id(no, shape))
            if cc.contains_any(text, cc.PLACEHOLDER_PHRASES):
                placeholder.append(cc.slot_id(no, shape))
    return vendor, placeholder


def ids_still_present(prs, ids):
    found = []
    for sid in ids:
        s, i = sid.split(":")
        no = int(s.lstrip("s"))
        if no <= len(prs.slides) and any(sh.shape_id == int(i) for sh, _, _ in cc.iter_shapes(prs.slides[no - 1].shapes)):
            found.append(sid)
    return found


def check_clone(path, state):
    prs = Presentation(path)
    return {
        "slides": len(prs.slides),
        "expected_slides": state["slides_count"],
        "slides_ok": len(prs.slides) == state["slides_count"],
        "sha256_ok": cc.sha256(path) == cc.sha256(state["reference_pptx"]),
    }


def check_remake(path, state, expected, check_purged):
    prs = Presentation(path)
    vendor, placeholder = scan_text(prs)
    links = [l for l in cc.external_rels(path) if l["type"] == "hyperlink" and "/slides/" in l["rels"]]
    author = state.get("author") or ""
    left = ids_still_present(prs, state.get("purged") or []) if check_purged else []
    cp = prs.core_properties
    return {
        "slides": len(prs.slides),
        "expected_slides": expected,
        "slides_ok": len(prs.slides) == expected,
        "purged_shapes_left": left,
        "purged_ok": not left,
        "external_links": len(links),
        "links_ok": not links,
        "vendor_text": vendor,
        "vendor_ok": not vendor,
        "placeholder_text": placeholder,
        "placeholder_ok": not placeholder,
        "author_ok": (cp.author or "") == author and (cp.last_modified_by or "") == author,
    }


def run_check(name, path, fn):
    result = {"name": name, "path": path, "status": "FAIL", "checks": {}}
    if not path or not os.path.exists(path):
        result["checks"]["exists_ok"] = False
        return result
    try:
        checks = fn()
    except Exception as e:
        result["checks"]["opens_ok"] = False
        result["checks"]["error"] = str(e)
        return result
    checks["size_mb"] = round(os.path.getsize(path) / 1e6, 2)
    result["checks"] = checks
    result["status"] = "PASS" if all(v for k, v in checks.items() if k.endswith("_ok")) else "FAIL"
    return result


def main():
    cc.setup_stdout()
    args = cc.base_parser("[Step 5] 최종 품질 검증").parse_args()
    state = cc.load_state(args.output)
    if not state.get("full_pptx"):
        sys.exit("[!] 2차 산출물 정보가 없습니다. Step 4를 먼저 실행해 주세요.")

    clone, full, exe = state.get("clone_pptx"), state.get("full_pptx"), state.get("exec_pptx")
    results = [
        run_check("1차 100% 복제본", clone, lambda: check_clone(clone, state)),
        run_check("2차 완결본", full, lambda: check_remake(full, state, state["slides_count"], True)),
        run_check("2차 Executive 압축본", exe,
                  lambda: check_remake(exe, state, len(state.get("exec_keep") or []), False)),
    ]
    cc.write_json(os.path.join(args.output, cc.F_QUALITY), results)

    for r in results:
        c = r["checks"]
        print(f"[{r['status']}] {r['name']}: {c.get('slides', '?')}/{c.get('expected_slides', '?')}장, {c.get('size_mb', '?')}MB")
        for k, v in c.items():
            if k.endswith("_ok") and not v:
                print(f"    - 실패 항목: {k}")
    if any(r["status"] != "PASS" for r in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
