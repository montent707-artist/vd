# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-11-03  | Claude          | [Step 2] 레퍼런스를 바이트 단위로 복제하고 SHA-256으로 동일성 확인
=============================================================================
"""

import os
import sys
import shutil

from pptx import Presentation

import cloner_common as cc


def main():
    cc.setup_stdout()
    args = cc.base_parser("[Step 2] 제1차 100% 무손실 복제").parse_args()
    ref, state = cc.resolve_reference(args)
    out_path = os.path.join(args.output, cc.F_CLONE)
    os.makedirs(args.output, exist_ok=True)

    if os.path.abspath(ref) == os.path.abspath(out_path):
        sys.exit("[!] 입력 파일과 1차 복제본 경로가 같습니다. 다른 --output 폴더를 지정해 주세요.")

    print(f"[*] 무손실 복제: {ref} -> {out_path}")
    shutil.copyfile(ref, out_path)

    src_hash, dst_hash = cc.sha256(ref), cc.sha256(out_path)
    if src_hash != dst_hash:
        sys.exit("[!] 복제본 해시가 원본과 다릅니다. 디스크 공간이나 파일 잠금을 확인해 주세요.")
    slides = len(Presentation(out_path).slides)

    cc.save_state(args.output, clone_pptx=os.path.abspath(out_path), clone_sha256=dst_hash, slides_count=slides)
    print(f"[+] 1차 복제본 완성: {slides}장, {os.path.getsize(out_path):,} bytes, SHA-256 일치 ({dst_hash[:12]}…)")


if __name__ == "__main__":
    main()
