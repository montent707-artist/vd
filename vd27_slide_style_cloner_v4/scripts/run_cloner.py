# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-09-22  | Antigravity AI  | 살아 숨쉬는 유정란 스킬 마스터 실행기 (run_cloner.py)
REV.02  | 2026-11-03  | Claude          | --input/--topic/--content를 각 단계에 전달, 에셋 생성·산출물 목록(Step 6) 추가,
        |             |                 | 단계 실패 시 어느 단계인지 알려 주고 멈춤
=============================================================================
"""

import os
import sys
import subprocess
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cloner_common as cc


def run_step(title, script, args):
    print(f"\n{title}", flush=True)
    cmd = [sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), script)] + args
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    code = subprocess.run(cmd, env=env).returncode
    if code != 0:
        sys.exit(f"\n[!] {title} 단계에서 멈췄습니다 (종료 코드 {code}). 위 메시지를 확인해 주세요.")


def write_manifest(output_dir):
    state = cc.load_state(output_dir)
    rows = [
        (cc.F_DNA, "비즈니스 DNA 및 내러티브 구조"),
        (cc.F_DESIGN, "디자인 시스템 8요소 명세서"),
        (cc.F_STYLE, "세부 스타일 및 컴포넌트 가이드"),
        (cc.F_TEMPLATE, "슬라이드별 템플릿 매핑 명세"),
        (cc.F_AUDIT_MD, "부족/과도/추가 3대 갭 분석서"),
        (cc.F_CONTENT_MAP, "새 문구·이미지를 넣는 콘텐츠 매핑표 (직접 채움)"),
        (cc.F_CLONE, "1차 무손실 100% 복제본"),
        (os.path.basename(state.get("full_pptx", "")), "2차 완결 실전본"),
        (os.path.basename(state.get("exec_pptx", "")), "2차 Executive 압축본"),
        (cc.F_QUALITY, "무결성 품질 검증 리포트"),
        (cc.F_PILLARS_JSON, "분석 원자료 (기계용)"),
        (cc.F_AUDIT_PLAN, "리메이크 계획 (기계용)"),
    ]
    path = os.path.join(output_dir, cc.F_MANIFEST)
    with open(path, "w", encoding="utf-8") as f:
        f.write("# [000] 산출물 목록\n\n")
        f.write(f"- 생성 시각: {datetime.now():%Y-%m-%d %H:%M}\n")
        f.write(f"- 원본: `{state.get('original_input', '')}`\n")
        if state.get("topic"):
            f.write(f"- 주제: {state['topic']}\n")
        f.write("\n| 파일 | 설명 | 크기 |\n|:---|:---|:---|\n")
        for name, desc in rows:
            p = os.path.join(output_dir, name)
            if name and os.path.exists(p):
                size = os.path.getsize(p)
                shown = f"{size / 1e6:.1f}MB" if size > 1e6 else f"{size / 1e3:.0f}KB"
                f.write(f"| `{name}` | {desc} | {shown} |\n")
    return path


def run_pipeline(input_path=None, output_dir="result", topic=None, content=None, exec_count=12, no_compress=False):
    cc.setup_stdout()
    print("=" * 60)
    print(" 🥚 [Slide Style Cloner] 유정란 스킬 자율 부화 파이프라인 가동")
    print("=" * 60, flush=True)
    os.makedirs(output_dir, exist_ok=True)

    # 새 실행마다 상태를 새로 잡는다 (콘텐츠 매핑표는 사용자가 채운 것이므로 유지)
    state_file = os.path.join(output_dir, cc.F_STATE)
    if os.path.exists(state_file):
        os.remove(state_file)

    common = ["--output", output_dir] + (["--topic", topic] if topic else [])
    first = common + (["--input", input_path] if input_path else [])

    run_step("[Step 1] 4대 설계도서(BUSINESS_DNA, DESIGN, STYLE, TEMPLATE) 추출", "extract_4pillars.py", first)
    run_step("[Step 2] 제1차 100% 무손실 정밀 복제", "clone_exact_100.py", common)
    run_step("[Step 3] 1:1 정밀 검증 및 부족/과도/추가 3대 갭 분석", "audit_and_propose.py",
             common + ["--exec-count", str(exec_count)])
    run_step("[Step 4-1] 도메인 맞춤 엠블럼/배지 에셋 생성", "generate_assets.py", common)
    run_step("[Step 4-2] 제2차 고도화 리메이크 (로고 제거 + 완결본/Executive 동시 생성)", "remake_enhanced.py",
             common + (["--content", content] if content else []) + (["--no-compress"] if no_compress else []))
    run_step("[Step 5] 최종 무결성 품질 감사 (Quality Audit Gate)", "verify_quality.py", common)

    print("\n[Step 6] 산출물 번호 체계 정리")
    manifest = write_manifest(output_dir)
    print(f"[+] 산출물 목록: {manifest}")

    print("\n" + "=" * 60)
    print(" 🐥 [부화 완료] 4대 설계도서 + 1차 복제본 + 2차 덱 2종이 만들어졌습니다.")
    print(f" 📂 산출물 저장 위치: {os.path.abspath(output_dir)}")
    print(f" ✏️  새 문구를 넣으려면 {cc.F_CONTENT_MAP}을 채운 뒤 같은 명령을 다시 실행하세요.")
    print("=" * 60)


if __name__ == "__main__":
    parser = cc.base_parser("Slide Style Cloner Self-Executing Pipeline")
    parser.add_argument("--content", "-c", default=None, help="콘텐츠 매핑표 경로 (기본: 출력 폴더의 [009])")
    parser.add_argument("--exec-count", type=int, default=12, help="Executive 압축본 장수 (기본 12)")
    parser.add_argument("--no-compress", action="store_true", help="대용량 이미지 축소를 하지 않음")
    args = parser.parse_args()
    run_pipeline(args.input, args.output, args.topic, args.content, args.exec_count, args.no_compress)
