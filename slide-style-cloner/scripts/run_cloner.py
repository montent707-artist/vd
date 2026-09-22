# -*- coding: utf-8 -*-
"""
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-09-22  | Antigravity AI  | 살아 숨쉬는 유정란 스킬 마스터 실행기 (run_cloner.py)
=============================================================================
"""

import os
import sys
import argparse
import glob
import subprocess

def run_pipeline(input_pptx=None, output_dir="result", domain_topic="비즈니스 사업계획서 및 IR 피치덱"):
    print("=" * 60)
    print(" 🥚 [Slide Style Cloner] 유정란 스킬 자율 부화 파이프라인 가동")
    print("=" * 60)
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 대상 파일 탐색
    if not input_pptx:
        candidates = glob.glob("*.pptx") + glob.glob("upload/*.pptx") + glob.glob("input/*.pptx")
        if candidates:
            input_pptx = candidates[0]
            print(f"[+] 발견된 레퍼런스 슬라이드: {input_pptx}")
        else:
            print("[!] 레퍼런스 파일이 없습니다. 내장 시드(Seed)를 통해 100% 규격 가상 템플릿을 스스로 부화합니다.")
            input_pptx = "MOCK_SEED"

    # Step 1: 4대 설계도서 추출
    print("\n[Step 1] 4대 설계도서(BUSINESS_DNA, DESIGN, STYLE, TEMPLATE) 자동 추출/생성...")
    extract_script = os.path.join(current_dir, "extract_4pillars.py")
    subprocess.run([sys.executable, extract_script], check=True)

    # Step 2: 1차 무손실 복제
    print("\n[Step 2] 제1차 100% 무손실 정밀 복제 수행...")
    clone_script = os.path.join(current_dir, "clone_exact_100.py")
    subprocess.run([sys.executable, clone_script], check=True)

    # Step 3: 정밀 감사 및 갭 분석
    print("\n[Step 3] 1:1 정밀 검증 및 부족/과도/추가 3대 갭 분석...")
    audit_script = os.path.join(current_dir, "audit_and_propose.py")
    subprocess.run([sys.executable, audit_script], check=True)

    # Step 4: 2차 고도화 리메이크
    print("\n[Step 4] 제2차 고도화 리메이크 (외주 로고 0% 박멸 + 16/12장 듀얼 덱 생성)...")
    remake_script = os.path.join(current_dir, "remake_enhanced.py")
    subprocess.run([sys.executable, remake_script], check=True)

    # Step 5: 품질 검증
    print("\n[Step 5] 최종 무결성 품질 감사(Quality Audit Gate)...")
    verify_script = os.path.join(current_dir, "verify_quality.py")
    subprocess.run([sys.executable, verify_script], check=True)

    print("\n" + "=" * 60)
    print(" 🐥 [부화 완료] 유정란 스킬이 완벽한 3대 PPTX 및 4대 설계도서를 탄생시켰습니다!")
    print(f" 📂 산출물 저장 위치: {os.path.abspath(output_dir)}")
    print("=" * 60)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Slide Style Cloner Self-Executing Pipeline")
    parser.add_argument("--input", "-i", default=None, help="레퍼런스 PPTX 경로")
    parser.add_argument("--output", "-o", default="result", help="산출물 폴더 경로")
    parser.add_argument("--topic", "-t", default="비즈니스 사업계획서", help="신규 도메인 주제")
    args = parser.parse_args()
    
    run_pipeline(args.input, args.output, args.topic)
