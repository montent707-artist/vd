---
name: slide-style-cloner
description: "Reverse engineer and clone executive presentation styles from any input PDF or PPTX slide deck, extract 4 Pillars (BUSINESS_DNA, DESIGN, STYLE, TEMPLATE), execute 1st 100% exact cloning, audit deficiencies/excesses/value-adds, and generate 2nd enhanced executive presentation slides (30/16-slide Full & 12-slide Executive). Features Surgical In-Place Remake, Vector Logo Purge, and In-Place Media Replacement. Self-contained with executable scripts for zero-dependency instant generation."
---

# 🥚 Slide Style Cloner v4.0 (살아 숨쉬는 완전 실행형 유정란 스킬)

> **💡 핵심 철학 (Core Philosophy)**:  
> 레퍼런스 템플릿의 **압도적인 3D 그래픽, 세련된 컬러 스키마, 감각적인 카드 레이아웃을 100% 온전히 보존**하면서, 외주사 로고와 불필요한 메타데이터를 정밀 박멸하고 실전 비즈니스 로직을 주입하는 **'외과수술형 2단계 슬라이드 복제(Two-Stage Surgical Remake)'** 전문 스킬입니다.

---

## 1. 🛡️ 안티패턴 및 4대 방어 수칙 (What NOT to Do)

1. **절대 백지에서 새로 그리지 마십시오 (Anti-Blank Canvas)**:
   - 템플릿의 틀을 피하겠다고 원본 템플릿 파일(.pptx)을 버리고 빈 캔버스에서 사각형 박스로 새로 그리면, 원본의 수십 MB급 고화질 3D 그래픽과 미학 에셋이 전부 소실되고 조잡한 상자 PPT가 됩니다.
2. **단순 텍스트 치환에 머물지 마십시오 (Anti-Blind Replacement)**:
   - 표지 상단에 벡터 도형(`GROUP`)으로 숨겨진 외주사(예스폼 등) 로고 및 `_rels` 하이퍼링크를 반드시 영구 박멸하십시오.
   - 도메인과 맞지 않는 외국인 스톡 모델 사진이나 어색한 사무실 사진을 방치하지 마십시오.
3. **단일 덱 발행에 안주하지 마십시오 (Multi-Deck Publishing)**:
   - 16장/30장 완결 실전본(Full Deck)과 12장 Executive 압축본(Pitch Deck)을 항상 쌍으로 발행하십시오.
4. **정적 문서(무정란)로 방치하지 마십시오 (Executable Self-Contained)**:
   - 스킬은 단순 설명서가 아니라, 자체 실행 엔진(`scripts/`)을 품고 언제든 스스로 부화하여 산출물을 낳아야 합니다.

---

## 2. 🧬 6대 엔드-투-엔드 자율 파이프라인 (6-Step Master Pipeline)

```
[입력: 레퍼런스 PPTX 또는 Mock 가상 템플릿]
                           │
                           ▼
[Step 1] 4대 핵심 설계도서 역공학 추출 (Extract 4 Pillars)
  - [001]_BUSINESS_DNA.md : 비즈니스 목표, 청중, 16장 내러티브 흐름 구조화
  - [002]_DESIGN.md       : 디자인 시스템 8요소(캔버스, 폰트, 컬러, 레이아웃 등) 명세
  - [003]_STYLE.md        : 모서리 반경, 테두리 두께, 드롭 섀도우, 테이블 스타일 규격
  - [004]_TEMPLATE.md     : 슬라이드별 컴포넌트 매핑 및 레이아웃 템플릿 맵
                           │
                           ▼
[Step 2] 제1차 100% 무손실 정밀 복제 (Exact 100% Clone)
  - 원본 파일 존재 시: 1:1 바이트 레벨 및 텍스트 토씨 보존 복제
  - 원본 파일 부재 시: Mock Generator가 4대 도서 기반 16장 가상 덱 자체 생성
  -> [001]_제1차_100퍼센트_완전복제본.pptx 생성
                           │
                           ▼
[Step 3] 1:1 정밀 검증 및 3대 갭 분석 (Audit & Gap Analysis)
  - 부족한 부분 (Deficiencies): 외주사 로고 잔재, C-Level 맞춤 에셋 부재 진단
  - 과도한 부분 (Excesses): 미사용 대용량 미디어, 과밀 줄글 텍스트 식별
  - 추가할 요소 (Value-Add): 결론형 액션 타이틀, 골드 메트릭 엠블럼, 12장 피치덱 제안
  -> [005]_1차복제_정밀검증_및_개선제안서.md 생성
                           │
                           ▼
[Step 4] 제2차 고도화 리메이크 실행 (Enhanced 2nd Remake)
  - 외주 로고 및 XML 메타데이터 0% 박멸 (Vector Shape Purge)
  - 1:1 바이트 레벨 도메인 맞춤 에셋 교체 (`ppt/media/`)
  - [006]_제2차_고도화_완결실전본(16장).pptx 생성
  - [007]_제2차_고도화_Executive압축본(12장).pptx 생성
                           │
                           ▼
[Step 5] 품질 무결성 감사 게이트 (Quality Audit Gate)
  - 슬라이드 수 무결성, 파일 용량 적정성, 미치환 플레이스홀더 0건 검증
  -> audit_result.json 발행
                           │
                           ▼
[Step 6] 산출물 번호 체계화 ([001]~[999]) 및 아카이빙
```

---

## 3. 📂 표준 산출물 번호 체계 (`[001]` ~ `[999]`)

```
result/
├── SKILL.md                                              # 유정란 스킬 마스터 명세서 v4.0
├── [001]_BUSINESS_DNA.md                                 # 비즈니스 DNA 및 내러티브 구조
├── [002]_DESIGN.md                                       # 디자인 시스템 8요소 명세서
├── [003]_STYLE.md                                        # 세부 스타일 및 컴포넌트 가이드
├── [004]_TEMPLATE.md                                     # 슬라이드별 템플릿 매핑 명세
├── [005]_1차복제_정밀검증_및_개선제안서.md               # 부족/과도/추가 3대 갭 분석서
├── [008]_1차_2차_복제_종합_베스트프랙티스_가이드.md       # 실전 슬라이드 복제 종합 노하우
├── audit_result.json                                     # 무결성 품질 검증 리포트
│
├── [001]_제1차_100퍼센트_완전복제본.pptx                 # 1차 무손실 100% 복제본
├── [006]_제2차_고도화_완결실전본(16장).pptx              # 2차 외주로고 박멸 완결 실전본
└── [007]_제2차_고도화_Executive압축본(12장).pptx         # 2차 12장 핵심 Executive 압축본
```

---

## 4. 🚀 살아 숨쉬는 자체 내장 실행 스크립트 풀 (`scripts/`)

- `scripts/run_cloner.py`: **[메인 진입점]** 단 한 줄로 6단계를 자동 부화시키는 CLI 러너
- `scripts/extract_4pillars.py`: 4대 설계도서 자동 역공학 파싱 엔진
- `scripts/clone_exact_100.py`: 1차 100% 복제기 (Mock Seed 내장)
- `scripts/audit_and_propose.py`: 1:1 무결성 진단 및 개선 제안서 생성기
- `scripts/remake_enhanced.py`: 2차 외과수술형 고도화 리메이크 엔진
- `scripts/generate_assets.py`: 도메인 맞춤형 고해상도 그래픽 에셋 생성기
- `scripts/verify_quality.py`: 최종 무결성 검증기

---

## 5. 💡 원클릭 부화 실행 명령어

```bash
# 파이썬 가상환경에서 단 한 줄로 전체 파이프라인 자율 구동
python "scripts/run_cloner.py" --output "result"
```
