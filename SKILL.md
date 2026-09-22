---
name: slide-style-cloner
description: "Reverse engineer and clone executive presentation styles from any input PDF or PPTX slide deck, extract 4 Pillars (BUSINESS_DNA, DESIGN, STYLE, TEMPLATE), execute 1st 100% exact cloning, audit deficiencies/excesses/value-adds, and generate 2nd enhanced executive presentation slides (30/16-slide Full & 12-slide Executive). Features Surgical In-Place Remake, Vector Logo Purge, and In-Place Media Replacement. Self-contained with executable scripts for zero-dependency instant generation."
---

# Slide Style Cloner v4.0 (살아 숨쉬는 완전 실행형 유정란 스킬)

> **🥚 유정란 스킬 선언문 (The Living Skill Manifesto)**:  
> 본 스킬은 단순한 텍스트 지침서(무정란)가 아닙니다. **자체 내장된 실행 엔진(`scripts/run_cloner.py`)과 역공학 추출기, 템플릿 복제기, 무결성 검증기를 품고 있어, 사용자의 단 한 줄 명령에도 스스로 부화하여 100% 완결된 산출물을 낳는 완전 자립형 유정란(Fertilized Skill)**입니다.

---

## 1. 🐣 스스로 부화하는 6대 자율 파이프라인

에이전트는 사용자가 PPT 복제나 스타일 추출을 요청할 때, 번거로운 백지 코딩 없이 아래 내장 스크립트를 즉시 실행하여 산출물을 부화시킵니다:

```bash
# 원클릭 자체 부화 실행 (파이썬 가상환경에서 즉시 구동)
python "C:\Users\user\.gemini\config\skills\slide-style-cloner\scripts\run_cloner.py" --input "<레퍼런스_PPTX>" --output "result" --topic "<도메인_주제>"
```

### 자율 구동 6대 단계 (Internal Lifecycle):
1. **[Step 1] 4대 설계도서 자동 산란**:
   - `[001]_BUSINESS_DNA.md`, `[002]_DESIGN.md`, `[003]_STYLE.md`, `[004]_TEMPLATE.md` 동시 추출.
2. **[Step 2] 제1차 100% 무손실 복제**:
   - 레퍼런스 파일이 있으면 바이트 단위 무손실 복제, 없으면 내장 Seed로 16장 가상 덱 생성.
3. **[Step 3] 1:1 정밀 감사 및 3대 갭 진단**:
   - 부족한 점(외주 로고 잔재), 과도한 점(미사용 미디어), 추가할 점(맞춤 엠블럼/피치덱) 자동 보고서 작성.
4. **[Step 4] 제2차 고도화 리메이크 (Surgical In-Place Remake)**:
   - 외주사 벡터 로고 0% 영구 삭제 + 1:1 도메인 맞춤 에셋 교체 + 16장 완결본 & 12장 Executive 피치덱 동시 빌드.
5. **[Step 5] 품질 무결성 자동 검증 (Quality Audit Gate)**:
   - 슬라이드 수, 파일 크기, 플레이스홀더 잔재 0건 통과 시 `[PASS]` 부여.
6. **[Step 6] 산출물 번호 체계화 ([001]~[999]) 및 자동 아카이빙**.

---

## 2. 🗂️ 자체 내장 자산 구조 (Self-Contained Bundle)

```
slide-style-cloner/
├── SKILL.md                          # 유정란 스킬 마스터 헌법
├── scripts/                          # 🧬 살아 숨쉬는 내장 실행 유전자 (Scripts)
│   ├── run_cloner.py                 # [메인 진입점] 원클릭 자율 부화기
│   ├── extract_4pillars.py           # 4대 설계도서 역공학 추출 엔진
│   ├── clone_exact_100.py            # 제1차 100% 무손실 복제 엔진 (Mock Fallback 내장)
│   ├── audit_and_propose.py          # 1:1 정밀 감사 및 3대 개선 제안서 생성기
│   ├── remake_enhanced.py            # 제2차 외과수술형 고도화 리메이크 엔진
│   ├── generate_assets.py            # 도메인 맞춤형 엠블럼/카드 에셋 생성기
│   └── verify_quality.py             # 무결성 검증 게이트
│
├── references/                       # 📚 공인 레퍼런스 설계 규격
│   ├── BUSINESS_DNA_EXAMPLE.md
│   ├── DESIGN_EXAMPLE.md
│   ├── STYLE_EXAMPLE.md
│   └── TEMPLATE_EXAMPLE.md
│
└── examples/                         # 💡 실전 산출물 예시 모음
```

---

## 3. 🛡️ 안티패턴 방어 규약

1. **백지 재작성 금지 (Anti-Blank Canvas)**: 원본의 수십 MB급 3D 그래픽/섀도우 레이아웃을 훼손하지 말 것.
2. **외주사 로고 방치 금지 (Vector Logo Purge)**: 슬라이드 내 `GROUP` 도형 및 메타데이터를 영구 삭제할 것.
3. **단일 덱 발행 금지 (Multi-Deck Strategy)**: 16/30장 완결 실전본과 12장 Executive 피치덱을 항상 쌍으로 발행할 것.
