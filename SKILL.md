---
name: slide-style-cloner
description: "Reverse engineer and clone executive presentation styles from any input PDF or PPTX slide deck, extract 4 Pillars (BUSINESS_DNA, DESIGN, STYLE, TEMPLATE), execute 1st 100% exact cloning, audit deficiencies/excesses/value-adds, and generate 2nd enhanced executive presentation slides (30/16-slide Full & 12-slide Executive). Features Surgical In-Place Remake, Vector Logo Purge, and In-Place Media Replacement."
---

# Slide Style Cloner v3.0 (슬라이드 스타일 복제 및 2단계 정밀 리메이크 마스터 스킬)

> **핵심 철학 (Core Philosophy)**:  
> 레퍼런스 템플릿의 **압도적인 3D 그래픽, 세련된 컬러 스키마, 감각적인 카드 레이아웃을 100% 온전히 보존**하면서, 4대 설계도서(BUSINESS_DNA, DESIGN, STYLE, TEMPLATE)를 역공학 추출하고, 1차 무손실 복제 -> 정밀 감사(부족/과도/추가 요소 분석) -> 2차 고도화 리메이크를 수행하는 **'외과수술형 2단계 슬라이드 복제(Two-Stage Surgical Remake)'** 전문 스킬입니다.

---

## 1. 안티패턴 및 핵심 방어 수칙 (What NOT to Do)

1. **절대 백지에서 새로 그리지 마십시오 (Anti-Blank Canvas)**:
   - 템플릿의 억지 틀을 피하겠다고 원본 템플릿 파일(.pptx)을 버리고 빈 캔버스에서 사각형 박스로 새로 그리면, 원본의 수십 MB급 고화질 3D 그래픽과 미학 에셋이 전부 날아가고 조잡한 관공서 상자 PPT가 됩니다.
2. **단순 텍스트 치환에 머물지 마십시오 (Anti-Blind Replacement)**:
   - 표지 상단에 벡터 도형(`GROUP`)으로 숨겨진 외주사(예스폼 등) 로고 및 메타데이터를 방치하지 마십시오.
   - 도메인과 맞지 않는 외국인 스톡 모델 사진이나 IT 사무실 사진을 그대로 두지 마십시오.
3. **단일 릴리즈에 안주하지 마십시오 (Multi-Deck Publishing)**:
   - 30장/16장 완결 실전본(Full Deck)과 12장 Executive 압축본(Pitch Deck)을 동시에 발행하십시오.

---

## 2. 6대 엔드-투-엔드 실행 파이프라인 (6-Step Master Pipeline)

```
[입력: 레퍼런스 PPTX 또는 Mock 가상 템플릿]
                           │
                           ▼
[Step 1] 4대 핵심 설계도서 역공학 추출 (Extract 4 Pillars)
  - [001]_BUSINESS_DNA.md (내러티브 및 비즈니스 로직)
  - [002]_DESIGN.md       (디자인 시스템 8요소 명세)
  - [003]_STYLE.md        (타이포, 컬러, 셰이프 스타일)
  - [004]_TEMPLATE.md     (슬라이드별 컴포넌트 매핑 맵)
                           │
                           ▼
[Step 2] 제1차 100% 무손실 정밀 복제 (Exact 100% Clone)
  - 원본 파일 존재 시: 1:1 바이트 및 텍스트 토씨 보존 복제
  - 원본 부재 시: Mock Generator가 4대 도서 기반 100% 규격 가상 덱 생성
  -> [001]_제1차_100퍼센트_완전복제본.pptx 생성
                           │
                           ▼
[Step 3] 1:1 정밀 검증 및 3대 갭 분석 (Audit & Gap Analysis)
  - 부족한 부분 (Deficiencies): 외주 로고 잔재, C-Level 맞춤 에셋 부재
  - 과도한 부분 (Excesses): 미사용 미디어 비효율, 과밀 텍스트
  - 추가할 요소 (Value-Add): 맞춤 인포그래픽, 12장 Executive 압축 덱
  -> [005]_1차복제_정밀검증_및_개선제안서.md 생성
                           │
                           ▼
[Step 4] 제2차 고도화 리메이크 실행 (Enhanced 2nd Remake)
  - 외주 로고 및 XML 메타데이터 0% 박멸 (Vector Shape Purge)
  - 1:1 바이트 레벨 도메인 맞춤 에셋 교체 (`ppt/media/`)
  - 결론 지향적 Action Title 및 Gold Metric 엠블럼 주입
  -> [006]_제2차_고도화_완결실전본(16/30장).pptx 생성
  -> [007]_제2차_고도화_Executive압축본(12장).pptx 생성
                           │
                           ▼
[Step 5] 무결성 및 품질 감사 게이트 검증 (Quality Audit Gate)
  - 외주 로고 잔재 0건, 플레이스홀더 잔재 0건, 슬라이드 수 무결성 확인
  -> audit_result.json 발행
                           │
                           ▼
[Step 6] 산출물 번호 체계화 ([001]~[999]) 및 아카이빙
```

---

## 3. 표준 산출물 및 번호 체계 (`[001]` ~ `[999]`)

```
result/
├── SKILL.md
├── [001]_제1차_100퍼센트_완전복제본.pptx
├── [001]_BUSINESS_DNA.md
├── [002]_DESIGN.md
├── [003]_STYLE.md
├── [004]_TEMPLATE.md
├── [005]_1차복제_정밀검증_및_개선제안서.md
├── [006]_제2차_고도화_완결실전본(16장).pptx
├── [007]_제2차_고도화_Executive압축본(12장).pptx
├── [008]_1차_2차_복제_종합_베스트프랙티스_가이드.md
└── audit_result.json
```

---

## 4. 원클릭 실행 스크립트 규격

- `src/[010]_extract_four_pillars.py`: 4대 도서 추출기
- `src/[011]_clone_exact_100.py`: 1차 100% 복제기 (Mock 내장)
- `src/[012]_audit_and_propose_improvements.py`: 정밀 감사 및 개선 제안서
- `src/[013]_clone_enhanced_v2.py`: 2차 고도화 리메이크 엔진
- `command/run_pipeline.ps1`: 원클릭 일괄 실행기
