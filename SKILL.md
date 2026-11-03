---
name: slide-style-cloner
description: "Reverse engineer and clone executive presentation styles from any input PDF or PPTX slide deck, extract 4 Pillars (BUSINESS_DNA, DESIGN, STYLE, TEMPLATE), execute 1st 100% exact cloning, audit deficiencies/excesses/value-adds, and generate 2nd enhanced executive presentation slides (16-slide Full & 12-slide Executive). Features Surgical In-Place Remake, Vector Logo Purge, and In-Place Media Replacement. Self-contained with executable scripts for zero-dependency instant generation."
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
   - 파일 속성(`docProps/core.xml`)에 남은 외주 플랫폼 원작자 정보도 함께 정리하십시오.
   - 도메인과 맞지 않는 외국인 스톡 모델 사진이나 어색한 사무실 사진을 방치하지 마십시오.
3. **단일 덱 발행에 안주하지 마십시오 (Multi-Deck Strategy)**:
   - 심층 발표·심사용 16장 완결 실전본과 임원 보고용 12장 Executive 압축본을 항상 쌍으로 발행하십시오.
4. **이미지를 지우고 새로 배치하지 마십시오 (In-Place Media Replacement)**:
   - 사진·그래픽을 바꿀 때는 `ppt/media/image*.png` 파일만 같은 이름으로 교체하십시오. 그래야 원본의 둥근 모서리 마스크, 그림자, 위치 좌표가 1mm도 깨지지 않습니다.
   - 로고 등 불필요한 객체는 `sp_elem.getparent().remove(sp_elem)`처럼 해당 도형만 제거하고, 나머지 레이아웃은 건드리지 마십시오.

---

## 2. 🐣 6단계 작업 파이프라인 (Internal Lifecycle)

1. **[Step 1] 4대 설계도서 추출**:
   - `[001]_BUSINESS_DNA.md`, `[002]_DESIGN.md`, `[003]_STYLE.md`, `[004]_TEMPLATE.md`를 원본에서 역공학으로 추출합니다.
2. **[Step 2] 제1차 100% 무손실 복제**:
   - 레퍼런스 파일이 있으면 무손실 복제, 없으면 내장 Seed로 16장 가상 덱을 생성합니다.
3. **[Step 3] 1:1 정밀 감사 및 3대 갭 진단**:
   - 부족한 점(외주 로고 잔재), 과도한 점(미사용 미디어, 빽빽한 텍스트), 추가할 점(맞춤 엠블럼, Action Title, 12장 압축본)을 `[005]` 보고서로 작성합니다.
4. **[Step 4] 제2차 고도화 리메이크 (Surgical In-Place Remake)**:
   - 외주사 벡터 로고 제거 + 도메인 맞춤 에셋 1:1 교체 + 16장 완결본 & 12장 Executive 압축본을 동시에 빌드합니다.
5. **[Step 5] 품질 무결성 검증 (Quality Audit Gate)**:
   - 슬라이드 수, 파일 크기, 플레이스홀더 잔재 0건을 확인해 통과 시 `[PASS]`를 `audit_result.json`에 기록합니다.
6. **[Step 6] 산출물 번호 체계화 및 정리**:
   - 아래 3장의 `[001]`~`[999]` 번호 체계로 저장하고, `[000]_산출물_목록.md`에 파일별 설명과 크기를 기록합니다.

---

## 3. 📂 표준 산출물 번호 체계 (`[001]` ~ `[999]`)

```
result/
├── SKILL.md                                              # 유정란 스킬 마스터 명세서 v4.0
├── [001]_BUSINESS_DNA.md                                 # 비즈니스 DNA 및 내러티브 구조
├── [002]_DESIGN.md                                       # 디자인 시스템 8요소 명세서
├── [003]_STYLE.md                                        # 세부 스타일 및 컴포넌트 가이드
├── [004]_TEMPLATE.md                                     # 슬라이드별 템플릿 매핑 명세
├── [000]_산출물_목록.md                                  # 이번 실행 산출물 목록 (자동 생성)
├── [005]_1차복제_정밀검증_및_개선제안서.md               # 부족/과도/추가 3대 갭 분석서
├── [008]_1차_2차_복제_종합_베스트프랙티스_가이드.md       # 실전 슬라이드 복제 종합 노하우
├── [009]_콘텐츠_매핑표.json                              # 새 문구·이미지를 넣는 표 (사용자가 채움)
├── audit_result.json                                     # 무결성 품질 검증 리포트
├── pillars.json / audit_plan.json                        # 단계 사이에 주고받는 분석 원자료
├── assets/                                               # 자체 엠블럼·배지 PNG
│
├── [001]_제1차_100퍼센트_완전복제본.pptx                 # 1차 무손실 100% 복제본
├── [006]_제2차_고도화_완결실전본(16장).pptx              # 2차 외주로고 박멸 완결 실전본
└── [007]_제2차_고도화_Executive압축본(12장).pptx         # 2차 12장 핵심 Executive 압축본
```

---

## 4. 🚀 자체 내장 실행 스크립트 풀 (`scripts/`)

- `scripts/run_cloner.py`: **[메인 진입점]** Step 1~6을 차례로 실행하는 CLI 러너. 실패하면 어느 단계에서 멈췄는지 알려 줍니다.
- `scripts/extract_4pillars.py`: 원본의 테마 색·글꼴, 실제 사용 색·글꼴·글자 크기, 반복 머리글, 슬라이드 유형, 글자 슬롯을 분석해 4대 설계도서 작성
- `scripts/clone_exact_100.py`: 1차 바이트 단위 복제 + SHA-256 동일성 확인
- `scripts/audit_and_propose.py`: 무결성 비교, 로고·링크·작성자·안내문구·대용량/미사용 미디어 진단, Executive 슬라이드 선별, `[009]` 매핑표 생성
- `scripts/remake_enhanced.py`: 로고·외부 링크 제거, 서식 유지 문구 교체, 제자리 이미지 교체, 대용량 이미지 4K 축소, 완결본 + Executive 압축본 저장
- `scripts/generate_assets.py`: 테마 색으로 주제 엠블럼·배지 PNG 생성
- `scripts/verify_quality.py`: 산출물을 다시 열어 슬라이드 수·로고 잔재·외부 링크·외주 문구·안내 문구·작성자 정보를 검사해 `audit_result.json` 기록
- `scripts/cloner_common.py`, `scripts/mock_seed.py`: 공통 함수, 원본이 없을 때 쓰는 16장 Mock 시드
- `scripts/requirements.txt`: 필요한 패키지 (python-pptx, Pillow, lxml / PDF는 pymupdf 추가)

> **⚠️ 알아 둘 점**:
> - **로고 자동 삭제 기준**: 외주·워터마크 문구(예스폼, www. 등)가 들어간 도형, 외부 링크가 걸린 작은 로고형 도형만 자동으로 지웁니다. 모서리의 작은 이미지는 학교·회사 자체 로고일 수 있어 `[005]`에 "확인 필요"로만 표시합니다. 지우려면 `[009]`의 `purge`에 넣으세요.
> - **새 내용은 자동으로 지어내지 않습니다.** 새 문구와 사진은 `[009]_콘텐츠_매핑표.json`에 넣어야 반영됩니다. `--topic`을 주면 표지 제목(덱 제목) 문구만 주제로 바꿉니다.
> - **PDF 입력**은 페이지를 그림으로 바꿔 넣으므로 글자 편집은 안 됩니다 (`pymupdf` 설치 필요). `.ppt`는 `.pptx`로 다시 저장해서 넣어야 합니다.
> - 이 PC에는 PowerPoint가 없어서, 결과 화면은 파일 구조 검사로만 확인했습니다. 발표 전에 PowerPoint로 한 번 열어 보세요.
> - `result/` 폴더에 있는 기존 `audit_result.json`은 예전 파일(`[002]`/`[003]`_AI_에이전트_캡스톤디자인)을 검사한 결과입니다. 파이프라인을 다시 실행하면 새로 써집니다.

---

## 5. 💡 원클릭 부화 실행 명령어

```bash
# 파이썬 가상환경에서 단 한 줄로 전체 파이프라인 자율 구동
python "scripts/run_cloner.py" --input "<레퍼런스_PPTX>" --output "result" --topic "<도메인_주제>"
```

- `--input`: 레퍼런스 PPTX 경로. 생략하면 현재 폴더, `upload/`, `input/`에서 PPTX를 찾고, 없으면 내장 Seed를 사용합니다.
- `--output`: 산출물 폴더 (기본값 `result`)
- `--topic`: 새 도메인 주제. 주면 표지 제목 문구와 파일 속성 제목을 이 주제로 바꿉니다.
- `--content`: 콘텐츠 매핑표 경로 (기본: 출력 폴더의 `[009]_콘텐츠_매핑표.json`)
- `--exec-count`: Executive 압축본 장수 (기본 12)
- `--no-compress`: 대용량 이미지를 줄이지 않고 원본 화질 그대로 둠

### 새 내용 넣는 순서
1. 위 명령을 한 번 실행합니다. 출력 폴더에 `[009]_콘텐츠_매핑표.json`이 생깁니다.
2. 매핑표의 `slots` → 각 슬롯의 `new`에 새 문구를 넣습니다 (`original`이 지금 문구). 덱 전체에서 바꿀 말은 `replace`, 바꿀 사진은 `media`(예: `"image3.jpeg": "C:/사진/새사진.jpg"`)에 넣습니다.
3. 같은 명령을 다시 실행합니다. 매핑표는 덮어쓰지 않고 그대로 반영됩니다.

> 출력 폴더를 기존 `result/`로 두면 같은 이름의 `[001]`·`[006]`·`[007]` PPTX를 덮어씁니다. 기존 결과물을 남기려면 `--output`에 새 폴더를 지정하세요.
