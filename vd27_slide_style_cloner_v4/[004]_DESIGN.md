<!--
=============================================================================
REVISION HISTORY TABLE
=============================================================================
REV NO  | DATE        | AUTHOR          | DESCRIPTION OF CHANGE
-----------------------------------------------------------------------------
REV.01  | 2026-09-22  | Antigravity AI  | 최초 작성 (3주차 캡스톤디자인 강의 PPT 디자인 시스템 역공학 분석 명세)
=============================================================================
-->

# [004] DESIGN.md - 슬라이드 디자인 시스템 역공학 명세서

## 1. 캔버스 규격 (Canvas Geometry)
- **종횡비**: 16:9 Wide-Screen
- **물리적 크기**: 13.333 inches × 7.500 inches (33.867 cm × 19.050 cm)
- **EMU 좌표계**: 12,192,000 EMU × 6,858,000 EMU
- **안전 여백 (Safe Margins)**: 
  - Top: 0.8" (731,520 EMU)
  - Bottom: 0.6" (548,640 EMU)
  - Left / Right: 0.8" (731,520 EMU)

---

## 2. 타이포그래피 시스템 (Typography Hierarchy)

| 레벨 | 폰트 패밀리 | 기본 크기 | 굵기 | 주요 용도 |
| :--- | :--- | :--- | :--- | :--- |
| **Hero Title** | RecipekoreaOTF Medium / GamtanRoad Gamtan | 36pt ~ 44pt | Bold | 표지 대제목, 메인 테마 |
| **Section Title** | GamtanRoad Dotum Bold | 28pt ~ 32pt | Bold | 슬라이드 상단 액션 타이틀 |
| **Sub Header** | GamtanRoad Dotum Bold | 18pt ~ 22pt | Bold | 카드 헤더, 섹션 소제목, 테이블 헤더 |
| **Body (본문)** | GamtanRoad Dotum Regular / +mn-lt | 14pt ~ 16pt | Regular | 설명 텍스트, 리스트 본문, 표 내용 |
| **Kicker / Badge** | Gwangyang Touching | 14pt ~ 16pt | Regular | 우측 상단 회차 표시 (`3강`), 카테고리 태그 |
| **Caption / Footnote** | GamtanRoad Dotum Regular | 10pt ~ 12pt | Regular | 출처, 참고문헌, 각주 (`*출처:...`) |

---

## 3. 컬러 팔레트 시스템 (Color Palette)

| 구분 | 색상 코드 (HEX) | 역할 및 적용 요소 |
| :--- | :--- | :--- |
| **Primary Base** | `#1B3B2B` (Forest Deep Green) | 숲/자연/에코 테마의 신뢰감 있는 딥 그린, 표지 및 헤더 강조 |
| **Secondary Brand** | `#3E7B54` (Sage Olive Green) | 서브 타이틀, 카드 헤더 배경, 강조 배지 |
| **Accent Gold** | `#E5A93C` (Warm Ochre Gold) | 핵심 수치, 우수상 엠블럼, 포인트 아이콘 |
| **Background Dark** | `#0D1F17` (Deep Dark Green) | 표지 및 엔딩 슬라이드 딥 다크 배경 |
| **Background Light**| `#F8FAF8` (Soft Off-White) | 본문 콘텐츠 슬라이드 배경 (가독성 최적화) |
| **Surface Card** | `#FFFFFF` (Pure White) | 3열 카드 박스, 테이블 셀 배경 (그림자 효과 포함) |
| **Text Primary** | `#222222` (Charcoal Black) | 본문 주요 텍스트 |
| **Text Muted** | `#666666` (Slate Gray) | 부가 설명, 각주, 테두리선 |

---

## 4. 슬라이드 레이아웃 구조 (Layout Archetypes)

1. **표지 레이아웃 (Cover Deck)**:
   - 좌측 상단 강좌 대제목 + 우측 상단 메타데이터(연도, 날짜, 시간)
   - 중앙 비주얼 일러스트레이션/그래픽 + 하단 교수자/소속 프로필 정보
2. **테이블 그리드 레이아웃 (Schedule Grid)**:
   - 15주차 강의 계획을 2개 테이블(1~8주차, 9~15주차)로 분할 배치
   - 상단 헤더 행(주차, 주제, 수업내용, 과제/비고) 구분 강조
3. **3단 분할 카드 레이아웃 (3-Column Card Architecture)**:
   - 3개의 동일 비율 카드 박스 (학생창의형 / 기업연계형 / 지역발전공익형)
   - 카드 상단 타이틀 배지 + 중간 특징/내용/목표 불릿 블록
4. **프로젝트 갤러리 카드 레이아웃 (Project Case Cards)**:
   - 슬라이드 상단 회차/과목명 바 + 중앙/우측 대형 프로젝트 목업 및 수상작 쇼케이스
