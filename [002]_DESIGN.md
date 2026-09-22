# [002] DESIGN.md - 디자인 시스템 8요소 명세서

## 1. 캔버스 및 그리드 (Canvas Geometry & Grid)
- **해상도 및 비율**: 16:9 Wide-Screen (13.33" x 7.50")
- **EMU 치수**: 12,192,000 EMU x 6,858,000 EMU
- **안전 마진**: Top 0.8" / Bottom 0.6" / Left 0.8" / Right 0.8"

## 2. 타이포그래피 계층 (Typography Hierarchy)
| 계층 | 폰트 패밀리 | 권장 크기 | 굵기 | 적용 대상 |
|:---|:---|:---|:---|:---|
| **Hero Title** | RecipekoreaOTF Medium / GamtanRoad | 36pt ~ 44pt | Bold | 표지 대제목 |
| **Section Title** | GamtanRoad Dotum Bold / Malgun Gothic | 26pt ~ 30pt | Bold | 슬라이드 상단 액션 타이틀 |
| **Card Subheader** | GamtanRoad Dotum Bold | 18pt ~ 22pt | Bold | 카드 상단 헤더, 표 머리글 |
| **Body Text** | GamtanRoad Dotum Regular / Malgun Gothic | 13pt ~ 15pt | Regular | 본문 설명, 리스트 불릿 |
| **Caption/Badge** | Gwangyang Touching / Dotum | 11pt ~ 12pt | Regular | 차수 배지, 출처, 각주 |

## 3. 컬러 팔레트 시스템 (Color Palette & Semantic Tokens)
| 색상 구분 | HEX 코드 | RGB | 용도 및 의미 |
|:---|:---|:---|:---|
| **Primary Base** | `#1B3B2B` | (27, 59, 43) | 딥 포레스트 그린 (신뢰, 안정, 메인 헤더) |
| **Secondary Brand** | `#3E7B54` | (62, 123, 84) | 세이지 올리브 그린 (카드 헤더, 서브 배지) |
| **Accent Highlight** | `#E5A93C` | (229, 169, 60) | 웜 오커 골드 (수상 엠블럼, 핵심 강조 지표) |
| **Background Dark** | `#0D1F17` | (13, 31, 23) | 표지 및 엔딩 슬라이드 딥 다크 배경 |
| **Background Light** | `#F8FAF8` | (248, 250, 248) | 본문 슬라이드 가독성 최적화 배경 |
| **Surface Card** | `#FFFFFF` | (255, 255, 255) | 3단 카드 박스, 컨테이너 표면 |
| **Text Primary** | `#222222` | (34, 34, 34) | 메인 본문 텍스트 |
| **Text Muted** | `#666666` | (102, 102, 102) | 부가 설명, 각주, 캡션 |

## 4. 레이아웃 아키타입 (Layout Archetypes)
- **Archetype A (표지/엔딩)**: 대제목 + 키 메타데이터 + 중앙 3D 비주얼 + 교수자 프로필
- **Archetype B (3단 분할 카드)**: 3열 균등 카드 배치 (학생창의형 / 기업연계형 / 지역발전형)
- **Archetype C (스케줄 그리드)**: 15주차 강의 계획 2단 분할 테이블
- **Archetype D (프로젝트 쇼케이스)**: 좌측 설명 + 우측 대형 수상작 갤러리 목업

## 5. 컴포넌트 & 카드 시스템 (Component System)
- **Pill Badge**: 둥근 알약형 회차 태그 (`3강`, `Step 01`)
- **Floating Card**: 16px 라운딩 처리된 순백색 정보 컨테이너
- **Metric Highlight**: 수상 내역 및 주요 성과 강조용 골드 하이라이트 박스

## 6. 시각 에셋 및 이미지 스타일 (Visual Assets)
- 고해상도 3D 비주얼 렌더링, 둥근 원형 아바타 마스크, 전문 아이콘 세트

## 7. 여백 및 간격 계층 (Spacing Hierarchy)
- 외부 마진 24pt, 카드 간 간격 16pt, 카드 내부 패딩 14pt

## 8. 깊이감 및 데코레이션 (Depth & Elevation)
- 미세 드롭 섀도우(Opacity 8%, Blur 12px) 및 1px 서브 테두리
