# 🥚 Slide Style Cloner v4.0 (`vd27_slide_style_cloner_v4`)

> 레퍼런스 PPTX의 디자인(그래픽·색·배치)은 그대로 두고, 외주 로고·링크·작성자 정보를 정리한 뒤 새 내용을 넣어 **16장 완결본 + 12장 Executive 압축본**을 만드는 실행형 스킬입니다.

자세한 규칙과 단계 설명은 [SKILL.md](./SKILL.md)를 보세요.

## 설치
```bash
cd vd27_slide_style_cloner_v4
uv pip install -r scripts/requirements.txt
```
(`pip`를 쓰려면 `python -m pip install -r scripts/requirements.txt`. PDF 입력에는 `pymupdf`를 추가로 설치합니다.)

## 실행
```bash
python scripts/run_cloner.py --input "원본.pptx" --output "result_v1" --topic "새 주제"
```

1. 처음 실행하면 `result_v1/`에 설계도서 4종, `[005]` 개선 제안서, 1차 복제본, 2차 덱 2종, `[009]_콘텐츠_매핑표.json`이 생깁니다.
2. `[009]` 매핑표의 `slots`·`replace`·`media`에 새 문구와 사진 경로를 넣습니다.
3. 같은 명령을 다시 실행하면 서식은 그대로, 내용만 바뀐 2차 덱이 만들어집니다.

## 폴더 구조
```
vd27_slide_style_cloner_v4/
├── README.md
├── SKILL.md
└── scripts/
    ├── run_cloner.py          # 전체 실행 (Step 1~6)
    ├── extract_4pillars.py    # Step 1 원본 분석 → 4대 설계도서
    ├── clone_exact_100.py     # Step 2 바이트 복제 + SHA-256 확인
    ├── audit_and_propose.py   # Step 3 갭 분석 + 매핑표
    ├── generate_assets.py     # 엠블럼/배지 PNG
    ├── remake_enhanced.py     # Step 4 2차 리메이크
    ├── verify_quality.py      # Step 5 품질 검사
    ├── cloner_common.py       # 공통 함수
    ├── mock_seed.py           # 원본이 없을 때 쓰는 16장 시드
    └── requirements.txt
```
