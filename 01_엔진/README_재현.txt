바름 보장분석 자동화 엔진 — 정본 패키지 (2026.06.08)
========================================================

[이 zip의 목적]
법칙이 창마다 흔들리지 않게, 모든 법칙을 코드(rules.py)에 박은 영구본.
새 창에서 이 zip을 풀어 올리면 동일하게 재현된다.

[구성]
  rules.py              ← 법칙 코드 정본 (Single Source of Truth)
  engine_v2.py          ← 추출·조립 엔진
  법칙_정본.md          ← 법칙 한글 정본 (코드의 거울)
  표준폼_골격.md        ← 16분류 90행 표준 폼 설명
  보장진단_두호.xlsx    ← 완성 결과물 예시
  noname__7_.pdf        ← 테스트 PDF (두호 14계약)
  orig/보장진단_두호.xlsx     ← 영구 폼 (일반암갱신·중입자·116대수술비 행 추가됨)
  orig/보장진단_두호.bak.xlsx ← 동적 행매핑(REMAP) 기준 백업 (지우지 말 것)
  참조/                 ← KB 제안서 (담보 대조용)

[재현 방법]
  1) pip install pymupdf openpyxl python-pptx xlrd --break-system-packages
  2) cp orig/보장진단_두호.xlsx ./보장진단_두호.xlsx   (작업본 복구)
  3) python3 engine_v2.py
  4) 결과 = 보장진단_두호.xlsx
  렌더 검증: soffice --headless --convert-to pdf 보장진단_두호.xlsx

[표는 손대지 말 것]
  폼 구조(칸)는 확정. 입력값만 고객 PDF마다 바뀐다.
  행 추가가 필요하면 orig 폼에만 추가 → 동적매핑이 자동 반영.

[PPT]
  보장분석지(PPT) 자동 채우기는 아직 엔진에 없음 (현재 엑셀까지만 자동).
  PPT 폼(form.pptx)은 별도 보관. PPT 자동화는 다음 단계.

[PPT 생성 — 추가]
  python3 make_ppt.py  →  보장분석지_두호.pptx
  (완성 엑셀의 담보별 비갱신/갱신 합계를 폼 라벨에 자동 기입, 색·슬래시 포함)
  렌더 확인: soffice --convert-to pdf 보장분석지_두호.pptx → pdftoppm
