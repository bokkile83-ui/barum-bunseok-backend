# BARUM 배포 번들 (v14-eq2-20260619)

## 이번 핵심 수정 (PPT≠엑셀 해결)
- read_excel_totals: 끝열 =SUM() 캐시(LibreOffice) 의존 제거 → 데이터셀 직접 합산. 재계산 실패해도 항상 PPT=엑셀.
- 직전 driver2 누적분 포함: 대인/대물 분리, 실손 항상 파랑, 질병수술비 위 구분선, PPT 이름박스 고정(auto_size NONE).

## ※ 보류
- 자부상 '12-14급 기준 합산' 규칙은 확정 전(현재 전체 합산). 채팅 답변 후 반영.

## 업로드: main.py / MASTER_보장분석_엑셀_영구표본.xlsx(필수) / PPT 2종 / requirements·Procfile·nixpacks
## 확인: /health → "v14-eq2-20260619"
## 지침: 지침_통합정본_20260619.md → 프로젝트 지식 업로드(대체)
