# BARUM 보장분석 백엔드 v2 (Railway 드롭인)
PDF → PyMuPDF 추출 → 캐논 분류(색상·CI) → 보장진단 엑셀 + 보장요약 PPT (zip 반환)

## 파일
- main.py            FastAPI(업로드UI + /analyze, 비번 env BARUM_PW 기본1009)
- barum_extract.py   PyMuPDF 추출엔진(★poppler 금지)
- barum_dict.py      담보 기본지식 + CANON(17카테고리75행) + is_ci
- barum_canon.py     담보명 → 캐논 표준담보명 분류기
- fill_canon.py      캐논 template 채움 + 색상(갱신=파랑/비갱신=검정/미가입=빨강) + CI=보라
- ppt_gen.py         채운 엑셀 → 보장요약 PPT(네이비+골드)
- template_canon.xlsx / canon_rows.json   캐논 폼 + 담보명→행 인덱스
- requirements.txt / Procfile / railway.json

## 배포(Railway)
1. 이 폴더 내용을 repo 루트에 그대로 (폴더로 감싸지 말 것)
2. Railway → repo 연결 → Deploy  3. (선택) 변수 BARUM_PW
4. URL 접속 → 비번 + PDF → 보장분석.zip(엑셀+PPT)

## 색상 규칙(법칙13)
갱신담보=파랑글자 · 비갱신=검정 · 미가입(합계0)=빨강 · CI결합담보=보라+메모"CI중대한(선지급)"

## 한계(v2)
- 캐논에 행 없는 담보(심장수술·뇌혈관수술·전이암 등)=보류시트로(추측금지)
- 갱신색은 계약레벨 아닌 담보 g플래그 기준 → 일부 갱신계약 검정 가능
- PPT=요약덱(네 11p 정식폼 아님). 정식폼 매핑은 별도
- 분류 미매핑 잔존(니치 특약) — 실제 PDF 별첨 추가시 사전 확장
