# BARUM 보장분석 백엔드 (Railway)

PDF 1개 업로드 → Claude 네이티브 PDF 분석(법칙 내장) → 계약별 색칠 엑셀 + 보장분석지 PPT ZIP.

## 배포 (Railway)
1. 이 폴더 전체를 repo(bokkile83-ui/barum-bunseok-backend)에 push 또는 Railway에 업로드.
2. 환경변수(Variables) 설정:
   - `ANTHROPIC_API_KEY` = sk-ant-... (필수)
   - `ACCESS_PW` = 1009 (게이트 비번, 기본 1009)
   - `MODEL` = claude-sonnet-4-6 (선택, 정확도↑면 claude-opus-4-7)
3. 시작 명령은 Procfile 자동: `uvicorn main:app --host 0.0.0.0 --port $PORT`
4. 배포 URL 접속 → 비번 입력 → PDF 업로드 → ZIP 다운로드.

## 파일
- main.py : FastAPI (/ 업로드폼, /analyze 처리)
- rules.py : 법칙 + JSON 스키마 (Claude 시스템 프롬프트)
- fill_excel.py : JSON → 계약별 그리드 엑셀 (색·콤마·선·심장 중복채움·메모)
- fill_pptx.py : 한장보장 폼(form.pptx) 합계 채움
- template.xlsx / form.pptx : 표본
- static/index.html : 업로드 게이트

## 핵심
- Claude가 PDF를 직접 읽어 계약별 세로 컬럼에 정확히 매핑(좌표추출 한계 없음).
- 심장 허혈성→급성심근경색·협심증 중복보장 자동 채움.
- 1~5종 종별 합산, n대·통합전이암 대표1개, 하이클래스=비급여주요치료비.
- 결과는 증권 최종확인.
