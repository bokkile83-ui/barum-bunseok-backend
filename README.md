# 바름 AI 보장분석 백엔드 (단일 Render 서비스)

PDF 보장분석지 → Claude(29+4조) → 색칠된 보장진단 엑셀 → (2단계: PPT 연계) → 채팅에 다운로드.

## Render 배포
1. 이 폴더를 GitHub 새 repo에 올린다.
2. render.com → New → Web Service → 그 repo 선택.
3. Build: `pip install -r requirements.txt`  /  Start: `uvicorn main:app --host 0.0.0.0 --port $PORT`
4. Environment → `ANTHROPIC_API_KEY` = sk-ant-... 추가.
   + 잠금 켜려면 `ACCESS_PW` = 1009 (원하는 비번)도 추가. 미설정 시 잠금 없이 누구나 사용.
5. 배포 후 나온 주소 열기 → PDF 업로드.

## 파일
- main.py        : FastAPI(페이지+분석+채움+파일반환)
- rules.py       : 29+4조 시스템프롬프트
- fill_excel.py  : JSON→색칠 엑셀 (검증완료)
- template.xlsx  : 빈 보장진단 템플릿
- static/index.html : 채팅 UI
- (2단계) fill_pptx.py + form.pptx : 엑셀값→PPT 폼 연계

## 2단계(PPT) 추가법
form.pptx(빈 폼) 등록 + fill_pptx.py(엑셀값→버블/메모) 추가하면 자동으로 PPT도 채팅에 같이 올라옴.
