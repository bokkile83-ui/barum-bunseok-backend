# v29 보장설명지 배포 (루트에 올릴 파일)
아래 5개를 GitHub repo "루트"에 올린다(폴더 만들지 말 것):
- main.py            ← v29 + 질병수술비패치 + 보장설명지 생성 훅
- report_weasy.py    ← 리포트 PDF 렌더(main.py가 import)
- coverage_benchmark.py ← 충족률 엔진(main.py가 import)
- requirements.txt   ← weasyprint 추가됨
- nixpacks.toml      ← weasyprint 시스템 의존성(pango/폰트). ★없으면 리포트 생성 실패

※ master.xlsx / ppt_form.pptx / chiryo_form.pptx 는 기존 repo에 이미 있으니 그대로 둔다.
배포 후 /health = v29-realson-20260626, 분석 시 4번째 파일 '보장설명지_OOO.pdf' 생성.
리포트 실패해도 엑셀+PPT 2개는 정상(비치명적 처리).
