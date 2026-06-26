# BARUM v29 전체 자료 (2026.06.26)

## 정본
- 지침_v29.md  ← 통신·매핑·실손세대·7개개선·체크리스트 정본

## code/
- coverage_benchmark.py  ← 충족률 엔진 + map_excel_to_report (검증완료)
- report_weasy.py        ← weasyprint 리포트(chromium불필요). page3 충족률 설명표 추가본
- 질병수술비_패치.md      ← 50→750 버그 패치(미적용, main.py에 박을 코드)
- main_current_v27asc.py ← 프로젝트 현행 main.py (v27-asc, 1312줄)
- main_handoff.py        ← 직전 핸드오프 main.py

## assets/
- MASTER_v29.xlsx        ← 빈맥 신설 마스터
- MASTER_v28_90행.xlsx   ← 현행 정본 90행
- cover_form.pptx, 로고 3종

## samples/
- 정기철_auto.pdf        ← 동적 map_excel_to_report 검증 렌더(3p)
- 정기철_weasy.pdf       ← 직전 하드코딩 렌더
- 보장진단_정기철.xlsx    ← 원천 엑셀
- p1~p3.png

## 상태
✅ 완료: 지침v29, 충족률엔진, report_weasy page3설명, 정기철 검증
★ 미적용(코드 박기 남음):
  1) 질병수술비 패치(code/질병수술비_패치.md)
  2) 실손 세대별·통합암·통합전이암칸·항암분리·묶음후유분리(지침 v29 §8)
  3) 리포트 4페이지 압축 + 보장설명지 신규섹션
  4) map_excel_to_report main.py 통합 + weasyprint requirements/nixpacks
  5) 배포 v27b→v29 (GitHub main.py 버전문자열 갱신 필요)

## 배포
Railway: web-production-4a155.up.railway.app  /health
GitHub: github.com/bokkile83-ui/barum-bunseok-backend

## ★추가(2026.06.26)
- code/main_v29patch_질병수술비.py ← 현행 main.py + 질병수술비 750버그 패치 적용본(문법검증 완료, 런타임 미검증).
  변경: 질병수술비 행 = 변형/종(1-5종)/대수술/일당/Ⅱ·Ⅲ 전부 제외 → 합산 누수 차단(과합산 방지, 의심분은 매핑 제외).
  ※ 실손 세대·통합암·통합전이암·항암분리·묶음후유 등 나머지 v29는 코드 미반영 → 지침_v29.md 보고 수기.

## ★배포 방법 (v29로 뜨게)
1. code/main.py 를 GitHub repo 루트 main.py 로 덮어쓰기(업로드).
   (버전문자열 v29-realson-20260626 포함 → /health 가 v29로 응답)
2. Railway 자동 재배포 대기 후 https://web-production-4a155.up.railway.app/health 확인.
3. {"ok":true,"version":"v29-realson-20260626"} 나오면 성공.
※ v27b로 계속 뜨면: GitHub에 main.py가 실제로 갱신됐는지 + Railway가 그 커밋을 배포했는지 확인.
