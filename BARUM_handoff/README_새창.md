# 바름(BARUM) 보장리포트 — weasyprint 통합 핸드오프
정리 2026.06.26 · 새 창에서 이어가기용

## 지금까지 한 것 (완료)
1. 마스터 엑셀 빈맥(I47·48) 행 신설 → MASTER_v29.xlsx (91행)
2. weasyprint로 chromium 없이 PDF 리포트 생성 검증 완료 (3페이지)
   - p1 보장현황(10영역) / p2 진단+갱신구조+보험료막대 / p3 부위별 충족률 도넛
   - assets/code/report_weasy.py = 작동하는 함수 (build_report_pdf)
   - samples/정기철_weasy.pdf = 실제 렌더 결과

## weasyprint 핵심 (검증됨)
- pip install weasyprint (v69)
- 시스템 의존성: pango, fontconfig, 나눔/Noto CJK 폰트
- 주의: CSS grid·float 누적 깨짐 → table 레이아웃 사용. SVG 도넛은 정상 렌더됨.

## 남은 작업 (새 창에서 할 것)
### ★1순위: map_excel_to_report(data, ppt_totals) 함수
완성 엑셀 끝열 합계 → 리포트 dict 자동 매핑.
현재 report_weasy.py는 정기철 값이 하드코딩(sample dict)되어 있음.
이걸 동적 매핑으로 바꿔야 어떤 고객 txt든 자동 생성됨.
필요 매핑:
  - coverage[]: 카테고리별 담보 → status(full/part/gap) 판정
  - donuts[]: 영역별 충족률 % (산정 로직 미확정 — 현재 구성 충실도 상대값)
  - strength[]/weak[]: 충실/취약 영역 자동 추출
  - premium_bars[]/renew_list[]: 계약별 보험료·갱신여부 (data['contracts']에서)

### 2순위: main.py /analyze 통합 (3줄)
build_ppt 다음에:
  from report_weasy import build_report_pdf, map_excel_to_report
  rep = map_excel_to_report(data, ppt_totals)
  rep.update({'branch':지점명,'manager':지점장명,'title':직책,'phone':휴대폰})
  rp = os.path.join(d, f'보장리포트_{cust}.pdf'); build_report_pdf(rep, rp)
  response['report_b64']=base64.b64encode(open(rp,'rb').read()).decode()
  response['report_name']=f'보장리포트_{cust}.pdf'

### 3순위: 배포 설정
requirements.txt 에 추가: weasyprint
nixpacks.toml:
  [phases.setup]
  nixPkgs = ["pango","fontconfig","nanum-gothic-coding"]

## 변수 (표지·푸터 공용)
{{고객명}} {{지점명}} {{지점장명}} {{직책}} {{휴대폰}}

## 정본 원칙 (유지)
- 등식2: PPT/리포트는 완성 엑셀만 읽음
- 추측금지: 권장 가입액·상품 단정 안 함 → 공백·방향만
- 제외 4종: 실효·미납해지·농업인NH·자동차
- 빈맥 = v29 신설 (구 '무행, 묶음제외' 폐기)

## 파일
code/report_weasy.py  ← 작동 함수 (이거 그대로 프로젝트에 추가)
code/main_current.py  ← 현재 배포 main.py (통합 대상)
assets/MASTER_v29.xlsx ← 정본 마스터 (빈맥 신설)
samples/정기철_weasy.pdf ← 목표 결과물
