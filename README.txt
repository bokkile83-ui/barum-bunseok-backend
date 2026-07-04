BARUM 보장분석 — 최종 산출물 (2026.07.04)

[1_code/]  ← Railway/GitHub 배포용 완전체
  main.py               = v30t-brainheart-20260704 (회귀 11/11 통과)
                          ※ 베이스 v30r + 버전태그. 현대 뇌혈관Ⅰ·Ⅱ·Ⅲ 이미 구현(_rmn 괄호오검출 차단)
  report_weasy.py       = 보장설명서 8p PDF 생성기 (P5 산정특례·순환계 컬럼 포함)
  master.xlsx           = 99행 정본 (빈맥40·산정특례심장45·산정특례뇌혈관35·외상성뇌출혈34)
  ppt_form.pptx         = PPT 폼
  coverage_benchmark.py = 충족률 벤치마크

[2_report/]
  BARUM_report_kimjingu_final_8p.pdf = 보장설명서 최종본 8p
    표지(담당줄삭제)/보장현황/AI진단/CI선지급80%/부위별충족률/
    담보별보장범위(각축 개별·산정특례 기준박스)/주요치료비변천사/상담워크시트

[3_guide/]
  guide_v30t.md          = 오늘까지 전체 확정 통합정본 (정본)
  guide_v30m_detail.md   = v30m 상세정본 (참고)
  master_sample_99row.xlsx

■ 오늘(07.04) 확정:
  · 뇌·심 담보축 = 각각 개별·각각 보상 (뇌 3축/심 4축)
  · 묶음=구성질환 중 1개 보상 / 묶음 외 단독=개별보상
  · 심장 특정Ⅰ·Ⅱ·Ⅲ = 회사별 범위 습득(고정 아님)
  · 현대 뇌혈관Ⅱ→뇌졸증 / Ⅰ→뇌혈관진단비 (괄호수식어 오검출 차단)
  · 빈맥=심장포함·전용행40 / 산정특례·순환계·외상성뇌출혈=각각 단독
  · 산정특례 코드범위 확정 / 지급조건(30일·5%)=[확인]

■ 배포: GitHub main 현재본 확인 → v30t 덮기 → Railway Deploy → /health로 v30t 확인
  보장분석실: https://web-production-4a155.up.railway.app
  /health:    https://web-production-4a155.up.railway.app/health
  GitHub:     https://github.com/bokkile83-ui/barum-bunseok-backend
