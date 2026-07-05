BARUM 최종본 (2026.07.05) = 어제 v31n-heart-20260705 + 오늘 report_pptx

★중요: 베이스 = v31n-heart-20260705 (어제 진짜 최종). 
  오늘 세션서 report_pptx.py(설명지=진단서PPT 일치)만 추가.

[1_code/]  ← 배포용
  main.py               v31n-heart-20260705 (어제 최종, v28 canon 정렬)
  report_weasy.py       보장설명지 PDF (CI분기 포함)
  report_pptx.py        ★오늘 신규★ 보장진단서 PPT (설명지와 동일 rep=100% 일치)
  coverage_benchmark.py rep 생성
  master.xlsx / ppt_form.pptx

[2_sample/] 주정은 설명지PDF·분석지PPT·엑셀 (v31n-heart-20260705 산출)
[3_guide/]  지침_정본_v31n-heart-20260705 = 정본 / 프로젝트_업로드_안내.md

■ v31n-heart-20260705 = v28 canon 기준 (어제 정렬):
  · 빈맥 = 마스터 무행·[확인] (v28정본) + 심장6사 묶음엔 지점장 7/1 지시로 포함
  · 전이암진단비·고액항암 = __무시__ 센티넬로 완전 드롭
  · 통합전이암 = 대표1개 (PPT·설명지 반영)
  · 세부가입 파서·오버랩 근절·CI판정·표준금액 제외 전부 포함

■ 오늘 추가(report_pptx.py):
  · 설명지 PDF = 진단서 PPT 동일 rep → 100% 일치
  · CI/비CI 자동 분기 (P4): CI 80%형 / 비CI Plan B
  · 검증: 주정은 CI 80%·사망3,000 정확

■ 배포:
  1. GitHub에 main.py + report_weasy.py + report_pptx.py(신규) + coverage_benchmark.py
  2. Railway Deploy 클릭
  3. /health로 버전 확인
  보장분석실: https://web-production-4a155.up.railway.app
  /health:    https://web-production-4a155.up.railway.app/health
