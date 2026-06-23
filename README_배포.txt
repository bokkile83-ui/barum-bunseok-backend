BARUM 배포 패키지 — v27b-exclude-20260624
=========================================
포함:
  main.py      : 패치본(자동차보험 제외강화 / 무보험차상해 오매핑삭제 / 실손입원 5,000 행단위고정)
  master.xlsx  : 마스터 엑셀 90행 (변경 없음, 참고용 동봉)

GitHub Upload 후 Railway 자동 재배포 → /health 가
  v27b-exclude-20260624
뜨면 완료.

★ 동봉 안 함 (이 환경에서 유효본 미보유 → GitHub 기존 파일 그대로 두면 됨):
  ppt_form.pptx, chiryo_form.pptx  (PPT 폼 2종)
  requirements.txt, Procfile 등 배포설정
  → main.py 만 교체하면 됨. 위 파일들은 건드리지 말 것.
