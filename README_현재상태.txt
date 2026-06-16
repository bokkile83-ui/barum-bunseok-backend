===== BARUM 보장분석 앱 — 현재 상태 / 다음 할 일 (2026.06.15) =====

[ 들어있는 파일 ]
- main.py ............... 최신 코드(FINAL, haiku + PPT등식2). 첫 줄에 FINAL 표식.
- requirements.txt / Procfile / nixpacks.toml ... 배포 설정
- MASTER_보장분석_엑셀_영구표본.xlsx ... 엑셀 마스터
- MASTER_보장분석지_PPT_빈폼.pptx / 치료비_정리_빈폼.pptx ... PPT 마스터
- 지침_통합정본_20260615.md ... 정본 지침
- 결과샘플_정기철.xlsx / .pptx ... 대역 매핑 결과(참고용, 엑셀 85%)

[ 출근 후 할 일 — 순서대로 ]
1. PC로 GitHub 접속: https://github.com/bokkile83-ui/barum-bunseok-backend
2. Add file -> Upload files -> main.py(첫 줄 FINAL 적힌 것) 올리고 Commit
   (복붙 말고 파일 통째 업로드. 같은 이름이라 덮어써짐)
3. Railway 자동 재배포 대기 (Deployments 초록불)
4. 앱(https://web-production-4a155.up.railway.app/)에 박조은.txt 업로드
5. 매핑이 12행 -> 40행대로 뛰면 = 자동화 성공
   12행이면 = haiku 실패 -> Railway View logs 캡처해서 가져오기

[ 핵심 사실 ]
- ANTHROPIC_API_KEY 등록 확인됨.
- 자동화 AI 코드는 Claude가 작성(llm_resolve/llm_extract). 사용자는 올리기만.
- 100%는 불가: 보험사별 PDF 구조 편차가 팩트.
  * 깨끗한 회사(NH/롯데/현대) = 자동 처리됨
  * 깨지는 회사(동양 종신 암 등) = [확인] 시트로 빠짐 -> 신인이 약관 보고 수기
  * 깨지는 별첨은 Adobe "엑셀로 내보내기"로 재변환하면 살 수도 있음
- 현실적 천장 = 약 90% (나머지는 신인이 [확인] 칸 채움)
