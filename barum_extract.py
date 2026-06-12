# -*- coding: utf-8 -*-
"""
barum_extract.py v3 — 2026.06.12
핵심 수정:
  1) 별첨 파싱 조건 완화: '(정상계약 리스트)' 완전일치 → '정상계약' 포함 여부로 변경
  2) 납입년수 파싱 다중패턴: '납입주기' 라벨 외 '월납/연납/횟수' 직접 탐색
  3) 갱신판정 강화: 담보명 '갱신형' 포함 → 갱신 우선
  4) NH농협생명·NH농협손보 필터 추가
  5) 다중페이지 별첨 병합 안전화
"""
import fitz, re, json, sys, datetime

AMOUNT  = re.compile(r'^[\d,]+$')
FOOTER  = re.compile(r'(지점|LP|☎|\d+/\d+페이지|상기 내용|본 리포트|참고용)')
GAENG   = re.compile(r'\(갱신')
YEARS   = re.compile(r'(\d{4})\.\d{2}\.\d{2}')
CNT_PAT = re.compile(r'^(\d+)/(\d+)$')

# NH농협 제외 (법칙 확정 2026.06.11)
NH_SKIP = re.compile(r'(NH농협생명|NH농협손보|농협생명|농협손해보험)')

def _clean(nm):
    """담보명 정제: 불필요한 접두/특수문자 제거"""
    nm = re.sub(r'^[\s\-·•▶▷]+', '', nm)
    nm = re.sub(r'\s+', ' ', nm).strip()
    return nm

def parse_page(text):
    """별첨 1페이지 파싱. '정상계약' 포함 여부로 조건 완화."""
    L = [l.strip() for l in text.split('\n') if l.strip()]
    if not L:
        return None
    # 헤더 탐지: '정상계약 리스트' 혹은 '보험서비스(상품)별 보장현황' 포함 페이지
    header_idx = None
    for i, l in enumerate(L):
        if '정상계약' in l or ('보장현황' in l and '별첨' in l):
            header_idx = i
            break
    if header_idx is None:
        return None

    # 회사명·상품명: 헤더 다음 2줄
    회사 = L[header_idx+1] if header_idx+1 < len(L) else ''
    상품 = L[header_idx+2] if header_idx+2 < len(L) else ''

    # NH농협 제외
    if NH_SKIP.search(회사) or NH_SKIP.search(상품):
        return None

    납입 = 만기 = ''
    for i, l in enumerate(L):
        if '납입주기' in l and i+1 < len(L):
            납입 = L[i+1]
        if '~' in l and YEARS.search(l):
            만기 = l
        # '납입기간' 직접 표기 패턴
        m = re.search(r'납입기간[:\s]*(.+)', l)
        if m and not 납입:
            납입 = m.group(1).strip()

    # 보장기간 시작점 탐색
    start = 0
    for i, l in enumerate(L):
        if l in ('보장기간', '담보명', '보장내용'):
            start = i + 2
            break

    covers = []
    buf = []
    for l in L[start:]:
        if FOOTER.search(l):
            break
        if AMOUNT.match(l) and buf:
            nm = _clean(' '.join(buf))
            if nm:
                covers.append([nm, int(l.replace(',', '')), bool(GAENG.search(nm))])
            buf = []
        elif not AMOUNT.match(l) and l not in ('원', '-'):
            buf.append(l)

    if not covers:
        return None

    return {'회사': 회사, '상품': 상품, '납입': 납입, '만기': 만기, '담보': covers}


def contract_renewal(c):
    """
    갱신판정 5규칙 (법칙 v2 §12):
    ① 담보명/상품명에 '갱신형' → 갱신
    ② 운전자 → 비갱신  (※ 2026.06.11: 운전자 규칙 미결이나 보수적으로 유지)
    ③ 9999 → 비갱신(종신)
    ④ 총회차 > 240 → 비갱신
    ⑤ 납입년수 = 보장년수 → 갱신, 다르면 비갱신
    """
    상품, 납입, 만기 = c['상품'], c['납입'], c['만기']

    # ① 갱신형 명시
    if '갱신형' in 상품:
        return '갱신'
    # 담보 중 하나라도 갱신형이면 계약 갱신
    if any(g for _, _, g in c.get('담보', [])):
        return '갱신'

    # ② 운전자
    if re.search(r'운전자(보험|상해)', 상품):
        return '비갱신'

    # ③ 종신
    if '9999' in 만기 or '종신' in 상품:
        return '비갱신(종신)'

    # 납입횟수 파싱 (다중 패턴)
    납입년 = 0
    # 패턴1: "N/M" 횟수 표기
    m = re.search(r'/(\d+)', 납입)
    if m:
        납입년 = round(int(m.group(1)) / 12)
    # 패턴2: "Ny납" 표기
    if not 납입년:
        m = re.search(r'(\d+)\s*[년y]', 납입)
        if m:
            납입년 = int(m.group(1))
    # 패턴3: "Nm납" → 개월
    if not 납입년:
        m = re.search(r'(\d+)\s*[월m]', 납입)
        if m:
            납입년 = round(int(m.group(1)) / 12)

    # ④ 총회차 > 240
    if 납입년 * 12 > 240:
        return '비갱신'

    # ⑤ 납입년 vs 보장년
    yrs = YEARS.findall(만기)
    if 납입년 and len(yrs) >= 2:
        보장년 = int(yrs[-1]) - int(yrs[0])
        return '갱신' if 납입년 == 보장년 else '비갱신'

    return '비갱신'


def is_expired(c):
    yrs = YEARS.findall(c['만기'])
    if not yrs:
        return False
    last_yr = int(yrs[-1])
    return last_yr < datetime.date.today().year and last_yr != 9999


def extract(pdf):
    d = fitz.open(pdf)
    raw = []
    for pg in d:
        parsed = parse_page(pg.get_text())
        if parsed:
            raw.append(parsed)

    # 동일 상품+만기 연속페이지 병합
    merged = []
    for p in raw:
        if (merged
                and merged[-1]['상품'] == p['상품']
                and merged[-1]['만기'] == p['만기']):
            merged[-1]['담보'] += p['담보']
        else:
            merged.append(p)

    # 만료계약 제외
    merged = [c for c in merged if not is_expired(c)]

    for i, c in enumerate(merged, 1):
        c['계약'] = i
        c['갱신구분'] = contract_renewal(c)
        c['담보수'] = len(c['담보'])
        c['갱신담보'] = [nm for nm, _, g in c['담보'] if g]

    return merged


if __name__ == '__main__':
    cs = extract(sys.argv[1] if len(sys.argv) > 1 else 'src.pdf')
    for c in cs:
        gd = f" ·칸별갱신 {len(c['갱신담보'])}건" if c['갱신담보'] else ""
        print(f"계약{c['계약']:2} [{c['갱신구분']:6}] {c['회사']:13}{c['상품'][:26]:26} 담보{c['담보수']:3}{gd}")
    json.dump(cs, open('contracts.json', 'w'), ensure_ascii=False, indent=1)
    print(f"\n총 {len(cs)}계약(만료제외) 추출 → contracts.json")
