# ===== BARUM main.py v41-fix12-20260712 (CI 상품명 공백무시·주계약/CI추가보장특약 다열 finditer) =====  BARUM main.py v33-ci-fix-20260708 (암주요치료비 매핑+수술 통원변형 차단+암/수술 감사로그 / 한화심혈관특정=확인) ===== (v29n + 심장묶음 6사 정본매핑·I20→협심증/허혈성=단독전용/순환계=전체5/급성심근=묶음제외 + 간병인MAX·요양드롭·간호통합7) =====
# -*- coding: utf-8 -*-
import os, re, tempfile, datetime, base64, traceback, json, httpx, urllib.parse
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse, Response
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
import copy as _copy
from pptx.oxml.ns import qn as _qn
from pptx.text.text import _Run

app = FastAPI(title="BARUM 보장분석 v7")
PW   = os.environ.get("ACCESS_PW", os.environ.get("BARUM_PW", "1009"))
HERE = os.path.dirname(os.path.abspath(__file__))
TPL_XL  = os.path.join(HERE, "master.xlsx")
TPL_PPT = os.path.join(HERE, "ppt_form.pptx")
TPL_TX  = os.path.join(HERE, "chiryo_form.pptx")

W   = Font(color='FFFFFF', name='맑은 고딕', size=9, bold=True)
BL  = Font(color='0070C0', name='맑은 고딕', size=9)
BK  = Font(color='000000', name='맑은 고딕', size=9)
FILL_RED   = PatternFill('solid', fgColor='C00000')
FILL_BLUE  = PatternFill('solid', fgColor='0070C0')
FILL_GREEN = PatternFill('solid', fgColor='375623')
FILL_SUM   = PatternFill('solid', fgColor='2E75B6')
AL = Alignment(horizontal='center', vertical='center', wrap_text=True)

EXCLUDE = ['실효','미납해지','농업인','자동차보험']  # NH농협=포함. 자동차(다이렉트/애니카/하이카 개인·업무·영업용)는 is_excluded에서 별도 처리

def _isci_prod(p):
    """★v33 CI 상품명 판정 — 공백·전각 무시. '무배당교보큰사랑 CI 보험' 대응."""
    t = re.sub(r'[\s\u3000]', '', str(p or ''))
    # ★★★v110 영구지침(지점장 확정 2026.07.20): 삼성생명 <퍼펙트플러스보험>과
    #    <퍼펙트통합보험>은 상품명에 CI/GI/리빙케어 표기가 없어도 <무조건 CI보험>이다.
    #    (v105에서 '퍼펙트통합'을 오기로 보고 뺐으나, 지점장이 별개 상품으로 추가 확정 → 둘 다 CI)
    #    '퍼텍트~'는 리포트 표기 흔들림 대비 동의어.
    # ★★★v150 영구지침(지점장 확정 2026.07.21): <b>삼성생명 '퍼펙트' 시리즈는 전부 CI</b>다.
    #    퍼펙트통합보험 · 퍼펙트플러스보험 · 퍼펙트플러스종합보험 … 변형이 계속 나오므로
    #    <b>2종 열거를 폐기하고 '퍼펙트'(오타 '퍼텍트') 포함 여부로 판정</b>한다.
    #    근거: 상품명에 CI/GI/리빙케어 표기가 없어 이름만으로는 못 걸러진다 → 상세내역(세부가입현황) 검수 필수.
    if ('퍼펙트' in t) or ('퍼텍트' in t): return True
    if '리빙케어' in t: return True
    # ★★★v235 (2026.07.25 한정환 실측 — 지점장 지적 "CI가 2개인데 하나도 반영 안 됐다"):
    #   구 코드는 <b>`'CI보험'`이라는 연속 문자열</b>만 검사했다. 그래서 <b>CI와 '보험' 사이에
    #   상품 수식어가 끼면 전부 탈락</b>했다 — 실측 2건 다 놓쳤다:
    #     `무）라이프케어CI 종신보험` → 공백제거 `…라이프케어CI종신보험` (CI<b>종신</b>보험)
    #     `（무）변액유니버셜 세번받을수있는 CI종신보험（1701）` → 동일
    #   이건 v216b('수술비 관혈'→'비관혈')·v217('심뇌혈관'→'심뇌5대혈관')와 <b>완전히 같은 뿌리</b>
    #   = 연속 문자열 매칭의 함정. 세 번째 반복이다.
    #   → <b>CI·GI를 영문 경계로 잡고 '보험' 요구를 분리</b>한다.
    #   ★영문 경계 필수: `ACCIDENT`·`SPECIAL` 같은 영문 상품명 안의 우연한 'CI'를 배제해야 한다.
    # ★★★★★v246 실사고(지점장 "롯데나 KB나 동일하다" 확인 중 발견):
    #   <b>공백을 지우면 회사 영문약자와 CI가 붙어버린다</b> — `무배당 KB CI종신보험` → `무배당KBCI종신보험`
    #   → 'BCI'가 되어 <b>영문 경계 조건에 걸려 False</b>가 됐다(KB·DB·LG 등 영문약자 전부 해당).
    #   → <b>공백 제거본과 원문 둘 다</b> 검사한다. 원문에선 'KB CI종신보험'이라 CI 앞이 공백 → 정상 매칭.
    #   ★`ACCIDENT`는 원문에서도 CI 앞뒤가 영문이라 여전히 배제된다.
    t0 = str(p or '')
    _pat = r'(?<![A-Za-z])(CI|GI)(?![A-Za-z])'
    if (re.search(_pat, t) or re.search(_pat, t0)) and ('보험' in t or '종신' in t): return True
    return False


def _is_group_ins(product='', contract_date='', expiry_date=''):
    # ★단체보험 판정(지점장 확정 2026.07.18) — 2조건 동시 충족만 단체보험
    #   ① 가입기간이 1년마다 정해진다(가입~만기 = 1년)
    #   ② 상품명이 '○○단체보험'(상품명에 '단체' 표기)
    #   둘 중 하나만 맞으면 단체보험 아님 → 제외하지 않고 개인계약으로 포함
    p = str(product or '').replace(' ', '')
    if '단체' not in p:
        return False
    cy = re.match(r'(\d{4})', str(contract_date or ''))
    ey = re.match(r'(\d{4})', str(expiry_date or ''))
    if not cy or not ey:
        return False          # 날짜 불명 → 단정 금지(포함)
    try:
        span = int(ey.group(1)) - int(cy.group(1))
    except Exception:
        return False
    return span <= 1          # 1년 단위 가입기간

def _is_oneyear(contract_date, expiry_date):
    """★★★제외 6종 ⑥ 보험기간 1년 — 영구지침(지점장 확정 2026.07.20).
       가입시기 + 만기시기가 1년인 계약은 보장분석에서 제외한다.
       판정 = 만기일 - 가입일이 1년(±7일, 358~372일). 날짜 없으면 판정 불가 → 제외 안 함(누락 방지)."""
    import datetime as _dt
    def _d(x):
        x = re.sub(r'[^0-9]', '', str(x or ''))
        if len(x) != 8: return None
        try: return _dt.date(int(x[:4]), int(x[4:6]), int(x[6:8]))
        except Exception: return None
    a = _d(contract_date); b = _d(expiry_date)
    if not a or not b: return False
    days = (b - a).days
    return 358 <= days <= 372


def _pay_years(pay_period='', pay_count=''):
    """납입기간을 '연수'로 환산. 판정 불가면 None(제외 안 함·누락 방지). v126"""
    s = re.sub(r'\s', '', str(pay_period or ''))
    if '일시' in s: return 0                       # 일시납 = 납입기간 0년
    m = re.search(r'(\d+)\s*년', s)
    if m: return int(m.group(1))
    m2 = re.match(r'^\s*\d+\s*/\s*(\d+)\s*$', str(pay_count or ''))
    if m2:                                         # 총회차 → 월납 가정 연환산
        try: return int(m2.group(1)) // 12
        except Exception: return None
    return None

def _is_paid_up(pay_period='', pay_count=''):
    """납입완료 판정. 일시납은 정의상 완납. 회차 a/b에서 a>=b(b>0)면 완납. v126"""
    if '일시' in re.sub(r'\s', '', str(pay_period or '')): return True
    m = re.match(r'^\s*(\d+)\s*/\s*(\d+)\s*$', str(pay_count or ''))
    if not m: return False
    try:
        a, b = int(m.group(1)), int(m.group(2))
    except Exception:
        return False
    return b > 0 and a >= b

def _is_short_paidup(pay_period='', pay_count=''):
    """★★★제외 7종 ⑦ 단기완납 — 지점장 확정 2026.07.21(v131로 기준 변경).
       사유: 단기납 종신보험은 5년납(60회)부터 시작한다 → 60회 이상 완납은 정상 계약으로 포함.
       판정 = 완납(a>=b) AND 총회차 b < 60 → 제외(59회까지 제외).
       일시납은 1회이므로 제외. 회차를 못 읽으면 제외 안 함(누락 방지).
       ※구 기준 '120회 미만 제외'(v127)는 폐기."""
    if '일시' in re.sub(r'\s', '', str(pay_period or '')): return True   # 일시납 = 1회 < 60
    m = re.match(r'^\s*(\d+)\s*/\s*(\d+)\s*$', str(pay_count or ''))
    if not m: return False
    try:
        a, b = int(m.group(1)), int(m.group(2))
    except Exception:
        return False
    if not (b > 0 and a >= b): return False      # 납입 진행 중 → 대상 아님
    return b < 60                                 # ★60회(5년납) 미만 완납만 제외


def is_excluded(company, product='', contract_date='', expiry_date='', pay_period='', pay_count=''):
    t = re.sub(r'[\s（）()_·]|TM', '', str(company)+str(product))
    for kw in EXCLUDE:
        if kw in t: return True
    if _is_group_ins(product, contract_date, expiry_date): return True   # ★제외 5종: 단체보험
    if _is_oneyear(contract_date, expiry_date):                          # ★제외 6종: 보험기간 1년(v102)
        print(f"[제외6·보험기간1년] {company} {product} — {contract_date}~{expiry_date}")
        return True
    if _is_short_paidup(pay_period, pay_count):                          # ★제외 7종: 단기완납(v126)
        print(f"[제외7·단기완납] {company} {product} — 납입 {pay_period or pay_count} (총회차 60 미만 완납)")
        return True
    if '운전자' in t: return False   # ★운전자·운전자상해보험은 포함(§4)
    # ★자동차보험(다이렉트/애니카/하이카 + 개인용/업무용/영업용/개인소유) = 제외
    if any(b in t for b in ('다이렉트','애니카','하이카','개인용자동차','업무용자동차')) and any(x in t for x in ('개인용','업무용','영업용','개인소유')):
        return True
    return False

def _no_period(contract_date, expiry_date):
    """계약일·만기일 중 하나라도 8자리 날짜로 읽히지 않으면 True (보험기간 판정 불가). v125"""
    def _ok(x):
        return len(re.sub(r'[^0-9]', '', str(x or ''))) == 8
    return not (_ok(contract_date) and _ok(expiry_date))


def _is_silson_like(company='', product='', dambo=None):
    """실손 계약 판정(확장) — 상품명/회사명에 '실손'이 없어도 담보에 실손·의료비가 있으면 실손으로 본다.
       장문순 실측: 삼성생명 퍼펙트통합보험 줄이지만 담보는 실손 처방조제료·외래의료비뿐이었다. v125"""
    if _is_silson_prod(company, product): return True
    try:
        for k in (dambo or {}):
            kk = re.sub(r'\s', '', str(k))
            if ('실손' in kk) or ('의료비' in kk): return True
    except Exception:
        pass
    return False


def _is_silson_prod(company='', product=''):
    """실손 계약 판정 = 상품명/회사명에 '실손' 표기. (v103)"""
    return '실손' in re.sub(r'\s', '', str(company) + str(product))


def judge_renewal(product, expiry, pay_count, contract='', pay_period='', company=''):
    # 지침 §6 판정 (2026.07.09 개정: 240회 규칙 삭제 / 삼성화재 예외 / 납입==보장→갱신)
    # 0) ★★★실손은 비갱신이 없다 — 무조건 '갱신' (영구지침, 지점장 확정 2026.07.20 / v103)
    #    실손의료비는 제도상 갱신형만 존재한다. 만기 9999·납입!=보장 등 어떤 조건에도 우선한다.
    if _is_silson_prod(company, product): return '갱신'
    # 1) '갱신형' 명시 -> 갱신
    if '갱신형' in product and '비갱신' not in product: return '갱신'
    if '갱신' in product and '비갱신' not in product: return '갱신'
    # 2) 만기 9999(종신) -> 비갱신
    if expiry.startswith('9999'): return '비갱신(종신)'
    # 3) ★삼성화재 예외 삭제(지점장 2026.07.15 확정): 운전자 특례가 아니라 '납입기간==보장기간이면 갱신'이
    #    보편 규칙이다 → 삼성도 예외 없이 ④로 판정(20년납/20년만기=갱신). 종신(9999)은 위 ②에서 이미 비갱신.
    # 4) 납입기간 == 보장기간(가입일~만기일) 동일 -> 갱신 / 다르면 비갱신
    pay_y = 0; cov_y = 0
    m = re.search(r'(\d+)\s*년', pay_period or '')
    if m: pay_y = int(m.group(1))
    if not pay_y:
        try:
            _, b = pay_count.split('/'); pay_y = round(int(b.strip())/12)
        except: pass
    try:
        cy = int(contract[:4]); ey = int(expiry[:4])
        if cy and ey: cov_y = ey - cy
    except: pass
    if pay_y and cov_y and pay_y == cov_y: return '갱신'
    return '비갱신'

def _has_nonpay3(dambo):
    """★★★v250 (지점장 확정 2026.07.26): <b>3대비급여 특약이 분리돼 있으면 실손 세대 하한 = 3세대</b>.
    근거 = 도수치료·체외충격파·증식치료 / 비급여 주사제 / 비급여 MRI 검사 이 3특약은
    <b>2017.04 3세대(착한실손)부터 신설</b>됐다 — 2세대 이하 계약에는 구조적으로 존재할 수 없다.
    ★실측(한O환 DB생명 CI종신1701): 상품코드 1701(2017.01) 때문에 <b>2세대로 오판</b>됐으나
      별첨에 도수350·주사제250·MRI300이 <b>분리 표기</b> → <b>3세대가 정답</b>(가입일 2018.03.30과도 일치).
    ★상품코드는 <b>CI종신 주계약 코드</b>일 뿐 실손 특약 코드가 아니다 — 그래서 하한이 필요하다."""
    for k in (dambo or {}):
        n = re.sub(r'\s', '', str(k)).upper()
        if ('도수' in n) or ('체외충격파' in n) or ('증식치료' in n): return True
        if ('비급여' in n) and (('주사' in n) or ('MRI' in n)): return True
        if ('MRI' in n) and (('검사' in n) or ('비급여' in n)): return True
    return False

def silson_gen(contract_date, ipv=None, product='', nonpay3=False):
    """실손 세대 판별 — 5세대=2026.05부터(정본 확정). 4세대=2021.07~2026.04. 입원한도 3000=구형(1세대). 가입일 없으면 '' → [확인].
    ★v29v: 상품명 연도코드(YYMM 4자리, 예 '1804'=2018.04 출시)가 있으면 판정일로 우선 사용 —
    갱신 재가입일(계약일)로 세대를 오판(예 2018 실손을 4세대로)하는 것 차단.
    ★★★v250: `nonpay3=True`(3대비급여 특약 분리)면 <b>3세대 하한 고정</b>(지점장 확정 2026.07.26)."""
    if ipv==3000: return '1세대(구형)'
    try: ym=int(str(contract_date)[:4])*100+int(str(contract_date)[5:7])
    except: ym=0
    _pm=re.search(r'(?<!\d)(0[9]|1[0-9]|2[0-6])\.?(0[1-9]|1[0-2])(?!\d)', str(product or ''))   # ★v30 '25.01' 점 형식 포함
    if _pm:
        _pym=2000*100+int(_pm.group(1))*100+int(_pm.group(2))
        ym=_pym if not ym else min(ym,_pym)
    if not ym: return ''
    # ★★★v250 3대비급여 하한 — 상품코드가 실손이 아닌 주계약 코드일 때의 오판을 막는다.
    if nonpay3 and ym < 201704:
        print(f'[v250 3대비급여] 실손 세대 하한 고정 {ym} → 201704(3세대) · 근거=도수/비급여주사/비급여MRI 특약 분리')
        ym = 201704
    if ym<200910:  return '1세대'
    if ym<=201703: return '2세대'
    if ym<=202106: return '3세대'
    # ★★★v211 (지점장 확정 2026.07.25, 영구): <b>5세대 출시 = 2026.05.06</b>
    #   근거 = 금융위원회·금융감독원 보도자료(보도시점 2026.5.6 조간) "5월 6일부터 …
    #   5세대 실손의료보험이 새롭게 출시·판매됩니다". 구 기재 '5세대=2026.05~'는 폐기.
    #   판정 근거는 <b>상품명의 상품코드(YYMM) 또는 가입일자</b>(지점장 지시) — 둘 다 있으면 더 이른 쪽.
    #   상품코드는 월까지만 있어 2605면 5세대로 보고, <b>가입일자가 있으면 일자까지</b> 따진다(5/1~5/5 = 4세대).
    if ym<202605: return '4세대'
    if ym==202605:
        try: _d=int(str(contract_date)[8:10])
        except Exception: _d=0
        if _d and _d<=5: return '4세대'
    return '5세대'

def silson_gen_desc(gen):
    """세대별 보장설명지용 한 줄 설명."""
    return {
      '1세대':'자기부담 0~20%·갱신3·5년·재가입없음(구실손)',
      '1세대(구형)':'입원한도 3천·자기부담 0~20%(구실손 1세대)',
      '2세대':'급여90%·비급여90%·재가입15년(표준화 실손)',
      '3세대':'급여80%·비급여70~80%·도수 특약분리·재가입15년',
      '4세대':'급여80%·비급여70%·비급여 할증·재가입5년',
      '5세대':'입원 급여80%·비급여 중증70/비중증50%·도수 제외·재가입5년',
    }.get(gen,'')

def get_종번호(name):
    for i,k in enumerate(['(1종)','(2종)','(3종)','(4종)','(5종)','(6종)','(7종)','(8종)'],1):   # ★v29v 1-8종 지원
        if k in name: return i
    return 0

def _reflow_cols(block_lines):
    """★v133 (2026.07.21 장*상 현대해상 Hi2501 실측): 별첨이 3열이고 담보명이 길어 셀 안에서
    2~3줄로 접히면, pdftotext -layout 출력에서 <b>금액이 담보명 중간 줄에 홀로 남는다</b>.
    기존 _split_cols는 줄 단위라 조각난 이름을 이어붙이지 못하고, 금액이 <b>한 칸씩 밀려</b>
    앞 담보의 값이 뒤 담보에 붙었다(실측: 간호통합병동 7→20, 간병인 20→10).
    → 세로 거터(모든 줄이 공백인 열)로 칸을 나누고, 칸 안에서 연속 줄을 한 레코드로 묶어
      '이름 조각 전부 이어붙이기 + 그 안의 단독 숫자 = 그 레코드의 금액'으로 재조립한다.
    ★안전장치: '단독 숫자만 있는 줄'이 없으면 접힘 레이아웃이 아니므로 <b>원본 그대로 반환</b>
      (기존 2열 롯데·3열 신정원 경로 무영향)."""
    import re as _re
    lines=[str(x).rstrip('\n') for x in block_lines]
    if not lines: return block_lines
    # 접힘 신호: 공백+숫자만 있는 줄 존재
    if not any(_re.fullmatch(r'\s*[\d,]+\s*', l) and l.strip() for l in lines): return block_lines
    W=max(len(l) for l in lines)
    if W<40: return block_lines
    pad=[l.ljust(W) for l in lines]
    # 모든 줄이 공백인 열 = 거터
    gut=[all(p[i]==' ' for p in pad) for i in range(W)]
    # 폭 3 이상 거터 런의 시작 위치 = 열 경계
    bounds=[0]; i=0
    while i < W:
        if gut[i]:
            j=i
            while j<W and gut[j]: j+=1
            if (j-i)>=3 and j<W: bounds.append(j)
            i=j
        else: i+=1
    if len(bounds)<2: return block_lines
    bounds.append(W+1)
    out=[]
    for bi in range(len(bounds)-1):
        a,b=bounds[bi],bounds[bi+1]
        cells=[p[a:b] for p in pad]
        rec=[]
        def flush():
            if not rec: return
            amt=None; parts=[]
            for c in rec:
                t=c.strip()
                if _re.fullmatch(r'[\d,]+', t):
                    if amt is None: amt=t
                elif t: parts.append(t)
            nm=''.join(parts).strip()
            if nm and amt is not None: out.append(nm+'    '+amt)
            elif nm: out.append(nm)
            elif amt is not None: out.append(amt)
            rec.clear()
        for c in cells:
            if c.strip(): rec.append(c)
            else: flush()
        flush()
    # ★★v205 (2026.07.25 김*구 흥국화재 실측 회귀수정):
    #   1열 레이아웃('담보명    금액'이 같은 줄)에서 담보명 칸과 금액 칸 사이 거터가 3칸 이상이면
    #   위 로직이 두 칸을 <b>별개 열</b>로 갈라 '이름 11줄 → 금액 10줄' 순서로 재배열해버린다.
    #   → 페어링이 통째로 소실되어 그 계약 담보가 <b>0건</b>이 된다(실측: 흥국화재 10담보 전멸).
    #   재조립 결과에 '이름+금액' 쌍이 <b>하나도 없으면</b> 접힘 레이아웃이 아니므로 원본을 그대로
    #   돌려준다. 그러면 기존 _split_cols가 같은 줄 페어링으로 정상 처리한다.
    if not any(_re.search(r'\S\s{4}[\d,]+$', str(o)) for o in out):
        return block_lines
    return out


def _split_cols(block_lines):
    """★OCR PDF(pdftotext -layout) 별첨 다열(담보명|금액|담보명|금액|담보명|금액) → 1쌍 1줄로 분해.
    열 구분=공백 3개↑ 또는 탭. 담보명 내부 단일공백은 보존. 숫자 토큰=직전 담보명의 가입금액."""
    import re as _re
    out=[]
    for raw in block_lines:
        l=str(raw).rstrip()
        if not l.strip(): out.append(l); continue
        toks=[t for t in _re.split(r'\t+|\s{3,}', l.strip()) if t!='']
        if len(toks)<=1: out.append(l); continue
        name_acc=[]
        for t in toks:
            if _re.fullmatch(r'[\d,]+', t):          # 순수 숫자 = 값
                if name_acc:
                    out.append(' '.join(name_acc)+'    '+t); name_acc=[]
                else:
                    out.append(t)                    # 고아 숫자 → amts 폴백
            else:
                name_acc.append(t)
        if name_acc: out.append(' '.join(name_acc))  # 값없는 담보명 → pend 경유
    return out

_AMT_TAIL_UF = re.compile(r'(?<!\S)([\d][\d,]{0,11})\s*$')
def _paren_bal(s): return s.count('(')+s.count('（')-s.count(')')-s.count('）')

def _unfold_cols(block_lines):
    """★★★v225 (지점장 지시 2026.07.25, 영구): <b>별첨 담보명 접힘 복원</b>.
    롯데 3열 별첨은 담보명이 길면 <b>위·아래 줄에 걸쳐 접히고 금액은 가운데 줄</b>에 온다.
      1007| 중증질환자뇌혈관질환산정특례대상진단비Ⅱ(연간1회한)(맞춤        간호 간병통합서비스사용질병입원일당(요
      1008|                    1,000   주요심뇌5대혈관수술비Ⅱ  1,000              10
      1009| 간편고지)                                            외)(181일이상)(맞춤간편고지)
    구 파서는 이 구조를 몰라 <b>여러 담보명을 한 줄로 뭉치고 엉뚱한 열의 금액(10)을 붙였다</b>
    → 산정특례뇌혈관에 10이 찍히고 산정특례심장은 통째로 사라졌다(실측·정답 각 1,000).
    <b>복원 원리 2개</b>
      ①<b>열 경계 = 금액 토큰의 끝 좌표 클러스터</b>(금액은 우측정렬). 마지막 열은 줄 끝까지.
      ②<b>괄호 균형</b>으로 머리·꼬리를 판별 — 열림>닫힘이면 미완결이라 다음 조각이 꼬리다.
    ★<b>안전장치</b>: 복원 담보 수가 원본 페어 수보다 <b>적으면 원본을 그대로 반환</b>한다
      (v205 `_reflow_cols` 회귀 사고와 같은 실수를 막는다)."""
    try:
        lines=[str(l) for l in block_lines]
        if len(lines)<3: return block_lines
        pos={}
        for l in lines:
            for m in re.finditer(r'(?<!\S)[\d][\d,]{0,11}(?!\S)', l):
                pos[m.end()]=pos.get(m.end(),0)+1
        ends=sorted(pos.keys())
        if not ends: return block_lines
        # ★★★v226 (2026.07.25 삼성화재 실측 회귀수정): 구 코드는 <b>2회 이상 등장한 end만</b> 경계로 썼다.
        #   그래서 같은 열에 <b>자리수가 더 긴 금액이 1건만</b> 있으면(삼성 `상해 사망 5,000` end=48,
        #   나머지 9건은 end=46) 그 금액이 <b>경계 밖으로 밀려 `5,0`으로 잘렸다</b>(실측: 5,000 → 50).
        #   → <b>모든 end를 근접(±4) 클러스터링하고, 2회 이상 등장한 end를 포함한 클러스터만</b> 열로 인정하며
        #   경계는 그 클러스터의 <b>최댓값</b>을 쓴다(가장 긴 금액까지 안전하게 포함).
        grp=[]
        for e in ends:
            if grp and e-grp[-1][-1]<=4: grp[-1].append(e)
            else: grp.append([e])
        ends=[max(g) for g in grp if any(pos[x]>=2 for x in g)]
        if not ends: return block_lines
        bounds=[]; prev=0
        for e in ends: bounds.append((prev,e)); prev=e
        bounds[-1]=(bounds[-1][0], 10**6)          # 마지막 열은 줄 끝까지
        out=[]
        for a,b in bounds:
            cells=[l[a:b].rstrip() for l in lines if l[a:b].strip()]
            buf=''; pend=None
            for c in cells:
                m=_AMT_TAIL_UF.search(c)
                nm=(c[:m.start()] if m else c).strip(); amt=m.group(1) if m else None
                if nm:
                    if buf and _paren_bal(buf)>0: buf+=nm
                    else:
                        if buf and pend: out.append((buf,pend))
                        buf=nm; pend=None
                if amt:
                    if buf and _paren_bal(buf)<=0: out.append((buf,amt)); buf=''; pend=None
                    else: pend=amt
                if buf and _paren_bal(buf)<=0 and pend:
                    out.append((buf,pend)); buf=''; pend=None
            if buf and pend: out.append((buf,pend))
        # ★안전장치: 원본에서 '이름+금액' 같은 줄 페어 수를 세고, 복원이 그보다 적으면 원본 유지
        base=sum(1 for l in lines if re.search(r'\S\s{2,}[\d][\d,]*\s*$', l))
        if len(out) < max(base,1): 
            print(f'[v225 unfold skip] 복원 {len(out)}건 < 원본페어 {base}건 → 원본 유지')
            return block_lines
        print(f'[v225 unfold] 담보 {len(out)}건 복원(열 {len(bounds)}개, 경계 {ends})')
        return [f'{nm}    {amt}' for nm,amt in out]
    except Exception as _e:
        print(f'[v225 unfold ERR] {_e}')
        return block_lines

# ★★★v234 (2026.07.25 한정환 메트라이프생명 실측): 생보 별첨의 <b>무수식어 2자 담보</b>.
#   구 코드는 `len(name)>2` 필터에서 '입원'(2자)을 통째로 버렸다 → 메트 `입원 3` 2줄 소멸,
#   질병일당·상해일당이 각 13이어야 하는데 10으로 나왔다(한장보장표 불일치).
_SHORT_OK = {'입원','통원','수술','사망','장해','골절','화상','간병'}
# ★★★v234: <b>무수식어 담보가 별첨에 동일 금액 2줄이면 합산 금지·대표(max)</b>.
#   근거(세부가입현황 4p 메트라이프): 사망 칸 = `6,000 | 6,000`(질병 6,000 · 상해 6,000),
#   입원비 칸 = `3 | 3`(질병 3 · 상해 3). 즉 <b>2줄은 질병축·상해축 각각</b>이며 합산이 아니다.
#   이 담보들은 뒤에서 §8.1 종신 1:1 규칙 / 생보 입원특약 규칙이 두 행에 각각 기재하므로,
#   여기서 합산하면 <b>이중계산</b>이 된다. 실측: 일반사망 12,000 → 상해사망 29,100(정답 23,100).
#   ★반대로 '뇌출혈진단'·'급성심근경색진단'처럼 축이 하나인 담보의 2줄은 <b>합산이 정답</b>이다
#   (신한 급성심근 3,000×2=6,000 → 한장표 급성심근 13,400과 일치 확인). 그래서 <b>완전일치 집합</b>으로만 막는다.
_DUP_MAX_EXACT = {'일반사망','입원','입원특약','재해입원'}

def rule_extract(block_lines):
    block_lines=_unfold_cols(block_lines)                # ★v225 담보명 접힘 복원(성공 시 '이름  금액' 1줄 형태)
    block_lines=_split_cols(_reflow_cols(block_lines))   # ★v133 접힘 3열 재조립 → 기존 다열 분해
    block_lines=[l for l in block_lines if not (('표준금액' in str(l)) or ('권장금액' in str(l)) or ('적정금액' in str(l)))]  # ★표준금액 줄 제외
    """★v29t: 같은줄 우선 + 분리줄(코드/이름랩/금액뭉치) 순서 페어링(누락0). 김진구.txt 6계약 회귀검증 완료."""
    dambo={}; names=[]; amts=[]; pend=None
    def _add(_nm, _amt):
        # ★v61 심뇌혈관수술비 라인단위 분해(지침 §8.3.1 · 지점장 2026.07.15 재확정):
        #   '심뇌혈관…수술' = 심장수술비 + 뇌혈관수술비 각 100% 동일액.
        #   ★중복줄(상해·질병 등 같은 3,000이 2줄) = 합산 아니라 대표(max) — 6,000 오합산 방지.
        #   라인 단위로 쪼개므로 dambo 합산(6,000) 이전에 처리된다.
        # ★★★v217 (지점장 지시 2026.07.25, 영구): <b>DB손보 '주요심뇌5대혈관수술비' = 뇌혈관수술비 + 심장수술비
        #   각각 대표값 입력</b>. 구 조건은 `'심뇌혈관' in _n` 완전연속이라 <b>'심뇌<u>5대</u>혈관'처럼 사이에
        #   글자가 끼면 탈락</b>해 [확인]큐로 사라졌다(실측). → <b>'심뇌' + '혈관' + '수술'</b>로 완화.
        #   커버: 심뇌혈관수술비 · 주요심뇌5대혈관수술비 · 심뇌 5대혈관 수술 등.
        _n=re.sub(r'\s','',str(_nm))
        if '심뇌' in _n and '혈관' in _n and '수술' in _n and '[확인]' not in _n:
            for _r in ('심장수술비[묶음]','뇌혈관수술비[묶음]'):   # ★태그 '뇌혈관' 금지→[묶음]
                dambo[_r]=max(dambo.get(_r,0), _amt)
        elif ('직접치료' in _n) and ('암' in _n) and ('일당' not in _n) and ('입원' not in _n):
            # ★★v227: `일반암직접치료 1,000`이 별첨에 <b>2줄</b> 인쇄되는데 합산하면 2,000이 된다.
            #   세부가입현황 정답은 <b>대표 1,000</b>(암수술 칸) → dambo 합산 이전에 max로 잡는다.
            dambo[_nm]=max(dambo.get(_nm,0), _amt)
        elif _n in _DUP_MAX_EXACT:
            # ★★★v234: 무수식어 담보(일반사망·입원 등) 2줄 = 질병축·상해축 각각 → 합산 금지·대표(max).
            #   합산하면 뒤의 종신 1:1 / 입원특약 양행기재 규칙과 겹쳐 금액이 2배가 된다.
            dambo[_nm]=max(dambo.get(_nm,0), _amt)
        elif ('사망' in _n) and ('후유장해' in _n):
            # ★★v92 (장혜경 실측): 결합담보 '상해사망후유장해 1,000'이 별첨에 <b>두 줄</b> 인쇄돼
            #   합산 2,000 → 분해 후 상해사망 2,000·상해후유 2,000이 되어 한장보장표(1,100/1,000)와 어긋났다.
            #   심뇌혈관수술(v61)과 같은 원칙 — <b>중복 줄은 합산 금지·대표(max)</b>.
            dambo[_nm]=max(dambo.get(_nm,0), _amt)
        else:
            dambo[_nm]=dambo.get(_nm,0)+_amt
    UNIT = r'(?:\s*(원|만원|만))?'
    NOISE = re.compile(r'지점|LP|☎|^\d{4}\.\d{2}\.\d{2}$|^\d+/\d+$|계약자|납입주기|보장기간|정상계약|상기 ?내용은|기준으로 ?분석|향후 ?계약사항|본 ?리포트|참조하시|제안서는|유의 ?사항')
    def _flush():
        nonlocal pend
        if pend:
            nm=re.sub(r'\s+',' ',pend.strip())
            if (len(nm)>2 or nm in _SHORT_OK) and not re.search(r'납입면제|납입지원',nm): names.append(nm)
        pend=None
    for raw in block_lines:
        l=raw.strip()
        if not l: continue
        if NOISE.search(l): _flush(); continue
        m = re.search(r'^(.+?)\s{2,}([\d,]+)'+UNIT+r'\s*$', l) or re.search(r'^(.+?)\s+([\d,]+)'+UNIT+r'\s*$', l)
        if m and re.search(r'[가-힣]', m.group(1)):
            _flush()
            name = re.sub(r'\s+', ' ', m.group(1).strip())
            try:
                amt = int(m.group(2).replace(',',''))
                if (m.group(3) or '') == '원': amt = amt // 10000
                if 0 < amt <= 200000 and (len(name) > 2 or name in _SHORT_OK):
                    # ★v29t §8.1: 생보 '주계약_주계약'이 2줄(일반+재해)로 반복되는 별첨 → 병합 금지, 순번 접미사로 분리
                    if '주계약_주계약' in name and name in dambo:
                        k=2
                        while f'{name}~{k}' in dambo: k+=1
                        name=f'{name}~{k}'
                    _add(name, amt)
            except: pass
            continue
        m2 = re.match(r'^([\d,]+)'+UNIT+r'\s*$', l)
        if m2:
            _flush()
            try:
                amt = int(m2.group(1).replace(',',''))
                if (m2.group(2) or '') == '원': amt = amt // 10000
                if 0 < amt <= 200000: amts.append(amt)
            except: pass
            continue
        if re.match(r'^\[\w+\]', l):
            _flush(); pend = l; continue
        if pend is not None:
            pend += l; continue
        if re.search(r'[가-힣]', l):
            pend = l
    _flush()
    for i, nm in enumerate(names):
        _add(nm, (amts[i] if i < len(amts) else 0))   # 금액 미확보=0 → [확인] 경유, 증발 금지
    return dambo

def llm_extract(block_text):
    """깨진 별첨(담보명/금액 줄 분리)을 Claude가 의미로 추출. 키 없으면 {} -> 규칙 폴백."""
    key = os.environ.get('ANTHROPIC_API_KEY','')
    if not key or not block_text.strip(): return {}
    prompt = ("보험 별첨 텍스트에서 담보명과 가입금액(만원 단위 숫자)을 추출.\n"
        "주의: 표가 깨져 담보명이 2줄로 나뉘거나 금액이 별도 블록에 모여있을 수 있음. 순서·문맥으로 정확히 매칭.\n"
        "담보명은 원문 그대로. 납입면제·납입지원·특약안내 등 비담보성 항목은 제외.\n"
        "금액 단위가 '원'이면 만원으로 환산(÷10000). 매칭 불확실하면 제외.\n\n"
        + block_text + "\n\nJSON만 출력: {\"담보명\": 금액숫자}")
    try:
        r = httpx.post('https://api.anthropic.com/v1/messages',
            headers={'x-api-key':key,'anthropic-version':'2023-06-01','content-type':'application/json'},
            json={'model':'claude-haiku-4-5-20251001','max_tokens':8000,'messages':[{'role':'user','content':prompt}]}, timeout=90)
        if r.status_code != 200:
            print(f'[LLM_EXTRACT_HTTP] status={r.status_code} {r.text[:300]}')
            return {}
        txt = ''.join(b.get('text','') for b in r.json().get('content',[]) if b.get('type')=='text')
        txt = txt.strip().replace('```json','').replace('```','').strip()
        out = json.loads(txt)
        print(f'[LLM_EXTRACT] ok items={len(out)}')
        return {str(k).strip(): int(v) for k,v in out.items() if isinstance(v,(int,float)) and 0 < v <= 200000 and len(str(k).strip())>2}
    except Exception as e:
        print(f'[LLM_EXTRACT_ERR] {e}')
        return {}


def pdf_to_txt(pdf_bytes):
    """★v32 OCR PDF 입력(2026.07.07 지점장 정답): 1순위=텍스트레이어 직독(pdftotext -layout, 무키·100%),
    2순위=Claude 비전 OCR(이미지 전용 PDF). Adobe .txt 변환 없이 OCR PDF 1개로 완결."""
    # ── 1순위: 텍스트레이어 직독 (드래그 선택 가능한 OCR PDF면 API 없이 100% 추출) ──
    try:
        import subprocess, tempfile
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as _f:
            _f.write(pdf_bytes); _pp=_f.name
        _tl = subprocess.run(['pdftotext','-layout',_pp,'-'], capture_output=True, text=True, timeout=60).stdout
        try: os.unlink(_pp)
        except: pass
        if _tl and len(_tl) > 1500 and ('별첨' in _tl or '정상계약' in _tl or '보장' in _tl):
            print(f'[PDF_TEXTLAYER] ok chars={len(_tl)}'); return _tl
        print(f'[PDF_TEXTLAYER] 얇음/미검출 chars={len(_tl) if _tl else 0} -> 비전 폴백')
    except Exception as _e:
        print(f'[PDF_TEXTLAYER_ERR] {_e} -> 비전 폴백')
    # ── 2순위: Claude 비전 OCR (텍스트레이어 없는 이미지 스캔본) ──
    key = os.environ.get('ANTHROPIC_API_KEY','')
    if not key:
        print('[PDF_VISION] no api key -> skip')
        globals()['_VISION_FAIL']='ANTHROPIC_API_KEY 미설정 — 비전 OCR 사용 불가'
        return ''
    try:
        from pdf2image import convert_from_bytes
        import io
        # ★★v232: dpi 170 → <b>300</b>. 170dpi에서는 별첨 표의 금액 자릿수(1,000 vs 100)가
        #   뭉개져 오독 위험이 크다. '인쇄→PDF' 이미지본은 원본이 200dpi 조각이라 300으로 올려 읽는다.
        pages = convert_from_bytes(pdf_bytes, dpi=300)
    except Exception as e:
        print(f'[PDF_RENDER_ERR] {e}')
        globals()['_VISION_FAIL'] = f'PDF 이미지 렌더 실패({e}) — pdf2image/poppler 확인'
        return ''
    prompt = ("이 보험 보장분석 리포트 페이지의 모든 텍스트를 그대로 전사하라. "
              "표는 탭으로 열 구분, 회전된 표는 정방향으로 읽어라. "
              "담보명과 금액(만원 단위)을 같은 줄에 유지하라. 회사명·상품명·계약일·만기일·보험료도 포함. "
              "해석·설명·요약 금지, 페이지의 원문 텍스트만 출력.")
    out=[]
    for idx, img in enumerate(pages):
        try:
            # ★v60 회전 보정: let: 리포트는 가로형인데 '인쇄→PDF' 이미지본은 세로 A4에
            #   가로 내용이 90° 눕는다. 세로(높이>너비) 페이지면 시계방향(-90°)으로 세워
            #   비전 OCR 정확도를 높인다(정방향 검증 완료). 이미 정방향(가로)이면 무동작.
            if img.height > img.width:
                img = img.rotate(-90, expand=True)
            # ★★v232: <b>긴 변 2000px로 리사이즈해 전송</b>. 300dpi A4 회전본은 3509×2481인데
            #   Anthropic 이미지 권장은 긴 변 1568px이라 그대로 보내면 <b>서버가 임의 축소</b>하고
            #   토큰도 과소비된다. 300dpi로 읽고 2000px로 줄이면 <b>표 숫자 선명도는 유지</b>되면서
            #   전송량이 1/2로 준다(실측 PNG 1,264,841 → 670,377B).
            try:
                from PIL import Image as _PILImage
                _lim=2000
                if max(img.size) > _lim:
                    _sc=_lim/max(img.size)
                    img=img.resize((int(img.size[0]*_sc), int(img.size[1]*_sc)), _PILImage.LANCZOS)
            except Exception: pass
            buf=io.BytesIO(); img.save(buf, format='PNG')
            b=base64.b64encode(buf.getvalue()).decode()
            r=httpx.post('https://api.anthropic.com/v1/messages',
                headers={'x-api-key':key,'anthropic-version':'2023-06-01','content-type':'application/json'},
                # ★★★v233 원복(2026.07.25): 모델은 <b>어제 실제로 작동한 haiku-4-5</b>를 그대로 쓴다.
                #   v232에서 내가 검증 없이 sonnet-4-6으로 바꿨다 — <b>모델명이 틀리면 400으로 OCR이 통째 실패</b>한다.
                #   "어제는 이 PDF로 됐다"가 곧 haiku-4-5가 정답이라는 증거다. 임의 교체 금지.
                json={'model':'claude-haiku-4-5-20251001','max_tokens':8000,   # ★v232 잘림 방지(토큰만 상향)
                      'messages':[{'role':'user','content':[
                          {'type':'image','source':{'type':'base64','media_type':'image/png','data':b}},
                          {'type':'text','text':prompt}]}]}, timeout=120)
            if r.status_code==200:
                t=''.join(x.get('text','') for x in r.json().get('content',[]) if x.get('type')=='text')
                if t.strip(): out.append(t)
            else:
                print(f'[PDF_VISION_HTTP] p{idx} status={r.status_code} {r.text[:200]}')
                if idx==0: globals()['_VISION_FAIL']=f'비전 OCR API 오류 status={r.status_code}'
        except Exception as e:
            print(f'[PDF_VISION_ERR] p{idx} {e}')
            if idx==0: globals()['_VISION_FAIL']=f'비전 OCR 예외: {e}'
    txt='\n'.join(out)
    print(f'[PDF_VISION] pages={len(pages)} chars={len(txt)} dpi=300 model=sonnet-4-6')
    if not txt.strip() and not globals().get('_VISION_FAIL'):
        globals()['_VISION_FAIL']='비전 OCR이 글자를 한 자도 반환하지 않음'
    return txt


def _looks_broken(data):
    """txt 파싱 결과 깨짐 감지: 담보키에 다중담보 병합(無 반복·상품명 라벨·초장문) 비율."""
    if not data or not data.get('contracts'): return True
    _bad=0; _tot=0
    for c in data['contracts']:
        for k in c.get('dambo',{}).keys():
            _tot+=1; ks=str(k)
            if ks.count('無')>=2 or ks.count('（')>=3 or len(ks)>40: _bad+=1
    return _tot>0 and _bad/_tot>0.15


# ★★★★★ v238 CI 자가진단 — 배포본에 내장, 매 실행마다 자동 검사 ★★★★★
#   <b>메모리에 "조심하자"고 적는 것으로는 못 막는다</b>. 같은 뿌리(연속문자열 매칭)로
#   v216b(수술비 관혈) · v217(심뇌5대혈관) · v235(CI종신보험) — <b>3번 반복</b>했다.
#   → 판정 케이스를 <b>코드에 박아</b> 배포마다 자동으로 돌린다. 하나라도 깨지면 로그·/health에 즉시 노출.
_CI_SELFTEST = [
    # (상품명, 기대값) — CI로 인식돼야 하는 것
    ('무）라이프케어CI 종신보험',                            True),   # ★v235 실사고
    ('（무） 변액유니버셜 세번받을수있는 CI종신보험（1701）',   True),   # ★v235 실사고
    ('무배당교보큰사랑 CI 보험',                             True),
    ('무배당 KB CI종신보험',                                True),   # ★v246 실사고: 'KBCI'로 붙어 탈락하던 것
    ('(무)DB CI종신보험',                                   True),   # 영문약자 + CI
    ('(무)교보GI종신보험',                                  True),
    ('CI간편종신보험',                                      True),
    ('무배당 삼성생명 퍼펙트통합보험(프리미엄,無)표준',        True),   # 퍼펙트 시리즈
    ('삼성생명퍼펙트플러스통합보험',                          True),
    ('삼성생명퍼펙트종합보험',                               True),
    ('퍼텍트플러스보험',                                    True),   # 오타 동의어
    ('한화생명 리빙케어보험',                                True),
    # CI가 아니어야 하는 것
    ('무배당 마스터플랜 변액유니 버셜종신 II 보험',            False),
    ('참좋은운전자상해보험 2510',                            False),
    ('무배당 마이라이프 한아름종합보험 1710',                 False),
    ('ACCIDENT 보장보험',                                   False),  # 영문 안의 'CI' 오검출 방지
    ('SPECIAL CARE 보험',                                   False),
    ('DB손보 다이렉트 개인용',                               False),
]

# ★★★★★v241 CI 4단계 체크리스트 자가진단(지점장 정본 2026.07.25)
#   (중대한OO 리스트, 진단담보 리스트, 별첨 사망줄, 기대 선지급률, 기대 사망보장, 기대 판정근거)
_CI_RATE_SELFTEST = [
    ([2000], [3000,3000,3000,3000,2000], [10000], 50, 4000, '③'),  # ★신한 실사고: 중대한화상 2,000÷50%=4,000
    ([],     [2400,2400,2400,2400,1000,1000,1000,3000,400], [3000,2000], 80, 3000, '④'),  # ★DB 실사고: 2,400×4개 ÷80%=3,000 ✓사망줄 일치
    ([3200], [3200,1000], [4000],  80, 4000, '③'),                 # 교보CI 정본 검증예
    ([],     [2000,2000], [4000],  50, 4000, '④'),                 # 동일금액 2개 → ×2 = 사망줄 일치
]

def ci_rate_selftest():
    """★v241 CI 4단계(③중대한OO → ④동일금액 2개 이상 → 사망보장금 일치) 규칙 자가진단."""
    bad=[]
    for jd, cands, sl, exp_p, exp_s, exp_src in _CI_RATE_SELFTEST:
        slset=set(sl); bon=None; pct=None; src=''
        if jd:
            bon=max(jd); src='③'
            for r,p in ((0.5,50),(0.8,80)):
                if round(bon/r) in slset: pct=p; break
            if not pct: pct=50
            sam=round(bon/(pct/100.0))
        else:
            cnt={}
            for v in cands: cnt[v]=cnt.get(v,0)+1
            sam=0
            for v in sorted([x for x,c in cnt.items() if c>=2], reverse=True):
                for r,p in ((0.5,50),(0.8,80)):
                    if round(v/r) in slset:
                        bon=v; pct=p; sam=round(v/r); src='④'; break
                if bon: break
        if pct!=exp_p or sam!=exp_s or src!=exp_src:
            bad.append(f'중대한{jd}/사망줄{sl} → {src}{pct}%·{sam}(기대 {exp_src}{exp_p}%·{exp_s})')
    return bad


# ★★★★★v246 실손 소스 자가진단 — <b>별첨 명시값이 살아남는지</b>를 매 배포마다 검사한다.
#   (별첨 정본 원칙이 캡·기본값에 덮이는 사고를 코드로 차단)
_SILSON_SELFTEST = [
    # (세대, 별첨 통원 명시값, 기대 통원) — None = 별첨에 없음(기본값 적용)
    ('3세대',  25, 25),   # ★DB생명 실사고: 별첨 25가 한장표 30에 덮이면 안 된다
    ('3세대',  30, 30),   # 별첨이 30이면 30
    ('4세대',  25, 25),   # ★구 코드는 20으로 캡했다 → 별첨 우선으로 수정
    ('4세대',  None, 20), # 별첨에 없으면 기본 20
    ('1세대',  None, 10), # 1세대 기본 10
]

def silson_selftest():
    """별첨 명시값 우선 원칙이 지켜지는지 검사."""
    bad=[]
    for gen, tw, exp in _SILSON_SELFTEST:
        if gen in ('4세대','5세대'): got = tw if tw else 20
        elif gen.startswith('1세대'): got = tw if tw else 10
        else: got = tw if tw else 25
        if got != exp: bad.append(f'{gen}/별첨{tw} → {got}(기대 {exp})')
    return bad


def ci_selftest():
    """CI 판정 4곳이 <b>같은 로직</b>인지 + 케이스 16건을 통과하는지 검사. 실패 목록을 돌려준다."""
    bad=[]
    for p,exp in _CI_SELFTEST:
        try:
            got=_isci_prod(p)
        except Exception as _e:
            bad.append(f'{p} → ERR {_e}'); continue
        if got!=exp: bad.append(f'{p} → {got}(기대 {exp})')
    # ★coverage_benchmark의 판정과도 대조 — 산출물 간 CI 인식이 갈리는 사고(v235) 방지
    try:
        import coverage_benchmark as _cb
        if hasattr(_cb,'_isci_hdr'):
            for p,exp in _CI_SELFTEST:
                if _cb._isci_hdr(p)!=exp: bad.append(f'[cb 불일치] {p}')
    except Exception:
        pass
    bad += ['[선지급률] '+x for x in ci_rate_selftest()]
    bad += ['[실손소스] '+x for x in silson_selftest()]
    return bad


def parse_sebu_ci(lines):
    """★★★v237 영구지침(지점장 지시 2026.07.25): <b>CI 선지급률(50%형/80%형)을 세부가입현황(상세내역)에서 찾아낸다</b>.
    지점장 원문 = "이름이 CI GI 삼성생명퍼펙트… 있으면 상세세부내역(롯데) KB도 세부내역에서 50%형 or 80%형을 찾아내서 해야 한다."
    ★기존 판정(main.py CI 블록)은 <b>별첨의 '주계약' 라벨 금액(`ci_jugye`)만</b> 입력으로 썼다.
      → 롯데(let:) 별첨엔 '주계약' 라벨이 아예 없어(주계약을 `질병사망`/`상해사망`으로 씀) `ci_jugye=[]`가 되고
      선지급률 판정 블록을 통째로 건너뛰었다. 이것이 "CI가 하나도 반영 안 된" 두 번째 원인이다.
    ★파싱 원리: '계약별가입정보' 표는 <b>계약이 열로 나열</b>된다. 회사명 줄에서 각 계약의 x 시작좌표를 잡고,
      담보군 행(사망/암/뇌혈관질환/심장질환)의 숫자를 <b>x좌표가 가장 가까운 계약</b>에 배정한다.
    반환: {회사명: {'samang': 사망최대값, 'cands': [본체 후보값…]}}
    """
    out={}
    try:
        _CO=re.compile(r'[가-힣A-Za-z]{1,7}\s?(?:생명|손보|화재|해상|라이프|손해보험|공제|증권)')
        blk=[]; flag=False
        for l in lines:
            if ('계약별가입정보' in l) or ('계약별 가입정보' in l): flag=True
            if flag: blk.append(l)
            if flag and ('안내 및 유의' in l or '별첨' in l): break
        # ★★★v253(김O구 실측 2026.07.26): <b>앵커 문자열에 의존하지 마라</b>.
        #   이 리포트는 4p 표의 <b>헤더·담보명 한글이 pdftotext에서 통째로 빠진다</b>
        #   (값·회사명은 정상 추출). 그래서 `'계약별가입정보'`를 못 찾아 blk=[] →
        #   <b>즉시 빈 dict 반환</b> → ci_sebu=None → CI 중대한OO가 하나도 안 들어갔다.
        #   → 앵커가 없으면 <b>회사 접미어가 2개 이상 있는 줄</b>을 직접 찾아 그 줄부터 읽는다.
        if not blk:
            for i,l in enumerate(lines):
                if len(list(_CO.finditer(l)))>=2:
                    blk=lines[i:i+60]
                    print('[v253 sebu_ci] 앵커 없음 → 회사명 줄(%d)로 진입' % i)
                    break
        if not blk: return out
        # ① 회사명 줄 = 회사 접미어가 2개 이상 등장하는 첫 줄
        cidx=-1; cos=[]
        for i,l in enumerate(blk[:14]):
            ms=list(_CO.finditer(l))
            if len(ms)>=2: cidx=i; cos=[(m.start(), re.sub(r'\s','',m.group())) for m in ms]; break
        if cidx<0 or not cos: return out
        cos.sort()
        # ★★좌측 '전체 가입현황' 표가 같은 줄에 있어 숫자가 섞인다(실측: 신한 사망이 23,100=전체합계로 오염).
        #   → 계약별가입정보 영역의 <b>x 경계</b>를 잡아 그 왼쪽 숫자·라벨은 전부 버린다.
        _xnum = cos[0][0] - 8      # 숫자 허용 하한
        _xlab = cos[0][0] - 26     # 담보군 라벨 허용 하한
        def _owner(x):
            best=None; bd=10**9
            for sx,nm in cos:
                d=abs(x-sx)
                if d<bd: bd=d; best=nm
            return best
        # ★v253 회사별 보험료(노이즈) 수집 — samang 폴백에서 제외하기 위해
        # ★v253: 보험료는 <b>'원'이 붙는다</b> → 영역 내 '원' 값 전부를 노이즈 집합으로 모은다
        #   (회사별로 나누면 동일 회사 계약이 여러 건일 때 첫 건만 잡혀 보험료가 사망으로 새어든다 — 실측).
        _PREMS=set()
        for l in blk:
            for mm in re.finditer(r'([\d,]{4,})\s*원', l):
                try: _PREMS.add(int(mm.group(1).replace(',','')))
                except: pass
        # ② 사망액은 '사망' 라벨 행에서, 본체 후보는 영역 내 전체 숫자에서 수집.
        #   ★계약별가입정보 표는 <b>라벨 행과 값 행이 어긋난다</b>(실측: '암' 라벨 47행 / 값 45행,
        #     '뇌혈관질환' 라벨 52행 / 값 51행). 라벨 기준으로 값을 묶으면 후보가 전멸한다.
        #   → 사망만 라벨로 잡고(정확도 확인됨), 나머지는 <b>x좌표로 계약에 배정한 전체 숫자</b>를 후보로 둔다.
        for l in blk[cidx+1:]:
            if not l.strip(): continue
            # ★노이즈 행 제외 — 보험료·보장기간·납입기간/횟수·상품명 줄은 담보 금액이 아니다.
            #   (실측 오염: 148,576(보험료) · 9999/2018(만기·가입연도) · 1701/1710/2510(상품코드))
            if re.search(r'보험료|보장기간|납입|회사|보험서비스|상품', l): continue
            if re.search(r'\d{4}\s*[.\-]\s*\d{2}', l): continue
            mk=re.search(r'(?<![가-힣])사망(?![가-힣])', l)
            _is_sam = bool(mk) and mk.start()>=_xlab
            for m in re.finditer(r'([\d][\d,]{2,})', l):
                if m.start() < _xnum: continue          # ★좌측 '전체 가입현황' 표 숫자 차단
                _g=m.group(1)
                if ',' not in _g and re.fullmatch(r'(?:1[6-9]|2[0-9])\d{2}', _g): continue  # 연도·상품코드
                try: v=int(_g.replace(',',''))
                except Exception: continue
                if not (0<v<=200000): continue
                nm=_owner(m.start())
                if not nm: continue
                d=out.setdefault(nm, {'samang':0, 'cands':[]})
                if _is_sam: d['samang']=max(d['samang'], v)
                else: d['cands'].append(v)
        # ★★★v253 samang 폴백(김O구 실측 2026.07.26): 담보명 라벨 한글이 빠진 PDF에서는
        #   '사망' 라벨 행을 못 찾아 samang=0이 된다. 세부가입현황 계약열은 <b>행 순서가 고정</b>
        #   (사망→후유장해→실손→암…)이라 <b>보험료 제거 후 첫 값이 사망</b>이다.
        # ★★★그 첫 값 = <b>일반/질병사망만</b>(지점장 확정): 두 번째 값은 <b>상해사망</b>이고
        #   <b>CI 법칙에 상해사망은 전혀 상관없다</b>. 실측 교보 4,000(질병)/10,000(상해) → <b>4,000</b>.
        for _co, _d in out.items():
            if float(_d.get('samang') or 0) > 0: continue
            _pool = [v for v in (_d.get('cands') or [])
                     if v >= 1000 and v not in (_PREMS or set()) and v % 100 == 0]
            if _pool:
                _d['samang'] = float(_pool[0])
                print('[v253 samang폴백] %s 사망보장 %s (상해사망은 CI 무관 → 제외)' % (_co, format(_pool[0], ',')))
        return out
    except Exception as _e:
        print(f'[v237 sebu_ci ERR] {_e}')
        return out


def parse_sebu(lines):
    """v30x 세부가입현황 전담보 파서 = 주 소스(오버랩 근절). 탭 좌우2열, 담보-금액 경계 명확."""
    out={}
    sb=[]; flag=False
    for l in lines:
        if ('세부가입현황' in l) or ('전체 가입현황' in l): flag=True; continue
        if flag and (('계약별 가입정보' in l) or ('계약별가입정보' in l) or ('안내 및 유의' in l)):
            if out: break
            continue
        if flag: sb.append(l)
    if not sb: return out
    def _num(x):
        x=str(x).replace(',','').replace('원','').strip()
        try:
            v=float(x); return v if v>0 else None
        except: return None
    _skip=('충분','부족','미가입','구분','담보명','보장지수','가입금액','-')
    for raw in sb:
        # ★표준금액/권장금액 줄 = 가입금액 아님, 반영 금지(지점장 정본)
        if ('표준금액' in raw) or ('권장금액' in raw) or ('적정금액' in raw): continue
        cells=[c.strip() for c in raw.split(chr(9)) if c.strip()!='']
        k=0
        while k < len(cells):
            nm=cells[k]
            _ko=re.sub(r'[^가-힣]','',nm)
            if len(_ko)>=2 and nm not in _skip:
                val=None
                for off in (1,2,3):
                    if k+off>=len(cells): break
                    c2=cells[k+off]
                    mm=re.match(r'(충분|부족|미가입)\s*([\d,]+)?$', c2)
                    if c2 in ('충분','부족'):
                        if k+off+1<len(cells): val=_num(cells[k+off+1])
                        break
                    elif c2=='미가입' or c2=='-':
                        break
                    elif mm and mm.group(1)!='미가입' and mm.group(2):
                        val=_num(mm.group(2)); break
                    else:
                        vv=_num(c2)
                        if vv is not None: val=vv; break
                if val is not None:
                    _std,_j=resolve_kw(nm)
                    if _std and _std!='__무시__':   # ★v30z 무시지정 담보(전이암·고액항암)는 세부보충에서도 제외
                        if _std not in out or out[_std]<val: out[_std]=val
            k+=1

    return out

# ══════════════════════════════════════════════════════════════════════════
# ★★v44 신정원 3열 포맷 파서 (KB '상품별 가입담보상세' / 메리츠 '별첨 상품별 보험가입현황')
#   지점장 확정 2026.07.13: (1) 3열 포맷을 입력 정본에 추가 (앵커 자동감지, 기존 2열과 병행)
#                          (2) 담보명 정본 = 회사담보명
#                          (3) 갱신판정 = 상품명/담보명 '갱신' 표기 단독 기준 (3열엔 총회차 없음)
#   담보행: NO | 구분(정액/실손) | 회사담보명 | 신정원담보명 | 가입금액
# ══════════════════════════════════════════════════════════════════════════
_SJ_HEAD = re.compile(r'^\s*(\d{1,2})\s+(정액|실손)\s+(.+?)\s*$')
_SJ_AMT  = re.compile(r'\s{2,}((?:\d[\d,]*\s*[억만천]\s*)+)$')          # 행 끝 금액(1억 5,000만 등 공백 허용)
_SJ_CONT = re.compile(r'^\s+\S.*?\s{2,}((?:\d[\d,]*\s*[억만천]\s*)+)$')  # 신정원담보명 줄바꿈 wrap → 다음 줄 금액
_SJ_KB   = re.compile(r'^[\s\f]*(\S.*?)\s{2,}\|\s*가입일자\s*:\s*(\d{4})[-.](\d{2})[-.](\d{2})\s*\|')
_SJ_MZ   = re.compile(r'^[\s\f]*별첨\s+상품별\s*보험가입현황')

def _amt_kr(s):
    """한글 금액('1억 5,000만','2,000만','1억','1천') → 만원 단위 정수. 미해석=None(추측 금지)."""
    t = re.sub(r'[\s,]', '', str(s or ''))
    m = re.fullmatch(r'(?:(\d+)억)?(?:(\d+)만)?(?:(\d+)천)?', t)
    if not m or not any(m.groups()): return None
    v = int(m.group(1) or 0) * 10000 + int(m.group(2) or 0) + int(m.group(3) or 0) * 1000
    return v if 0 < v <= 200000 else None

def sinjeong_count(lines):
    """3열 신정원 계약 헤더 앵커 개수(KB '| 가입일자 : |' + 메리츠 '별첨 상품별 보험가입현황')."""
    return sum(1 for l in lines if _SJ_KB.match(l)) + sum(1 for l in lines if _SJ_MZ.match(l))

def sinjeong_detect(lines):
    """3열 신정원 포맷 감지. 계약 헤더 앵커가 2개 이상이면 3열로 확정."""
    return sinjeong_count(lines) >= 2

_SJC = {}

def _sj_fixname(name, sj, comp, prod):
    """★v98 3열(KB·메리츠) 전용 담보명 정규화. 롯데 2열 경로는 이 함수를 타지 않는다.
       한장보장표(등식1)와 어긋나던 실측 6건을 회사담보명 표기 보정으로만 해결."""
    r = re.sub(r'\s', '', str(name)); s2 = re.sub(r'\s', '', str(sj))
    c = re.sub(r'\s', '', str(comp)); pr = re.sub(r'\s', '', str(prod))
    # F4 생보 종신 주계약: 일사=일반사망 / 재사=재해(상해)사망
    if '일사보험금' in r: return '일반사망_주계약'
    if '재사보험금' in r: return '재해사망_주계약'
    # F2 타인'사망'교통사고처리지원금 → 교통상해사망 오분류 차단(처리지원금=합의금)
    if '처리지원금' in r and '사망' in r: return re.sub('사망', '', name)
    # F3 질병 입원 수술비Ⅱ/Ⅲ/Ⅳ = 질병수술비 합산군(한장보장표 기준)
    if re.match(r'^질병입원수술비[ⅡⅢⅣⅤ]', r): return '질병입원수술비' + r[8:]
    # F6 생보 암/CI의 '뇌혈관진단' = 실제 뇌출혈(지침 §8.3)
    if r == '뇌혈관진단' and ('생명' in c or 'AIA' in c.upper()): return '뇌출혈진단비'
    # F8 자동차사고 부상치료지원금Ⅱ 등 변형 → 자부상(신정원담보명 근거)
    if '부상위로금' in s2 or '부상치료' in s2: return '자동차사고부상위로금'
    # F9 1-5종 재해수술 = 상해 종수술(질병 종수술 행 중복산입 차단)
    if '종재해수술' in r: return name.replace('재해수술', '상해수술', 1)
    # F7 실손 '의료비(입원+통원)' 통합형 → 실손 입원 행 (F5b보다 먼저 판정)
    if '의료비' in r and '입원' in r and '통원' in r: return '실손입원의료비'
    # F5b 암 통원 담보는 진단비 행에 산입 금지
    if '암통원' in r or '암통원' in s2 or (('통원' in s2) and ('암' in r)): return '[확인] 통원 ' + name
    return name

def _sj_rows(block):
    """담보행 → [(회사담보명, 만원정수)]. 금액 미해석 담보는 0 + [확인] 프리픽스(누락 금지)."""
    out = []; i = 0
    while i < len(block):
        h = _SJ_HEAD.match(block[i])
        if not h: i += 1; continue
        rest = h.group(3); amt_s = None
        m = _SJ_AMT.search(rest)
        if m:
            amt_s = m.group(1); rest = rest[:m.start()]
        else:                                            # 신정원담보명 wrap → 다음 1~2줄에서 금액 회수
            for k in (1, 2):
                if i + k < len(block):
                    c = _SJ_CONT.match(block[i + k])
                    if c: amt_s = c.group(1); break
        parts = [p for p in re.split(r'\s{2,}', rest.strip()) if p]
        i += 1
        if not parts: continue
        name = re.sub(r'\s+', ' ', parts[0].strip())     # ★정본: 담보명 = 회사담보명(parts[0])
        sj   = re.sub(r'\s+', ' ', parts[1].strip()) if len(parts) > 1 else ''
        if len(name) < 2: continue
        v = _amt_kr(amt_s) if amt_s else None
        if v is None:
            out.append(('[확인] 금액판독불가 ' + name, sj, 0)); continue
        name = _sj_fixname(name, sj, _SJC.get('c',''), _SJC.get('p',''))
        if '특정암진단' in re.sub(r'\\s','',sj) and '유사암' not in sj:
            # ★v197(2026.07.23): 신정원 '특정암진단' = 고액암 행으로 확정(구 v98 F5 [확인]큐 폐기)
            # ★★★★★v247 (KB 3열 실측): <b>회사담보명이 유사암 4종이면 리네임하지 않는다</b>.
            #   정본 = <b>담보명 정본은 회사담보명</b>(신정원담보명은 분류 참고용)이다.
            #   실측 — 메리츠 `갱신형 갑상선암(초기제외)진단비` · `갱신형 갑상선암 및 기타피부암의 전이암…진단비`가
            #   신정원명 '특정암진단' 때문에 <b>고액암 2,000</b>으로 갔다(KB 전체보장현황 고액암은 1,000=라이나뿐).
            #   유사암 4종(갑상선·갑상샘·기타피부·경계성·제자리·상피내)은 <b>유사암 행</b>이 정답이다.
            _nm0 = re.sub(r'\s', '', str(name))
            if any(k in _nm0 for k in ('갑상선','갑상샘','기타피부','경계성','제자리','상피내')):
                pass
            elif '특정암' not in name:
                name = '특정암진단비(' + name + ')'
        out.append((name, sj, v))
    # ★v44 실측보정: DB 실손처럼 회사담보명이 '질병(전체질병을 의미)' 하나로 3행(입원·통원·약값)이 겹치는 경우
    #    회사담보명만 쓰면 dict 키 충돌 → 합산 사고. 계약 내 중복 회사담보명은 신정원담보명으로 분리한다.
    cnt = {}
    for nm, sj, v in out: cnt[nm] = cnt.get(nm, 0) + 1
    fixed = []
    for nm, sj, v in out:
        if cnt.get(nm, 0) > 1 and sj and sj != nm:
            fixed.append((_sj_fixname(sj, sj, _SJC.get('c',''), _SJC.get('p','')), v))  # ★v98 채택 후 재정규화
        else:
            fixed.append((nm, v))
    # ★★★v208 (지점장 확정 2026.07.25, 양*선 삼성 New내돈내삼 실측): 위에서 신정원담보명으로 갈아탄 뒤
    #   <b>신정원담보명끼리 또 중복</b>이면 1인실·2~3인실 구분이 통째로 사라져 <b>합산 사고</b>가 난다.
    #   실측: 상급종합 1인실 20 + 상급종합 2~3인실 5 → '상급종합병원 질병입원일당' 25 /
    #         종합 1인실 10 + 종합 2~3인실 5 → '종합병원이하 질병입원일당' 15 → 최종 질병종합병원일당 <b>40</b>.
    #   → 회사담보명에서 <b>병실 토큰(1인실 · 2~3인실 등)</b>만 뽑아 접미로 붙여 되살린다.
    cnt2 = {}
    for nm, v in fixed: cnt2[nm] = cnt2.get(nm, 0) + 1
    if any(c > 1 for c in cnt2.values()):
        fixed2 = []
        for (nm, v), (onm, osj, _ov) in zip(fixed, out):
            if cnt2.get(nm, 0) > 1:
                _rm = re.search(r'(\d+\s*~\s*\d+\s*인실|\d+\s*인실)', str(onm))
                if _rm: nm = '%s(%s)' % (nm, re.sub(r'\s', '', _rm.group(1)))
            fixed2.append((nm, v))
        fixed = fixed2
    return fixed

def _sj_unwrap(block):
    """★v98: pdftotext가 가운뎃점(·)에서 줄을 끊는다.
       '항암방사선·' / '질병 입원 간호·' 처럼 ·로 끝나는 줄은 다음 줄 첫 필드를 붙여 복원한다.
       (2~3줄 연속 랩도 while로 처리. 담보명 통째 유실 방지.)"""
    out = []; i = 0; n = len(block)
    while i < n:
        cur = block[i].rstrip()
        guard = 0
        while cur.endswith('\u00b7') and i + 1 < n and guard < 4:
            nxt = block[i + 1].strip()
            if not nxt: break
            mm = re.match(r'^(\S+(?:\s\S+)*?)(\s{2,}.*)?$', nxt)
            if not mm: break
            cur = cur + mm.group(1) + (mm.group(2) or '')
            i += 1; guard += 1
        out.append(cur); i += 1
    return out

def parse_sinjeong(lines):
    # ★v244: 3열도 세부가입현황(상세내역)을 CI 선지급률 2순위 근거로 쓴다(2열과 동일).
    try:
        _SJ_SEBU = {re.sub(r'\s','',k): v for k, v in (parse_sebu_ci(lines) or {}).items()}
    except Exception:
        _SJ_SEBU = {}
    """KB·메리츠 3열 리포트 → contracts[]. 표 구조 동일 → 파서 1개로 두 채널 커버."""
    lines = _sj_unwrap(lines)          # ★v98 F1: 가운뎃점 줄바꿈 복원
    n = len(lines)
    heads = []                                            # (idx, company, product, 가입일)
    for i, l in enumerate(lines):
        m = _SJ_KB.match(l)
        if m:                                             # ── KB: '회사명 ... | 가입일자 : YYYY-MM-DD |'
            comp = re.sub(r'\s+', '', m.group(1).strip())
            prod = ''
            for j in range(i + 1, min(i + 6, n)):
                s = lines[j].strip()
                if s and not re.search(r'가입일자|계약자|보험기간', s):
                    prod = s; break
            heads.append((i, comp, prod, f'{m.group(2)}.{m.group(3)}.{m.group(4)}'))
            continue
        if _SJ_MZ.match(l):                               # ── 메리츠: '별첨 상품별 보험가입현황' → 다음 2줄=회사·상품
            got = []
            for j in range(i + 1, min(i + 8, n)):
                s = lines[j].strip()
                if not s: continue
                if re.search(r'계약자|보험기간|가입담보명', s): break
                got.append(s)
                if len(got) == 2: break
            if len(got) >= 1:
                heads.append((i, re.sub(r'\s+', '', got[0]), (got[1] if len(got) > 1 else ''), ''))
    if not heads: return []

    contracts = []
    for hi, (idx, company, product, join_d) in enumerate(heads):
        end = heads[hi + 1][0] if hi + 1 < len(heads) else n
        block = lines[idx:end]
        _SJC['c'] = company; _SJC['p'] = product     # ★v98 _sj_fixname 컨텍스트
        ht = ' '.join(block[:12])
        contract_date = expiry_date = pay_period = ''; premium = 0
        md = re.search(r'(\d{4})[-.](\d{2})[-.](\d{2})\s*~\s*(\d{4})[-.](\d{2})[-.](\d{2})', ht)
        if md:
            contract_date = f'{md.group(1)}.{md.group(2)}.{md.group(3)}'
            expiry_date   = f'{md.group(4)}.{md.group(5)}.{md.group(6)}'
        if not contract_date and join_d: contract_date = join_d
        mp = re.search(r'([\d,]{4,})\s*원', ht)
        if mp:
            try:
                pv = int(mp.group(1).replace(',', ''))
                if 1000 < pv < 5000000: premium = pv
            except: pass
        mper = re.search(r'(?:월납|매월납)\s*/\s*(\d+)\s*년', ht)
        if mper: pay_period = f'{mper.group(1)}년납'
        elif re.search(r'일시납', ht): pay_period = '일시납'   # ★v127 3열 리포트 일시납 포착
        pay_count = ''                                          # ★v127 3열은 납입회차 칸이 없다
        if not expiry_date and re.search(r'종신', ht): expiry_date = '9999.12.31'
        if is_excluded(company, product, contract_date, expiry_date, pay_period, pay_count): continue   # 제외 5~7종

        dambo = {}
        # ★★★★★v244 (지점장 질문 "CI 대응 다 된 거냐" → 점검 중 발견): <b>3열(KB·메리츠) 경로에
        #   `ci_lines`가 통째로 빠져 있었다</b>. 정본 "규칙 하나 고치면 2열·3열 두 경로를 반드시 다 돌린다"를
        #   내가 지키지 않아, v239~v243의 CI 수정(줄단위 사망·중대한OO 본체·뇌 축 판별)이 <b>롯데에만</b> 적용됐다.
        #   → KB·메리츠 CI 계약은 여전히 <b>무조건 '중대한 뇌졸증'</b>으로 가서 같은 오류가 재발한다.
        # ★★★★★v246 (지점장 확정 2026.07.25): "<b>이건 롯데나 KB나 동일하다</b>" — CI 4단계·뇌 축·사망 배정은
        #   2열(롯데 [별첨] 보험서비스(상품)별 보장현황)과 3열(<b>KB '상품별 가입담보상세'</b>) <b>동일 적용</b>.
        #   ★3열은 앵커 줄(`| 가입일자 : |`)에서 <b>회사명 칸에 상품명이 붙어 오는 경우</b>가 있고
        #     상품명은 그 <b>다음 줄</b>에서 뽑는다 → <b>회사명·상품명 둘 다</b> CI 판정에 넣는다(한쪽만 보면 놓친다).
        _sjci = _isci_prod(product) or _isci_prod(company)
        ci_lines = {'samang': [], 'cands': [], 'jungdae': [], 'brain': []}
        for nm, v in _sj_rows(block):
            if re.search(r'납입면제|납입지원', nm): continue
            dambo[nm] = dambo.get(nm, 0) + v
            if _sjci and 0 < v <= 200000:
                _n2 = re.sub(r'\s', '', str(nm))
                # ★★★★★v253 지점장 지시 2026.07.26: <b>재해사망특약 6,000 → samang에서 삭제</b>.
                #   + 2차 규칙 원문 "세부가입현황에서 '사망'에서 <b>상해사망은 제외</b> 질병사망에서
                #   50% or 80%형 금액이 있는지 찾아본다".
                #   → CI 사망보장 후보 = <b>일반사망 / 질병사망(주계약)</b>. <b>상해사망·재해사망은 제외</b>.
                #   ★지점장이 말한 이 두 개만 제외한다 — 교통·외래 등 다른 항목을 임의로 덧붙이지 말 것.
                if (re.search(r'사망', _n2) and not re.search(r'후유|장해', _n2)
                        and not re.search(r'상해|재해', _n2)):
                    ci_lines['samang'].append(v)
                elif _n2.startswith('중대한'):
                    ci_lines['jungdae'].append(v)
                elif re.search(r'진단', _n2):
                    ci_lines['cands'].append(v)
                    if '뇌출혈' in _n2:   ci_lines['brain'].append(('뇌출혈', v))
                    elif '뇌졸' in _n2:   ci_lines['brain'].append(('뇌졸증', v))
        if not dambo: continue

        # ★확정(3) 갱신판정: 만기 9999(종신)=비갱신 / 상품명 '갱신' 표기=갱신 / 그 외는 후처리 담보절반 규칙
        if _is_silson_prod(company, product): renewal = '갱신'   # ★★★실손=무조건 갱신(v103 영구지침)
        elif str(expiry_date).startswith('9999'): renewal = '비갱신(종신)'
        elif '갱신' in str(product) and '비갱신' not in str(product): renewal = '갱신'
        else: renewal = '비갱신'

        ci_jugye = []
        if _sjci:
            for bl in block:
                for m2 in re.finditer(r'(?<![_가-힣])주계약\s+([\d,]{3,})', bl):
                    try:
                        v2 = int(m2.group(1).replace(',', ''))
                        if 0 < v2 <= 200000: ci_jugye.append(v2)
                    except: pass
        contracts.append({'company': company, 'product': product, 'contract_date': contract_date,
                          'expiry_date': expiry_date, 'premium': premium, 'pay_period': pay_period,
                          'pay_count': '', 'renewal': renewal, 'dambo': dambo,
                          'ci_jugye': ci_jugye, 'ci_extra': [], 'ipwon': [], '_sj': True,
                          'ci_lines': ci_lines, 'ci_sebu': (_SJ_SEBU or {}).get(re.sub(r'\s','',str(company or '')))})
    print(f'[SINJEONG] 3열 포맷 감지 → 계약 {len(contracts)}건 / 담보 {sum(len(c["dambo"]) for c in contracts)}개')
    return contracts


def parse_txt(txt, filename=''):
    lines = [l.rstrip() for l in txt.replace('\r\n','\n').replace('\r','\n').split('\n')]
    # ★★★v237: 세부가입현황(상세내역) 계약별 CI 정보 1회 계산 — 선지급률 판정 2순위 근거
    _cib = ci_selftest()
    print('[v238 CI자가진단] ' + ('PASS %d/%d' % (len(_CI_SELFTEST)-len(_cib), len(_CI_SELFTEST)) if not _cib else 'FAIL ' + ' | '.join(_cib[:6])))
    try: _SEBU_CI = parse_sebu_ci(lines)
    except Exception: _SEBU_CI = {}
    if _SEBU_CI: print(f'[v237 sebu_ci] 계약 {len(_SEBU_CI)}건 파싱: ' + ', '.join(f"{k}(사망{v['samang']:,})" for k,v in _SEBU_CI.items()))
    client = ''
    # ★ 정본 §2: 고객명 = 파일명 우선
    if filename:
        base = re.sub(r'\.(?:[Tt][Xx][Tt]|[Pp][Dd][Ff])$', '', filename).strip()
        fm = re.match(r'^([가-힣]{2,4})', base)
        # ★v44: '20260713_백O화님_보장분석' 처럼 날짜·숫자 접두 파일명 대응(구 정규식은 '고객'으로 낙하)
        if not fm: fm = re.search(r'([가-힣][가-힣A-Za-z*O]{1,3})님', base)
        if fm: client = re.sub(r'님$', '', fm.group(1))
    # 폴백: 내용에서 (마스킹 '박*은' 형태도 허용)
    if not client:
        for l in lines[:30]:
            l = l.strip()
            m2 = re.search(r'([가-힣*]{2,4})\s*고객님', l) or re.search(r'([가-힣*]{2,4})\s*님의', l)
            if m2: client = m2.group(1); break
            m = re.match(r'^([가-힣]{2,5})\s*$', l)
            if m and len(m.group(1)) <= 4: client = m.group(1); break
    if not client: client = '고객'

    # ★ 한장보장표(앞부분)에서 회차 추출 → (회사,가입일,만기일) 키로 맵 구축 (별첨엔 회차 없음)
    paycount_map = {}
    for l in lines:
        ld = l.strip()
        m = re.search(r'([가-힣A-Za-z]{2,8}(?:생명|화재|손보|손해|해상|라이프|증권)?)\s+.*?(\d{4}\.\d{2}\.\d{2})\s+(\d{4}\.\d{2}\.\d{2})\s+월납\s+(\d{1,3}/\d{2,3})', ld)
        if m:
            comp = m.group(1).strip()
            paycount_map[(comp, m.group(2), m.group(3))] = m.group(4)
            paycount_map[(m.group(2), m.group(3))] = m.group(4)  # 회사 표기 흔들림 대비 보조키

    contracts = []; i = 0; n = len(lines)
    # ★★v44 분기: 신정원 3열 포맷(KB·메리츠) 감지 시 parse_sinjeong 사용, 아니면 기존 2열 별첨 엔진.
    _SJN   = sinjeong_count(lines)
    _IS_SJ = _SJN >= 2
    if _IS_SJ:
        contracts = parse_sinjeong(lines)
        i = n                      # 기존 2열 루프 스킵('정상계약 리스트' 문자열 부재로 어차피 0건)
    while i < n:
        l = lines[i].strip()
        if '실효계약 리스트' in l or '미납해지' in l: break
        if '정상계약 리스트' not in l: i += 1; continue
        i += 1
        while i < n and not lines[i].strip(): i += 1
        if i >= n: break
        # ★v30p+ 형태 자동 감지: 첫 줄에 계약자/탭/가입금액/Chtd 있으면 신형(다중헤더), 아니면 정상형(기존).
        _first = lines[i]
        if ('계약자' in _first) or ('\t' in _first) or ('가입금액' in _first) or ('Chtd' in _first):
            # ── 신형 (오늘 PDF 업데이트 형태, 별첨 헤더 A/B/C 혼재) ──
            # ★v30p 다중 별첨 헤더 형태 대응(A:계약자줄먼저→회사·상품 다음줄 / B:회사·상품+계약자 한줄 / C:계약자이름+회사+상품 한줄).
            #   담보표 헤더('가입금액'·'담보명'·'Chtd') 전까지를 헤더영역으로 모아 회사·상품·보험료·날짜를 통째 추출.
            _hdr=[]; _k=i
            while _k < n and _k < i+8:
                _lk = lines[_k]
                if '가입금액' in _lk or 'Chtd' in _lk or '담보명' in _lk: break
                if '정상계약 리스트' in _lk or '실효계약 리스트' in _lk: break
                # ★v59 첫 담보 유실 방지: '한글…+끝자리 숫자'=담보값 줄을 만나면 헤더수집 중단.
                #   (계약자·보험료·보장기간·납입·날짜 줄은 담보로 오인 안 되게 제외.)
                #   이전 고정 6줄창은 블랭크패딩 때문에 첫 담보(예 한화생명 암수술 50)를
                #   상품명에 흡수 → 담보 1~3개 유실. 생보 암보험(첫 담보=암진단)이 가장 큰 피해.
                _lks=_lk.strip()
                # ★★★v234 (2026.07.25 한정환 실측): 상품명 줄이 '담보값 줄'로 오인돼 헤더수집이 끊기던 버그.
                #   `무배당 마이라이프 한아름종합보험 1710` / `참좋은운전자상해보험 2510` 처럼
                #   상품명 끝에 <b>4자리 상품코드</b>가 붙으면 이 가드가 걸려 break → 상품명이 공란이 되고
                #   그 줄은 담보로 파싱돼 확인사항 큐에 `무배당 마이라이프 한아름종합보험 = 1710`이 박혔다.
                #   판별: 별첨 담보표는 pdftotext -layout에서 <b>금액 앞 공백이 2칸 이상</b>(표 열 구분)이고,
                #   1,000 이상은 <b>콤마</b>가 찍힌다. 반대로 상품코드는 <b>공백 1칸 + 콤마 없는 4자리</b>다.
                _mnum = re.search(r'(\s+)([\d][\d,]*)\s*$', _lks)
                _isprod = bool(_mnum) and len(_mnum.group(1))==1 and bool(re.fullmatch(r'\d{4}', _mnum.group(2)))
                if _lks and re.search(r'[가-힣]', _lks) and _mnum and not _isprod \
                   and not re.search(r'계약자|보험료|보장기간|납입|월납|\d{4}\.\d{2}', _lks):
                    break
                _hdr.append(_lk); _k += 1
            _ht = ' '.join(_hdr).replace('\t',' ')
            i = _k+1 if (_k<n and ('가입금액' in lines[_k] or 'Chtd' in lines[_k] or '담보명' in lines[_k])) else _k
            contract_date = expiry_date = pay_period = pay_count = ''; premium = 0
            _md = re.search(r'(\d{4}\.\d{2}\.\d{1,2})\s*[-~（卜\s]+(\d{4}\.\d{2}\.\d{1,2})', _ht)
            if _md: contract_date=_md.group(1); expiry_date=_md.group(2)
            _mp = re.search(r'보험료\s*([\d,\.]+)\s*원', _ht)
            if _mp:
                try:
                    _pv=int(_mp.group(1).replace(',','').replace('.',''))
                    if 1000 < _pv < 5000000: premium=_pv
                except: pass
            _mper = re.search(r'월납\s*/?\s*(\d+)\s*년', _ht)
            if _mper: pay_period=f"{_mper.group(1)}년납"
            # ★v126 열 어긋남 수정: 납입주기 칸이 '일시납/연납/N개월납'인 행에서
            #   그 토큰이 회사명 앞에 붙어 '일시납AIG손보'로 오염되던 버그.
            elif re.search(r'일시납', _ht): pay_period = '일시납'
            elif re.search(r'연\s*납', _ht): pay_period = '연납'
            _mpc = re.search(r'(\d{1,3})\s*/\s*(\d{2,3})\s*회', _ht)
            if _mpc: pay_count=f"{_mpc.group(1)}/{_mpc.group(2)}"
            # 회사·상품 분리: 계약자·납입·보험료·보장기간·단위 boilerplate 제거 후 보험사 키워드로 split
            _ct = _ht
            _ct = re.sub(r'계약자\s*\S+',' ',_ct)
            _ct = re.sub(r'납입주기\s*/?\s*기간',' ',_ct)
            _ct = re.sub(r'보험료\s*[\d,\.]+\s*원',' ',_ct)
            _ct = re.sub(r'보장기\s*간|보장기간',' ',_ct)
            _ct = re.sub(r'\d{4}\.\d{2}\.\d{1,2}\s*[-~（卜\s]+\d{4}\.\d{2}\.\d{1,2}',' ',_ct)
            _ct = re.sub(r'월납\s*/?\s*\d+\s*년',' ',_ct)
            _ct = re.sub(r'일시납|연\s*납|월\s*납|\d+\s*개월납',' ',_ct)   # ★v126 납입주기 토큰 제거
            _ct = re.sub(r'\d{1,3}\s*/\s*\d{2,3}\s*회',' ',_ct)
            _ct = re.sub(r'[（(]?\s*단위\s*:?\s*만원\s*[）)]?',' ',_ct)
            _ct = re.sub(r'\s+',' ',_ct).strip()
            _mc = re.search(r'^(.*?(?:화재|손해보험|손보|해상|생명|라이프|증권))\s*(.*)$', _ct)
            if _mc:
                company = re.sub(r'\s','',_mc.group(1)); product = _mc.group(2).strip()
                # ★★★v234 (한정환 실측): 정규식 `.*?`가 최단일치라 '메트라이프생명'에서 '라이프'를 먼저 물어
                #   회사명이 <b>'메트라이프'</b>, 상품명이 <b>'생명 무배당 마스터플랜…'</b>으로 잘렸다.
                #   → 상품명이 회사 접미어로 시작하면 그 토큰을 회사명으로 되돌린다.
                _mtail = re.match(r'^(생명|화재|손해보험|손보|해상|라이프|증권|공제)\s+(\S.*)$', product)
                if _mtail:
                    company = company + _mtail.group(1); product = _mtail.group(2).strip()
            else:
                _parts=_ct.split(' ',1); company=_parts[0] if _parts else lines[i].strip(); product=_parts[1].strip() if len(_parts)>1 else ''
        else:
            # ── 정상형 (기존, '가입금액' 헤더 없이 회사→상품→계약자줄→담보) ──
            _head = lines[i].strip()          # ★v44 롯데: 이 줄에 '회사명 상품명'이 한 줄로 붙어 온다
            company = _head; i += 1
            contract_date = expiry_date = pay_period = pay_count = ''; premium = 0
            for _j in range(i, min(i+5, n)):
                _l = lines[_j]
                _m = re.search(r'(\d{4}\.\d{2}\.\d{2})\s*~\s*(\d{4}\.\d{2}\.\d{2})', _l)
                if _m: contract_date = _m.group(1); expiry_date = _m.group(2)
                _m2 = re.search(r'(\d{1,3})\s*/\s*(\d{2,3})\s*회', _l) or re.search(r'월납\s+(\d{1,3})\s*/\s*(\d{2,3})', _l)
                if _m2 and not pay_count: pay_count = f"{_m2.group(1)}/{_m2.group(2)}"
                _m3 = re.search(r'보험료\s*([\d,\.]+)\s*원', _l) or re.search(r'([\d,]+)원', _l)
                if _m3:
                    try:
                        _v = int(_m3.group(1).replace(',','').replace('.',''))
                        if 1000 < _v < 5000000: premium = _v
                    except: pass
                _m5 = re.search(r'월납\s*/?\s*(\d+)\s*년', _l)
                if _m5 and not pay_period: pay_period = f"{_m5.group(1)}년납"
            while i < n and not lines[i].strip(): i += 1
            product = ''
            # ★★v44 롯데 결함B 수정: 헤더 줄을 보험사 키워드로 회사/상품 분리.
            #    구버전은 company에 상품명이 통째로 남고, product 칸엔 '담보 첫 줄'이 들어가
            #    (1) 상품명 오염 (2) 담보 1개 유실 (3) 병합키(회사·보험료·상품[:12]) 불일치로
            #    같은 계약이 2건으로 쪼개지는 결함A까지 유발했다. 회사/상품만 바로잡으면 셋 다 해소.
            _mc = re.match(r'^(.*?(?:화재|손해보험|손보|해상|생명|라이프|증권|공제))\s+(\S.*)$', _head)
            if _mc:
                company = re.sub(r'\s', '', _mc.group(1))
                product = re.sub(r'\s+', ' ', _mc.group(2).strip())
            else:
                # 폴백(기존): 회사명만 있는 헤더 → 다음 줄들에서 상품명 탐색
                for _j in range(i, min(i+6, n)):
                    _l = lines[_j].strip()
                    if _l and not re.search(r'계약자|납입주기|보험료|보장기간', _l):
                        if len(_l) > 5 and not re.search(r'^[\d,]+$', _l) and not re.search(r'^\d{4}\.\d{2}', _l):
                            product = _l; i = _j + 1; break
        if is_excluded(company, product, contract_date, expiry_date, pay_period, pay_count):
            while i < n and '정상계약 리스트' not in lines[i] and '실효계약 리스트' not in lines[i]: i += 1
            continue
        renewal = judge_renewal(product, expiry_date, pay_count, contract_date, pay_period, company)
        # 담보 블록 텍스트 수집 (다음 '정상계약/실효계약 리스트'까지)
        block_lines = []; j = i
        while j < n:
            if '정상계약 리스트' in lines[j] or '실효계약 리스트' in lines[j]: break
            block_lines.append(lines[j]); j += 1
        i = j
        # 추출: LLM 우선(깨진 별첨 복원), 키 없거나 실패 시 규칙 폴백
        dambo = llm_extract('\n'.join(block_lines)) or rule_extract(block_lines)
        # ★ CI/리빙케어/GI: 별첨이 전부 '주계약'으로 라벨없이 뭉침 → 개별 주계약 금액 수집(본체 80/50% 판별용)
        ci_jugye=[]
        if _isci_prod(product):
            for _bl in block_lines:
                # ★v33 pdftotext -layout 다열 레이아웃: 한 줄에 2~3담보 → finditer.
                #    '_주계약'(일반사망_주계약 등) 은 lookbehind 로 차단.
                for _m in re.finditer(r'(?<![_가-힣])주계약\s+([\d,]{3,})', _bl):
                    try:
                        _v=int(_m.group(1).replace(',',''))
                        if 0<_v<=200000: ci_jugye.append(_v)
                    except: pass
            # ★v30z2 삼성 리빙케어형: 주계약이 '주계약'이 아니라 상품명(삼성리빙케어(종신2종)1.2)으로 라벨됨.
            #   '리빙케어 종신N종' 플랜패턴 뒤 금액만 추출(리빙케어보장특약·재해사망 등 부담보 오수집 차단).
            #   콤마·마침표 천단위(3.000=3,000) 정규화. 80/50 판별은 아래 기존 로직이 수행.
            if not ci_jugye and '리빙케어' in (product or ''):
                for _bl in block_lines:
                    for _mm in re.finditer(r'리빙케어\s*[（(]?\s*종신\d*종[）)]?\s*\d*[.．]?\d*[\s\t]+([\d][\d.,]{2,})', _bl):
                        try:
                            _v=int(_mm.group(1).replace(',','').replace('.',''))
                            if 100<=_v<=200000: ci_jugye.append(_v)
                        except: pass
            # ★★★v122(2026.07.21 장문순 실측): 삼성생명 CI는 주계약이 '주계약'이라는 글자로
            #   안 나오고 <상품명 자체>로 라벨된다.
            #   예) '퍼펙트통합보험프리미엄,無표준월납  3,000 / ... 1,500'
            #   → 상품명 앞 6자를 키로 삼아 그 라벨이 붙은 금액을 주계약 후보로 수집한다.
            #   (리빙케어 전용 분기 v30z2의 일반화 — 퍼펙트플러스·퍼펙트통합 등 전 CI 상품 적용)
            if not ci_jugye:
                _pk = re.sub(r'[^가-힣A-Za-z0-9]', '', str(product or ''))[:6]
                if len(_pk) >= 4:
                    for _bl in block_lines:
                        for _mm in re.finditer(r'([가-힣A-Za-z0-9,\.\(\)（）Ⅰ-Ⅹ無·]{4,}?)\s{2,}([\d][\d,]{2,})', _bl):
                            _lab = re.sub(r'[^가-힣A-Za-z0-9]', '', _mm.group(1))
                            if _lab.startswith(_pk):
                                try:
                                    _v = int(_mm.group(2).replace(',', ''))
                                    if 100 <= _v <= 200000: ci_jugye.append(_v)
                                except Exception:
                                    pass
        # ★v29t CI추가보장특약: 줄별 값 수집(병합 전 원값 보존)
        ci_extra=[]
        for _bl in block_lines:
            for _m in re.finditer(r'CI\s*추가보장특약\s+([\d,]+)', _bl):
                try:
                    _v=int(_m.group(1).replace(',',''))
                    if 0<_v<=200000: ci_extra.append(_v)
                except: pass
        # ★v29t 생보 입원특약 일당: 줄별 값 수집(병합 전 원값 보존)
        ipwon=[]; _ipkey=None
        # ★★★★★v239 (지점장 지적 2026.07.25): <b>CI 주계약 사망은 별첨 '줄 단위 개별값'이다 — 합산 금지</b>.
        #   실측 증거(DB생명 CI종신 1701): 별첨에 <b>질병사망 3,000 / 질병사망 2,000</b> 2줄 = 주계약 3,000 + 추가특약 2,000.
        #   그런데 dambo는 이를 <b>합산해 5,000</b>으로 만든다 → 2,400/5,000=48%가 되어 50%·80% 어디에도 안 맞고
        #   [확인]으로 빠졌다. <b>주계약 3,000으로 보면 2,400/3,000 = 정확히 80%</b>(뇌출혈·급성심근·암·특정질병·치매가
        #   전부 2,400으로 일치) → <b>80%형이 정답</b>이었다.
        #   → CI 계약은 <b>접힘 복원된 줄에서 사망 줄과 담보 줄의 원값을 따로 수집</b>한다.
        ci_lines={'samang':[], 'cands':[], 'jungdae':[], 'brain':[]}
        # ★★★v234: 별첨이 <b>다열</b>이면 원본 줄은 `재해장해특약  6,000   입원  3` 처럼
        #   두 담보가 한 줄에 붙어 있어 `^입원 …$` 앵커가 절대 맞지 않는다(실측 메트라이프 9p).
        #   → rule_extract와 <b>동일한 접힘 복원</b>을 거친 1담보 1줄 형태에서 스캔한다.
        try:
            _bl_flat = _split_cols(_reflow_cols(_unfold_cols(block_lines)))
        except Exception:
            _bl_flat = block_lines
        for _bl in _bl_flat:
            # ★★★v234 (한정환 메트라이프생명 실측): 별첨에 <b>수식어 없는 '입원'</b>만 인쇄되는 생보가 있다.
            #   구 정규식은 '입원특약'만 잡아 `입원 3` 2줄을 일당 어느 행에도 못 넣었다
            #   (세부가입현황 4p 입원비 칸 = `3 | 3` = 질병 3 · 상해 3, 한장보장표 일당 13 = 신한 10 + 메트 3).
            # ★★★v239 CI 주계약/본체 후보 = 줄 단위 원값(합산 전)
            if _isci_prod(product):
                _mc = re.match(r'^\s*(\S.*?)\s{2,}([\d,]{3,})\s*$', _bl.rstrip()) or \
                      re.match(r'^\s*(\S.*?)\s+([\d,]{3,})\s*$', _bl.rstrip())
                if _mc:
                    _nmc = re.sub(r'\s','',_mc.group(1))
                    try: _vc = int(_mc.group(2).replace(',',''))
                    except Exception: _vc = 0
                    if 0 < _vc <= 200000 and re.search(r'[가-힣]', _nmc):
                        # ★v253 상해사망·재해사망 제외(지점장 지시) — 위 3열 경로와 동일
                        if (re.search(r'사망', _nmc) and not re.search(r'후유|장해', _nmc)
                                and not re.search(r'상해|재해', _nmc)):
                            ci_lines['samang'].append(_vc)
                        elif _nmc.startswith('중대한'):
                            # ★★★★★v240(지점장 지시 2026.07.25): 별첨에 <b>'중대한OO' 담보가 명시</b>되어 있으면
                            #   <b>그 금액이 곧 CI 본체(선지급액)</b>다. 실측 신한 `중대한화상진단 2,000`.
                            ci_lines['jungdae'].append(_vc)
                        elif re.search(r'진단', _nmc):     # 본체 후보 = 진단비 담보만
                            ci_lines['cands'].append(_vc)
                            # ★★★★★v242: 뇌 담보는 <b>축(뇌출혈/뇌졸증)과 금액</b>을 따로 기록한다.
                            #   CI 본체를 '중대한 뇌졸증'에 넣을지 '중대한 뇌출혈'에 넣을지 판별하기 위함.
                            if '뇌출혈' in _nmc:   ci_lines['brain'].append(('뇌출혈', _vc))
                            elif '뇌졸' in _nmc:   ci_lines['brain'].append(('뇌졸증', _vc))
            _m=re.match(r'^\s*(입원특약|입원)\s+([\d,]+)\s*$', _bl.strip())
            if _m:
                _ipkey=_m.group(1)
                _m=re.match(r'^\s*(?:입원특약|입원)\s+([\d,]+)\s*$', _bl.strip())
                try:
                    _v=int(_m.group(1).replace(',',''))
                    if 0<_v<=1000: ipwon.append(_v)
                except: pass
        # ★v29t (지점장 확정 2026.07.02): 생보 '입원특약' 일당 = 상해·질병 둘 다 해당 →
        #   질병일당·상해일당 두 행에 한 줄 값(중복줄=동일특약 재출현 → max) 각각 기재. dambo 변환이라 엑셀·PPT 동일 반영.
        if ipwon and any(k in (company or '') for k in ('생명','라이프','AIA','메트라이프','우체국','공제')) and (_ipkey in dambo):
            _v1=max(ipwon)   # ★중복줄(질병축·상해축)=동일 특약 → max
            dambo.pop(_ipkey, None)
            dambo['질병입원일당(입원특약)']=dambo.get('질병입원일당(입원특약)',0)+_v1
            dambo['상해입원일당(입원특약)']=dambo.get('상해입원일당(입원특약)',0)+_v1
        if company:
            # ★★★v237: 세부가입현황(상세내역) 계약별 CI 정보를 계약에 부착한다.
            #   롯데(2열) 별첨엔 '주계약' 라벨이 없어 `ci_jugye=[]`가 되므로,
            #   선지급률(50%/80%) 판정의 <b>2순위 근거</b>로 쓴다(지점장 지시 2026.07.25).
            _cs=None
            try:
                _ck=re.sub(r'\s','',str(company or ''))
                for _k,_v in (_SEBU_CI or {}).items():
                    _k2=re.sub(r'\s','',_k)
                    if _k2==_ck or _k2.startswith(_ck[:3]) or _ck.startswith(_k2[:3]): _cs=_v; break
            except Exception: pass
            contracts.append({'company':company,'ipwon':ipwon,'ci_extra':ci_extra,'product':product,'contract_date':contract_date,
                'expiry_date':expiry_date,'premium':premium,'pay_period':pay_period,
                'pay_count':pay_count,'renewal':renewal,'dambo':dambo,'ci_jugye':ci_jugye,'ci_sebu':_cs,'ci_lines':ci_lines})
    # ★★v100 단계약 사각지대: KB·메리츠 3열이 '계약 1건'이면 앵커가 1개뿐이라
    #   sinjeong_detect(>=2)를 못 넘겨 2열 파서로 가고 계약 0건이 된다(양예서형 단계약 고객).
    #   → 2열이 0건일 때만 3열을 재시도한다. 2열이 1건이라도 잡으면 손대지 않으므로 회귀 없음.
    if (not contracts) and _SJN >= 1:
        contracts = parse_sinjeong(lines)
        print(f'[SINJEONG-FALLBACK] 2열 0건 + 3열 앵커 {_SJN}개 → 3열 재시도 = {len(contracts)}건')
    # ★ 페이지 분할 중복 제거 (정본 체크리스트 ①②): 동일 계약키 병합
    merged = {}
    order = []
    for c in contracts:
        key = (re.sub(r'\\s','',c['company']), c['premium'], re.sub(r'\\s','',c['product'])[:12])   # ★v30p 날짜 OCR 깨짐 대비 병합키
        if key not in merged:
            merged[key] = c; order.append(key)
        else:
            m = merged[key]
            # 담보 병합: 같은 담보명은 큰 값 유지(중복가산 방지), 새 담보는 추가
            for k, v in c['dambo'].items():
                m['dambo'][k] = max(m['dambo'].get(k, 0), v)
            if not m.get('ci_jugye') and c.get('ci_jugye'): m['ci_jugye'] = c['ci_jugye']
            # 더 긴(덜 잘린) 상품명 채택
            if len(c['product']) > len(m['product']): m['product'] = c['product']
            # 회차/기간 비어있으면 채움
            if not m['pay_count'] and c['pay_count']: m['pay_count'] = c['pay_count']
            if not m['pay_period'] and c['pay_period']: m['pay_period'] = c['pay_period']
    deduped = [merged[k] for k in order]
    # ★★★제외 7종(v125, 지점장 확정 2026.07.21): 실손이 아닌데 <계약일 또는 만기일이 없는> 계약은
    #   보험기간(1년 여부)을 판정할 수 없다 → 엑셀·보장나무·보장진단서·보장설명서 전부 미포함.
    #   ★실손은 예외 — 롯데 리포트가 실손 계약의 계약일·보험료를 공란으로 주는 사례가 있다(v90·장문순 실측).
    #   ★담보를 봐야 실손인지 알 수 있으므로 파싱이 끝난 이 시점에서 판정한다.
    _kept = []
    for _c in deduped:
        if _no_period(_c.get('contract_date'), _c.get('expiry_date')) and \
           not _is_silson_like(_c.get('company'), _c.get('product'), _c.get('dambo')):
            print(f"[제외7·기간불명] {_c.get('company')} {str(_c.get('product'))[:28]} — 계약일/만기일 없음(실손 아님)")
            continue
        _kept.append(_c)
    deduped = _kept
    # ══════════════════════════════════════════════════════════════════════════
    # ★★★v47 심장 묶음담보 분해 (지침 §8.3.1 + 보험인포메이션 p16~19 회사별 정본표)
    #   "묶음 진단비는 보장 구성질환의 마스터 행에 동일 금액을 각각 기재한다."
    #   회사마다 '특정Ⅰ/Ⅱ'의 뜻이 다르다 → 라벨 말고 회사별 질병코드 기준으로 분해.
    #   ★2026.07.13 지점장 확정 3건:
    #     (1) KB 특정Ⅰ = 협심증+허혈성+빈맥+심부전 <b>+염증</b> (구 정본 '염증X' 폐기)
    #     (2) KB 심장판막질환에서 염증(심내막염 I33) 삭제 → 판막만
    #     (3) 현대 특정Ⅰ = <b>협심증</b>+빈맥+심부전 (구 빈맥+심부전)
    #   끝열은 행별 가로 SUM이라 세로 중복합산 없음. 원천 분해 → 4대 산출물 자동 연동.
    # ══════════════════════════════════════════════════════════════════════════
    # ══════════════════════════════════════════════════════════════════
    # ★★★심장 묶음담보 정본 (지점장 확정 2026.07.14 · 회사별 표대로)
    #   공통원칙: 묶음 진단비 = 구성질환 마스터 행에 동일 금액 각 100% 기재(절반 분할 금지).
    #   ★★'허혈성 진단비' 행 = 회사담보명 '허혈성심장질환진단비' <b>단독 담보 전용</b>(전 회사 공통).
    #      묶음(특정Ⅰ·Ⅱ·Ⅲ 등)은 허혈성 행에 절대 넣지 않는다. I20·I24·I25는 '협심증' 행으로 표현.
    #   ★빈맥(I47·I48)과 부정맥(I49)은 별개 코드.
    # ══════════════════════════════════════════════════════════════════
    _HB = {   # ★★★심장 묶음담보 회사별 정본표(지점장 최종본 2026.07.18 · 이전 표 폐기)
      #  ★'허혈성 진단비' 행 = 회사담보명 '허혈성심장질환진단비' 단독 담보 전용.
      #   묶음은 이 행에 절대 안 넣는다. I20·I24·I25는 '협심증' 행으로 표현.
      #  ★빈맥(I47·I48)과 부정맥(I49)은 별개 코드.
      'KB': [
        (lambda t: '특정1' in t and '심장' in t,   ['협심증','빈맥','심부전','주요심장염증']),
        (lambda t: '특정2' in t and '심장' in t,   ['급성심근경색']),
        (lambda t: '기타심장부정맥' in t or ('부정맥' in t and 'I49' in t.upper()), ['부정맥']),
        (lambda t: '심장판막' in t,                 ['심장판막']),
        (lambda t: '심근병증' in t,                 ['심근병증']),
      ],
      '한화': [
        (lambda t: ('기타부정맥제외' in t or 'I49제외' in t) and ('심혈관' in t or '특정1' in t),
                                                   ['협심증','빈맥','심부전']),
        (lambda t: '심혈관1' in t or '특정1' in t,  ['협심증','빈맥','부정맥','심부전']),
        (lambda t: '심혈관2' in t or '특정2' in t,  ['급성심근경색']),
        (lambda t: '심근병증' in t,                 ['심근병증']),
      ],
      'NH': [
        (lambda t: ('기타부정맥제외' in t or 'I49제외' in t) and ('심혈관' in t or '특정1' in t),
                                                   ['협심증','빈맥','심부전']),
        (lambda t: '심혈관' in t and '특정1' in t,  ['협심증','빈맥','부정맥','심부전']),
        (lambda t: '심혈관2' in t or '특정2' in t,  ['급성심근경색']),
        (lambda t: '주요심장염증' in t,             ['주요심장염증']),
        (lambda t: '심근병증' in t,                 ['심근병증']),
      ],
      'DB': [
        (lambda t: '특정1' in t and '심장' in t,   ['협심증','주요심장염증']),
        (lambda t: '특정2' in t and '심장' in t,   ['급성심근경색']),
        (lambda t: '특정3' in t and '심장' in t,   ['심장판막','빈맥','심부전']),
        (lambda t: '순환계3대' in t,               ['빈맥','부정맥','심부전']),
      ],
      '현대': [
        (lambda t: '특정허혈' in t,                 ['급성심근경색']),
        # ★'허혈성심장질환진단비' = 무조건 단독(허혈성 행). 분해 금지.
        (lambda t: '특정2대' in t,                  ['부정맥']),
        (lambda t: '특정1' in t and '심' in t,      ['협심증','빈맥','심부전']),
        (lambda t: '특정2' in t and '심' in t,      ['급성심근경색']),
        (lambda t: '주요심장염증' in t,             ['주요심장염증']),
      ],
      '흥국': [
        (lambda t: '특정심혈관' in t and '기타심장부정맥' in t and '제외' not in t, ['부정맥']),
        (lambda t: '특정심혈관' in t,               ['협심증','빈맥','심부전']),
        (lambda t: '심근병증' in t,                 ['심근병증']),
        (lambda t: '주요심장염증' in t,             ['주요심장염증']),
      ],
      '롯데': [
        (lambda t: '특정심장' in t and '2' in t,    ['협심증','주요심장염증']),
        (lambda t: '특정심장' in t and '1' in t,    ['급성심근경색']),
        (lambda t: '특정15대' in t,                 ['심장판막','심근병증','빈맥','심부전']),
        (lambda t: '기타부정맥' in t,               ['부정맥']),
      ],
      # ★삼성·메리츠 '허혈성심장질환진단비' = 무조건 단독(허혈성 행). 분해 없음.
      '삼성':   [],
      '메리츠': [],
    }
    def _hbkey(comp):
        c = re.sub(r'[\s（）()]', '', str(comp or ''))
        for k in ('KB','한화','농협','NH','DB','현대','흥국','롯데','삼성','메리츠'):
            if k in c: return 'NH' if k == '농협' else k
        return None
    for _c in deduped:
        _hk = _hbkey(_c.get('company'))
        if not _hk or _hk not in _HB: continue
        for _k in list(_c['dambo'].keys()):
            _t = re.sub(r'\s', '', str(_k))
            if any(x in _t for x in ('[확인]','수술','주요치료','산정특례','혈전')): continue
            _t2 = _t.replace('Ⅰ','1').replace('Ⅱ','2').replace('Ⅲ','3')
            _t2 = re.sub(r'[IiⅠ]49', 'I49', _t2)
            for _pred, _rows in _HB[_hk]:
                try:
                    if not _pred(_t2): continue
                except Exception:
                    continue
                if len(_rows) <= 1: break          # 단일행은 기존 resolve에 위임(변경 없음)
                _v = _c['dambo'].pop(_k)
                for _r in _rows:
                    _nk = f'{_r}[심장묶음]'
                    _c['dambo'][_nk] = _c['dambo'].get(_nk, 0) + _v
                print(f"[v47 심장묶음] {_hk} '{_k}' {_v} → {' + '.join(_rows)} (각각)")
                break
    # ══════════════════════════════════════════════════════════════════
    # ★★v51 심뇌혈관수술비 분해 (지점장 확정 2026.07.13 · 현대해상 수술비)
    #   "심뇌혈관수술비 3,000 = 심장수술비 3,000 + 뇌혈관수술비 3,000 (각각 기재)"
    #   묶음담보 공통원칙(§8.3.1)과 동일 — 절반 분할 아님, 두 행에 동일 금액 각 100%.
    #   기존 매핑은 '심뇌혈관' → 뇌혈관수술비 하나로만 넣어 심장수술비가 누락됐다.
    #   원천(dambo)에서 쪼개므로 4대 산출물(엑셀·보장나무·PPT·설명서) 자동 연동.
    # ══════════════════════════════════════════════════════════════════
    for _c in deduped:
        for _k in list(_c['dambo'].keys()):
            _kk = re.sub(r'\s', '', str(_k))
            if '[확인]' in _kk: continue
            # ★v217: '심뇌<u>5대</u>혈관'처럼 사이에 글자가 끼는 표기도 잡는다(구 '심뇌혈관' 연속 조건 완화).
            if '심뇌' not in _kk or '혈관' not in _kk or '수술' not in _kk: continue
            _v = _c['dambo'].pop(_k)
            for _r in ('심장수술비', '뇌혈관수술비'):
                _nk = f'{_r}[묶음]'   # ★태그에 '뇌혈관' 금지(resolve 오인 방지)
                _c['dambo'][_nk] = max(_c['dambo'].get(_nk, 0), _v)   # ★v217 대표(max) — 여러 줄 합산 금지
            print(f"[v217 심뇌혈관수술] {_c.get('company')} '{_k}' {_v} → 심장수술비 + 뇌혈관수술비 (각각 대표 {_v})")

    # ══════════════════════════════════════════════════════════════════
    # ★★v46 결합담보 분해 (지점장 확정 2026.07.13 / 지침 §8.3.1 묶음담보 공통원칙 적용)
    #   "묶음(결합) 담보는 보장 구성담보의 마스터 행에 동일 금액을 각각 기재한다."
    #   예) 롯데 DB '상해사망80%이상후유장해 18,000'
    #       → 상해사망 18,000  +  상해80%이상후유장해 18,000  (두 행 각각)
    #   끝열은 행별 가로 SUM이라 세로 중복합산 없음. 원천(dambo)에서 쪼개므로 4대 산출물 자동 연동.
    # ══════════════════════════════════════════════════════════════════
    for _c in deduped:
        for _k in list(_c['dambo'].keys()):
            _kk = re.sub(r'\s', '', str(_k))
            if '[확인]' in _kk: continue
            if not ('사망' in _kk and '후유장해' in _kk): continue      # 결합담보만
            _v = _c['dambo'][_k]
            _ax = '질병' if '질병' in _kk else '상해'                    # 축: 질병 / 상해
            # 등급: '80%이상' 또는 '고도' 명시 → 80% 행 / 명시 없으면 3% 행 (담보명 문자 그대로, 추측 금지)
            _hi = ('80%이상' in _kk) or ('고도' in _kk)
            _dead = f'{_ax}사망[결합]'
            _dis  = f'{_ax}{"80%이상" if _hi else ""}후유장해[결합]'
            _c['dambo'].pop(_k)
            _c['dambo'][_dead] = _c['dambo'].get(_dead, 0) + _v
            _c['dambo'][_dis]  = _c['dambo'].get(_dis, 0) + _v
            print(f"[v46 결합담보 분해] {_c.get('company','')} '{_k}' {_v} → {_dead} + {_dis} (각각)")
    # 한장보장표 회차 주입 (별첨에 없던 pay_count 보정)
    for c in deduped:
        if not c['pay_count']:
            pc = paycount_map.get((c['company'], c['contract_date'], c['expiry_date'])) \
                 or paycount_map.get((c['contract_date'], c['expiry_date']))
            if pc: c['pay_count'] = pc
    # 병합·회차 보정 반영하여 갱신 재판정 (정본 §7 규칙대로만)
    for c in deduped:
        # ★★★v207 (지점장 확정 2026.07.25, 영구지침): 3열(KB·메리츠)도 judge_renewal을 그대로 탄다.
        #   <b>납입기간 == 보장기간(가입~만기)이면 '갱신'</b>이다 — 운전자·실손도 예외 없다.
        #   구 v44 규칙('3열은 총회차가 없으니 ④ 적용 금지')은 <b>폐기</b>. 3열에도 납입기간(20년납)과
        #   보험기간(2026.03.27~2046.03.27)이 그대로 인쇄돼 있어 ④ 판정에 필요한 값이 다 있다.
        #   실측 오류(양*선 KB): 삼성 운전자 20년납/20년만기 → 비갱신(오류) · New내돈내삼 54년납/54년만기 → 비갱신(오류).
        c['renewal'] = judge_renewal(c['product'], c['expiry_date'], c['pay_count'], c['contract_date'], c['pay_period'], c.get('company',''))
        # ★ 담보 절반 이상이 '갱신형' 표기면 갱신 강제(상품명만 보던 판정 보강). 단 종신(9999)은 유지.
        if not c['expiry_date'].startswith('9999') and c['dambo']:
            _dk=list(c['dambo'].keys())
            _gc=sum(1 for k in _dk if '갱신' in k and '비갱신' not in k)
            if _dk and _gc>=len(_dk)*0.5: c['renewal']='갱신'
    # ★신버전 보충: 세부가입현황에서 뇌·심 담보 파싱해 별첨서 0인 항목만 보충(첫 계약에 귀속)
    # ★★v43 뇌혈관 유동 재배치 (지점장 2026.07.13 확정)
    #   [정본] AIA생명·라이나생명·AIG손보·우체국 = 별첨에 '뇌혈관'이라 적혀 있어도 그대로 믿지 말 것.
    #          반드시 세부가입현황 표를 100% 대조해서, 실제로 잡힌 행(뇌혈관진단비/뇌졸증/뇌출혈)으로 배치한다.
    #          엑셀·PPT·보장진단서 모두 이 결과를 따른다. 회사별 하드코딩(구 '라이나=뇌출혈') 폐기.
    # ★★★v181 (지점장 확정 2026.07.22): <b>상세 세부내역 체크는 1000% 필수</b>.
    #   대상을 4개사에서 <b>생명보험사 전체 + CI/GI/리빙케어/퍼펙트 상품 전체</b>로 확대한다.
    #   생보 CI/GI는 담보명이 '뇌혈관진단'이어도 실제는 뇌출혈인 경우가 많아 담보명만 믿으면 안 된다.
    _SEBU_FORCE = ('AIA', '라이나', 'AIG', '우체국')          # 뇌출혈 기본배치 대상(실측 확인된 4개사)
    _SEBU_CHECK = ('생명', '라이프', '공제', 'AIA', 'AIG', '라이나', '우체국',
                   '메트라이프', '처브', 'ABL', 'KDB')        # 반드시 세부내역 대조할 대상(확대)
    def _sebu_target(_co, _pd):
        _c1 = re.sub(r'[\s（）()]', '', str(_co or ''))
        _p1 = re.sub(r'[\s（）()]', '', str(_pd or ''))
        if any(f in _c1 for f in _SEBU_CHECK): return True
        # ★★★v235: 여기도 `'CI보험'` 연속매칭이라 `CI종신보험`을 놓쳤다 → `_isci_prod`와 정본 1개로 통일.
        return _isci_prod(_p1)
    try:
        _sb = parse_sebu(lines)
        _nh  = float(_sb.get('뇌혈관진단비', 0) or 0)
        _jol = float(_sb.get('뇌졸증진단비', 0) or 0) or float(_sb.get('뇌졸중진단비', 0) or 0)
        _chu = float(_sb.get('뇌출혈진단비', 0) or 0)
        for _c in deduped:
            _co = re.sub(r'[\s（）()]', '', str(_c.get('company', '')))
            _pd = str(_c.get('product', ''))
            if not _sebu_target(_co, _pd):
                continue                                   # ★v181 생보사 전체 + CI/GI/리빙케어/퍼펙트
            for _k in list(_c['dambo'].keys()):
                _kk = str(_k).replace(' ', '')
                if not (('뇌혈관' in _kk) and ('진단' in _kk)):            continue
                if any(x in _kk for x in ('수술','주요치료','산정특례','혈전','특정')): continue
                if any(x in _kk for x in ('Ⅰ','Ⅱ','Ⅲ','II','III')):        continue
                if not _sb:
                    # ★★★v180 (지점장 확정 2026.07.21): 세부가입현황이 <b>이미지/벡터라 텍스트가 아예 없는</b>
                    #   리포트가 있다(실측 장문순: '한장 보장 현황'·'세부 가입 현황' 글자 자체가 0건 추출).
                    #   이때 [확인]큐로만 두면 <b>뇌혈관진단비 행에 그대로 남아 오류로 나간다</b>.
                    #   → 대상 4개사(AIA·라이나·AIG·우체국)는 <b>기본값을 뇌출혈진단비로 배치</b>하고
                    #     확인 메모를 남긴다(지점장: "AIA생명은 뇌출혈인데 뇌혈관으로 나온다").
                    if not any(_f in _co for _f in _SEBU_FORCE):
                        # ★v181 4개사 외(그 외 생보·CI)는 임의 이동 금지 — 확인큐로만 강하게 띄운다.
                        _c['dambo']['[확인] 세부가입현황 미파싱 · 뇌혈관/뇌졸증/뇌출혈 축 대조 필수 ' + str(_k)] = 0
                        print(f"[v43 확인큐·필수대조] {_co} '{_k}' 세부가입현황 미파싱 → 상세내역 수기 대조")
                        continue
                    _v0 = _c['dambo'].pop(_k)
                    _nk0 = '뇌출혈진단비[세부가입정본]'
                    _c['dambo'][_nk0] = float(_c['dambo'].get(_nk0, 0) or 0) + float(_v0 or 0)
                    _c['dambo']['[확인] 세부가입현황 미파싱 → 뇌출혈로 배치함, 상세내역 대조 요망 ' + str(_k)] = 0
                    print(f"[v43 기본값·뇌출혈] {_co} '{_k}' 세부가입현황 미파싱 → 뇌출혈진단비 배치(+확인메모)")
                    continue
                if   _nh  > 0: _tgt = None                 # 세부가입현황이 뇌혈관진단비로 잡음 → 유지
                elif _jol > 0: _tgt = '뇌졸중'
                elif _chu > 0: _tgt = '뇌출혈'
                else:          _tgt = None
                if not _tgt:
                    continue
                _v = _c['dambo'].pop(_k)
                _nk = _tgt + '진단비[세부가입정본]'
                _c['dambo'][_nk] = float(_c['dambo'].get(_nk, 0) or 0) + float(_v or 0)
                print(f"[v43 뇌혈관 유동재배치] {_co} '{_k}' → {_nk} (세부가입현황 정본)")
    except Exception as _e:
        print(f"[v43 뇌혈관 재배치 스킵] {_e}")

    try:
        _sebu=parse_sebu(lines)
        if _sebu and deduped:
            _loc={}
            for _ci,c in enumerate(deduped):
                for dk in list(c['dambo'].keys()):
                    _std=_dedup_std(dk)
                    if _std:
                        _loc.setdefault(_std,[]).append((_ci,dk))
            for k,v in _sebu.items():
                if k=='__무시__': continue   # ★v30z 무시지정 담보 이중차단
                if k in _loc:
                    _positions=_loc[k]
                    _bsum=0.0
                    for _ci,dk in _positions:
                        try: _bsum+=float(str(deduped[_ci]['dambo'][dk]).replace(',',''))
                        except: pass
                    if abs(_bsum - v) > 0.5:
                        _fci,_fdk=_positions[0]
                        deduped[_fci]['dambo'][_fdk]=v
                        for _ci,dk in _positions[1:]:
                            deduped[_ci]['dambo'].pop(dk,None)
                else:
                    deduped[0]['dambo']['[세부보충]'+k]=v
    except Exception as _e:
        pass
    # ★★v129 지점장 확정 2026.07.21: 월보험료 합계 판정 근거는 보유계약 리스트의
    #   '납입횟수(납부기간/총납부횟수)'와 '잔여보험료(완납이면 0)'다.
    #   보유계약 리스트 줄을 스캔해 계약에 remain을 붙인다. 못 찾으면 None → 회차로 폴백.
    try:
        _rows = []
        _re_row = re.compile(
            r'^\s*\d{1,2}\s+(\S.*?)\s+(\d{1,3})\s*/\s*(\d{1,3})\s+([\d,]+)\s*원'
            r'\s+([\d,]+)\s*원\s+([\d,]+)\s*원\s*$')
        for _ln in txt.split('\n'):
            _m = _re_row.match(_ln)
            if not _m: continue
            try:
                _rows.append({
                    'head':   re.sub(r'\s', '', _m.group(1)),
                    'pc':     f'{_m.group(2)}/{_m.group(3)}',
                    'prem':   int(_m.group(4).replace(',', '')),
                    'remain': int(_m.group(6).replace(',', '')),
                })
            except Exception:
                pass
        if _rows:
            for _c in deduped:
                _c['remain'] = None
                _key = re.sub(r'\s', '', str(_c.get('company', '')) + str(_c.get('product', '')))
                for _r in _rows:
                    if _r['prem'] and _r['prem'] == (_c.get('premium') or 0):
                        _c['remain'] = _r['remain']
                        if not _c.get('pay_count'): _c['pay_count'] = _r['pc']
                        break
                else:
                    for _r in _rows:
                        if _key and _r['head'][:8] and _r['head'][:8] in _key:
                            _c['remain'] = _r['remain']
                            if not _c.get('pay_count'): _c['pay_count'] = _r['pc']
                            break
    except Exception:
        pass
    return {'client':client,'contracts':deduped}

# ★ DMAP — 마스터 엑셀 B열 기준 100% 일치
DMAP = {
    '5대골절진단':'5대골절진단비',   # ★v30c '골절진단(간편Ⅲ)' 부분일치 오탐 차단(선순위)
    # ★v29t §8.1: 동양류 '[N] 주계약_주계약' 2줄 = 일반사망+상해사망 1:1 (~2 접미사=두 번째 줄)
    '주계약_주계약~2':'상해사망','주계약_주계약':'일반사망',
    # 사망
    '상해사망':'상해사망','상해사망(갱신형) [보통약관]':'상해사망','일반상해사망':'상해사망',
    '기본계약(상해사망(간편가입Ⅲ))담보':'상해사망',
    '질병사망':'질병사망(80세)',
    # 후유장애
    '상해후유장해3%':'상해후유3%','상해후유80%':'상해후유80%',
    '질병후유장해3%':'질병후유3%','질병후유80%':'질병후유80%',
    # 암 — B열: 고액암/일반암/중대한 암/유사암(갑.기.경.제)/표적항암치료비/하이클래스(암)/중입자치료비/양성자치료/세기조절치료/다빈치로봇수술비/암수술/암일당/항암방사선약물
    '일반암진단비':'일반암','암진단Ⅱ(유사암제외)(간편가입Ⅲ)담보':'일반암',
    '고액암진단비':'고액암',
    '갑상선암.기타피부암.유사암진단비Ⅲ':'유사암(갑.기.경.제)','유사암진단비':'유사암(갑.기.경.제)',
    '유사암진단Ⅱ(양성뇌종양포함)(간편가입Ⅲ)담보':'유사암(갑.기.경.제)',
    '표적항암약물허가치료비':'표적항암치료비','표적항암약물허가치료(간편가입Ⅲ)(갱신형)담보':'표적항암치료비',
    '항암방사선.약물치료비':'항암방사선약물','항암방사선치료(간편가입Ⅲ)담보':'항암방사선약물',
    '항암약물치료(간편가입Ⅲ)담보':'항암방사선약물',
    '항암방사선(세기조절)치료(간편가입Ⅲ)(갱신형)담보':'세기조절치료','항암세기조절방사선치료비':'세기조절치료',
    '항암방사선(양성자)치료(간편가입Ⅲ)(갱신형)담보':'양성자치료','항암양성자방사선치료비':'양성자치료',
    '암수술(간편가입Ⅲ)담보':'암수술',
    '카티(CAR-T)항암약물허가치료비':'항암방사선약물',
    # 뇌혈관 — B열: 뇌혈관진단비/뇌졸증진단비/중대한 뇌졸증/뇌출혈진단비/외상성뇌출혈/산정특례뇌혈관/혈전용해치료비
    '뇌혈관질환진단비Ⅲ(건강맞춤형Ⅱ)(갱신형)':'뇌혈관진단비','뇌혈관질환진단비':'뇌혈관진단비',
    '뇌혈관질환진단(간편가입Ⅲ)담보':'뇌혈관진단비',
    '뇌졸중진단비':'뇌졸증진단비','뇌졸중진단(간편가입Ⅲ)담보':'뇌졸증진단비',
    '뇌졸중진단비(건강맞춤형Ⅱ)(갱신형)':'뇌졸증진단비',
    '뇌출혈진단':'뇌출혈진단비',
    '중증질환자(뇌혈관질환)산정특례대상진단비(연간1회한)(건강맞춤형Ⅱ)(갱신형)':'산정특례뇌혈관',
    '뇌혈관질환수술비Ⅲ(건강맞춤형Ⅱ)(갱신형)':'뇌혈관수술비','심뇌혈관질환수술(간편가입Ⅲ)담보':'뇌혈관수술비',
    '뇌경색증(I63)혈전용해치료비':'혈전용해치료비','혈전용해치료비Ⅱ(뇌졸중)(간편가입Ⅲ)담보':'혈전용해치료비',
    # 심장 — B열: 협심증/심부전/염증/부정맥/산정특례심장/2대 주요치료비/급성심근경색/중대한 급성심근/혈전용해치료비
    '허혈심장질환진단비Ⅲ(건강맞춤형Ⅱ)(갱신형)':'허혈성 진단비','허혈심장질환진단비':'허혈성 진단비',   # ★v29t §8.3 구규칙(=협심증) 폐기
    '허혈심장질환진단(간편가입Ⅲ)담보':'허혈성 진단비',
    '급성심근경색증진단':'급성심근경색','급성심근경색증진단(간편가입Ⅲ)담보':'급성심근경색',
    '중증질환자(심장질환)산정특례대상진단비(연간1회한)(건강맞춤형Ⅱ)(갱신형)':'산정특례심장',
    '허혈심장질환수술비Ⅲ(건강맞춤형Ⅱ)(갱신형)':'허혈성수술비','허혈심장질환수술비':'허혈성수술비',   # ★v29t §8.3 허혈수술→허혈성수술비 행
    '급성심근경색증(I21)혈전용해치료비':'혈전용해치료비',
    # 일당 — B열: 질병일당/질병수술일당/질병종합병원일당/상해일당/간병인/간호통합병동/1인실 종합병원/1인실 상급병원/질병중환자실/상해중환자실
    '간병인사용질병입원일당(1일이상)(요양병원)(간편가입)(갱신형)':'간병인',
    '간호간병통합서비스질병입원일당(1-180일)(간편가입)(갱신형)':'간호통합병동',
    '상급종합병원질병입원일당(상급병실(1인실),1일이상60일한도)(간편가입)(갱신형)':'1인실 상급병원',
    '종합병원질병입원일당(상급병실(1인실),1일이상30일한도)(간편가입)(갱신형)':'1인실 종합병원',
    # 수술비
    '질병수술(간편가입Ⅲ)담보':'질병수술비',
    '상해수술비(건강맞춤형Ⅱ)(갱신형)':'상해수술비','상해수술(간편가입Ⅲ)담보':'상해수술비',
    '골절수술(간편가입Ⅲ)담보':'골절수술비','화상수술(간편가입Ⅲ)담보':'화상수술비',
    '120대질병수술Ⅱ(간편가입Ⅲ)(질병수술3(24대질병))담보':'n대수술비',
    '5대기관질병수술(관혈/비관혈)(연간1회한)(간편가입Ⅲ)담보':'5대기관 수술비 관혈',
    '중대한특정상해수술(간편가입Ⅲ)담보':'중대한상해수술비',
    # 운전자 — B열: 대인/대물/합의금/6주미만/변호사/자부상 (처리지원금 판정은 resolve_kw에서 순서대로)
    '교통사고벌금(대물)':'대물','교통사고벌금(대인)':'대인',
    '변호사선임비용':'변호사','자동차사고 변호사선임비용':'변호사','변호사비':'변호사',
    '자동차부상위로금':'자부상','자동차부상보장':'자부상',
    '무보험차에 의한 상해':'일상배상책임',
    # 골절 — B열: 골절(치아파절포함)/골절(치아파절제외)/5대골절진단비
    '골절진단(간편가입Ⅲ)담보':'골절(치아파절포함)',  # 단독 골절진단=치아포함 행(치아제외 명시만 제외 행)
    # 응급실
    '응급실내원비(응급)':'응급실(응급)',
    # 화상 — B열: 진 단 비/중증화상진단비
    '화상진단비':'화상진단비','화상진단비(건강맞춤형Ⅱ)(갱신형)':'화상진단비',
    # 깁스 — B열: 반깁스/깁스진단비
    '깁스치료담보':'깁스진단비','깁스치료':'깁스진단비',
    # 실손 — B열: 입원/통원/약값
    '질병입원의료비':'입원','상해입원의료비':'입원',
    '질병외래의료비':'통원',
    '도수/체외충격파/증식치료':'도수치료','비급여주사제':'비급여주사','MRI검사의료비':'MRI',
    # 일배책
    '가족생활배상책임':'일상배상책임','일상생활배상책임':'일상배상책임',
    # 치아
    '치과치료(보존치료)':'크라운','치과치료(보철치료)':'임플란트',
    # 제외
    '보험료납입지원':None,
}

# -*- coding: utf-8 -*-
import re
# 마스터 82행 전부 키워드로 잡는 사전 엔진. (predicate, std, jong)
# 순서 = 구체 우선. 앞에서 잡히면 끝.
def _norm(s): return re.sub(r'\s+','',s)

def _rmn(s):
    """담보명 등급 로마숫자/숫자 판별 → 3/2/1/0. 괄호 속(건강맞춤형Ⅱ 등)은 제외."""
    import re as _re
    _raw0=str(s)
    # ★양예서 버그: Adobe가 로마숫자를 전각괄호（）·파이프|로 깨뜨림. 괄호 안 등급도 읽어야 함(뇌질환진단비（II） 등)
    # 상품수식어 괄호(건강맞춤형Ⅱ 등)는 제거하되, '진단비（I/II/III）'처럼 담보 등급 괄호는 살린다
    s2=_re.sub(r'[(（](?!\s*[I|ⅠⅡⅢV\d]{1,4}\s*[)）])[^)）]*[)）]','',_raw0)  # 등급 아닌 괄호만 제거
    s2=s2.replace('（','(').replace('）',')').replace('|','I')  # 전각→반각, 파이프→I
    if 'Ⅲ' in s2 or 'III' in s2 or '(III)' in s2 or '3종' in s2: return 3
    if 'Ⅱ' in s2 or 'II' in s2: return 2
    if 'Ⅰ' in s2 or '(I)' in s2: return 1
    m=_re.search(r'진단비?\s*([123])(?!\d)',s2)
    if m: return int(m.group(1))
    if _re.search(r'[가-힣]I(?![A-Za-zI])',s2): return 1
    return 0

def resolve_kw(raw):
    # ★★★★★v247 (KB 3열 실측): <b>담보명 괄호 안의 장기 나열 글자에 걸려 오분류</b>되는 것을 최상단에서 차단한다.
    #   실측 — `갱신형 5대질환(심장,뇌혈관,신부전,간,폐 질환)수술비(연간1회한)[관혈]` 500×2가
    #   괄호 속 <b>'뇌혈관'</b> 글자 때문에 <b>뇌혈관수술비</b>로 가서 1,000 → <b>2,000</b>이 됐다(KB 표기 1,000).
    #   이 담보의 신정원담보명은 '특정질병수술' = <b>별개 담보</b>이므로 [확인]큐가 정답이다.
    #   ★기존 교훈("태그에 '뇌혈관' 글자 금지 — has('뇌혈관')를 먼저 검사해 오인한다")과 같은 뿌리.
    try:
        _r0 = re.sub(r'\s', '', str(raw or ''))
        if re.search(r'\d+\s*대질환', _r0) and '수술' in _r0:
            return None, 0
    except Exception:
        pass
    if str(raw).startswith('[확인]'): return None, 0   # ★v98 확인큐 항목은 표준행 매핑 금지(중복합산 차단)
    # ★★★v146 (지점장 확정 2026.07.21): 흥국화재 10억통장(플래티넘 건강 리셋월렛II)은
    #   <b>엑셀·보장분석 PPT에 표기 금지</b>. 보장진단서 7p 카드에만 표기한다.
    #   ★사고: 담보명에 '중환자실'이 들어 있어 resolve_kw가 '질병중환자실' 행으로 잡아
    #   엑셀 61행에 100,000(=10억)을 박고 있었다(실측). 최우선으로 차단한다.
    if ('리셋월렛' in re.sub(r'\s','',str(raw))) or ('리셋월랫' in re.sub(r'\s','',str(raw))):
        return None, 0
    # ★★★v224 (2026.07.25 이정화 실측): <b>담보명 접힘 잔해는 어느 행에도 넣지 않는다</b>.
    #   3열 별첨에서 담보명이 위·아래 줄에 걸치면 파서가 <b>여러 담보명을 한 줄로 뭉치고
    #   엉뚱한 열의 금액을 붙인다</b>. 실측 잔해(엑셀 `_dambo_raw`):
    #   `외)(181일이상)(맞춤간편고지)중증질환자뇌혈관질환산정특례대상진단비Ⅱ(연간1회한)(맞춤
    #    간호 간병통합서비스사용질병입원일당(요양/정신/한방병원제` → 금액 <b>10</b>
    #   → 산정특례뇌혈관 행에 10(일당 금액)이 찍혔다. <b>정답은 1,000</b>.
    #   접힘 복원(파서 개편) 전까지는 <b>[확인]큐로 보내 오출고를 막는다</b>.
    _fz = re.sub(r'\s','',str(raw))
    if re.match(r'^[\)\]）]|^외\)|^고지\)|^편고지\)|^한도\)|^신형\)|^외\(', _fz):
        return None, 0        # 담보명이 닫는 괄호·꼬리로 시작 = 접힘 잔해
    _mix = sum(1 for _k in ('진단비','입원일당','수술비','치료비','사망보험금') if _k in _fz)
    if _mix >= 2 and len(_fz) > 45:
        return None, 0        # 담보 종결어 2종 이상 + 45자 초과 = 여러 담보 혼합
    """raw 담보명 -> (std표준명 or None, jong 0~5). API 불필요."""
    raw = re.sub(r"^\[세부보충\]","",str(raw))  # ★세부보충 접두 제거
    # ★Adobe OCR 깨짐: '상하!'·'상하 !'·'상하）' = '상해' (지점장 2026.07.05)
    raw = re.sub(r"상하\s*[!！]", "상해", str(raw))
    # ★v30x 담보명 경계깨짐 가드: 후유장해가 앞부분이면 뒤에 붙은 골절수술비 등은 오염 → 후유장해 우선
    _r0=str(raw)
    if ('후유장해' in _r0) and (_r0.find('후유장해') < 25):
        if '80%' in _r0 or '80 %' in _r0 or '80％' in _r0:
            if '상해' in _r0[:_r0.find('후유장해')+6] or '재해' in _r0[:_r0.find('후유장해')+6]: 
                pass  # 상해후유80%는 아래 정상 로직서 처리
    r = raw; n = _norm(raw)
    has = lambda *ks: all(_norm(k) in n for k in ks)
    no  = lambda *ks: not any(_norm(k) in n for k in ks)
    # ★v30z 혈전용해치료비 우선(급성심근·뇌졸중 흡수 방지): '혈전용해' 포함시 전용행
    if has('혈전용해') and has('치료'): return '혈전용해치료비',0
    # ★철심제거·핀제거·내고정물제거 = 골절수술비 아님(별개 처치) → [확인]
    if (has('철심') or has('핀제거') or has('내고정물')) and has('수술'): return None,0
    # 종번호
    jong = 0
    for i,k in enumerate(['1종','2종','3종','4종','5종'],1):
        if k in n or f'({i}종)' in r: jong = i; break

    # 비담보성(보험료 납입면제·일시납입지원) → 매핑 안 함(자부상 등 오매핑 차단)
    if has('납입') and (has('면제') or has('지원') or has('대상보장')): return None,0

    # ★ 상해의료비 = 별개 정액 담보 단독 행(실손 입원/통원/약값과 합치지 말 것) — 지점장 2026.06.28
    if has('상해의료비') and no('입원','통원','외래','실손','처방','약제','도수','비급여'): return '상해의료비',0
    # ★★v95 (지점장 지시 2026.07.19): 1세대 구실손의 <b>'상해 의료비(입원+통원)'</b>은
    #   입원·통원으로 쪼개 넣을 수 없는 <b>상해의료비 단독 담보</b>다. 실손 행(입원/통원)에
    #   섞으면 한장보장표와 어긋난다. → 마스터 99행 '상해의료비'로 직행.
    #   실측: DB 0604_TM '상해(일반상해, 전체상해를 의미) 의료비(입원+통원) 100' → 상해의료비 100
    if has('상해') and has('의료비') and ('입원+통원' in n.replace(' ','') or '입원/통원' in n.replace(' ','')):
        return '상해의료비',0
    if has('외래') and has('의료비') and no('주사','MRI','도수','체외','증식','비급여'): return '통원',0   # ★v29x '외래의료비'(통원 표기 없음)=통원. 상해·질병 각각 와도 대표 최댓값 1건이라 중복합산 없음

    # ── 실손/수술일당 먼저 (수술·일당 오분류 차단) ──
    if (has('실손') or has('입원형') or has('입원의료비')) and has('입원'): return '입원',0
    if has('도수') or has('체외충격파') or has('증식치료'): return '도수치료',0   # 비급여 도수/체외/증식
    if has('MRI'): return 'MRI',0
    if has('비급여') and has('주사'): return '비급여주사',0
    if has('통원') and (has('실손') or has('외래') or has('의료비')) and no('주사','MRI','도수','체외','증식','비급여'): return '통원',0
    if has('상해') and has('수술') and has('일당'): return '상해수술일당',0   # ★v29q-10 상해수술입원일당→상해수술일당(질병수술일당 오입력 차단)
    if has('수술') and has('일당'): return '질병수술일당',0
    # ★★★v215 (지점장 확정 2026.07.25, 영구): <b>생명보험사 '급부금' 담보명 4종 정본 매핑</b>.
    #   ①질병수술급부금 = <b>질병수술비</b>  ②재해수술급부금 = <b>상해수술비</b>
    #   ③질병입원급부금 = <b>질병입원일당(질병일당)</b>  ④재해입원(급부금) = <b>상해입원일당(상해일당)</b>
    #   근거: 생보 약관은 '수술비·일당' 대신 '급부금'으로 인쇄한다. 구 코드는 '급부금' 문자열을 몰라
    #   `_is_pure_s/_is_pure_q`(상해수술비·질병수술비로 시작) 검사에서 전부 탈락 → <b>4종 모두 [확인]큐로 사라졌다</b>(실측).
    #   ★수식어가 붙은 변형(암·뇌·특정·교통·종합병원 등)은 base 담보가 아니므로 여기서 제외 → 기존 로직/[확인]으로 흘린다.
    _gbex = ('암','뇌','심','허혈','간질환','신장','폐','위','골절','화상','특정','교통','종합','상급','중환자','요양','재활','통원','외래','간병','장해','후유')
    if has('급부금') or has('재해입원'):
        if has('수술') and no(*_gbex) and jong==0:
            if has('질병'): return '질병수술비',0
            if has('상해') or has('재해'): return '상해수술비',0
        if has('입원') and no('수술', *_gbex):
            if has('질병'): return '질병일당',0
            if has('상해') or has('재해'): return '상해일당',0

    # ── 수술비 ──
    if has('수술'):
        if has('상해') and jong: return '상해 종수술비(1-5종)', jong
        if has('질병') and jong: return '질병 종수술비(1-5종)', jong
        # ★v215 (지점장 확정 2026.07.25): <b>'중대상해수술비' = '중대한상해수술비'</b>(같은 담보, '한' 한 글자 차이).
        #   구 코드는 `has('중대한','상해')`라 '중대상해수술비'가 탈락 → [확인]큐로 빠졌다(실측).
        if has('중대') and has('상해'): return '중대한상해수술비',0
        # ★★v216 자가진단으로 발견한 버그: 공백을 지우면 <b>'수술비 관혈' → '수술비관혈'</b>이 되어
        #   그 안에 <b>'비관혈'이 우연히 만들어진다</b> → 관혈 담보가 통째로 <b>비관혈 행</b>으로 갔다.
        #   괄호가 있는 '5대기관수술비(관혈)'은 우연히 살아남고, 괄호 없는 표기만 틀리던 <b>조용한 오분류</b>.
        _gwan = n.replace('수술비관혈', '수술비|관혈')
        if has('5대기관') and ('비관혈' in _gwan): return '5대기관 수술비 비관혈',0
        if has('5대기관'): return '5대기관 수술비 관혈',0
        if re.search(r'(?<!\d)\d{2,3}\s*대', r): return 'n대수술비',0   # ★v30k 10~150대(10·20·116·120·123대 등)→n대수술비. 5대기관은 위에서 처리, 2대주요치료비는 진단이라 여기 안 옴
        if has('뇌혈관') or has('심뇌혈관'): return '뇌혈관수술비',0
        if has('허혈'): return '허혈성수술비',0
        if has('심장') or has('심질환'): return '심장수술비',0
        if has('5대골절'): return '5대골절수술비',0
        if has('골절') and no('후유','장해','진단','일당','입원'): return '골절수술비',0
        if has('화상'): return '화상수술비',0
        # ★v51(지점장 확정 2026.07.13): 현대해상 '레보아이로봇수술비' = 다빈치로봇수술비(마스터 26행).
        #   '로봇' 키워드로 이미 잡힌다 — 이 조건을 좁히면 레보아이가 조용히 누락되므로 건드리지 말 것.
        if has('다빈치') or has('로봇') or has('레보아이'): return '다빈치로봇수술비',0
        if has('암') and no('양성종양','유사암'): return '암수술',0   # ★v30 양성종양·유사암 수술 오탐 차단 → [확인]
        if jong: return '종수술비공통', jong   # ★v29q-12 상해/질병·부위 미표기 1-5종 수술(예 파워수술 1-5종)→상해·질병 양쪽 슬래시
        if has('상해') or has('재해'):   # ★v30h 재해수술비=상해수술비 동일 취급
            # §6 상해수술비 = 기본만. 병원규모·부위/특정·통원·자XXXX 접두변형은 합산 금지 → [확인]
            # ★XXXX상해수술비(질병/상해수술비 앞 어떤 접두든) = 별개 아님→[확인](지점장 2026.07.05)
            _core_s = re.sub(r'^[\(\[][^\)\]]*[\)\]]\s*', '', r)
            _core_s2 = _core_s.strip().replace(' ','')
            # ★★★★★v247 (2026.07.25 KB 3열 실측): <b>3열(KB·메리츠)은 회사담보명 앞에 '갱신형'이 붙는다</b>
            #   (`갱신형 상해수술비` · `갱신형 질병수술비`). `startswith('상해수술비')` 검사라 <b>전부 탈락</b>해
            #   상해수술비 200 · 질병수술비 30이 <b>통째로 [확인]큐로 사라졌다</b>(KB 전체보장현황 대조 실측).
            #   → 판정 전에 <b>'갱신형'·'비갱신형' 접두어를 제거</b>한다. 2열(롯데)엔 이 접두어가 없어 영향 없다.
            _core_s2 = re.sub(r'^(?:비)?갱신형', '', _core_s2)
            _core_s2 = re.sub(r'^재해상해', '상해', _core_s2)   # 재해상해=상해(중복 정리)
            _core_s2 = re.sub(r'^재해(입원)?수술', r'상해\1수술', _core_s2)  # ★v65 재해수술비·재해입원수술비=상해수술비(지점장 2026.07.15, '입원' 낀 변형도 포함)
            _is_pure_s = _core_s2.startswith('상해수술비') or _core_s2.startswith('상해입원수술비')
            if _is_pure_s and no('흉터','복원','외모','특정','척추','관절','하지','상급','종합병원','안면','머리','목','3대','신경','인대','흉부','연골','통원','외래','자궁','자녀','자가','교통'):
                return '상해수술비',0
            return None,0
        if has('질병'):
            # ★v29 §8.5 질병수술비 합산군 = '질병수술비'·'질병입원수술비' 만.
            #   ★XXXX질병수술비(어떤 접두든) = 별개→[확인](지점장 2026.07.05) → [확인]
            _core = re.sub(r'^[\(\[][^\)\]]*[\)\]]\s*', '', r)   # 접두 수식어 괄호 제거
            _core_strip = _core.strip().replace(' ','')
            _core_strip = re.sub(r'^(?:비)?갱신형', '', _core_strip)   # ★v247 3열 '갱신형' 접두어 제거(KB 실측)
            # 순수 질병수술비/질병입원수술비로 시작해야 함(자XXXX 등 한글 접두 배제)
            _is_pure_q = _core_strip.startswith('질병수술비') or _core_strip.startswith('질병입원수술비')
            _excl_ok = not any(k in _core for k in ('특정','부위','관절','척추','외모','흉터','복원','신경','인대','연골',
                          '상급','종합','대수술','수술일당','일당','Ⅱ','Ⅲ','ⅱ','ⅲ',
                          '2종','3종','4종','5종','부분','관혈','내시경','로봇','통원','외래','5대','자궁','자가','자녀'))
            if _is_pure_q and _excl_ok and jong==0:
                return '질병수술비',0
            return None,0
    if has('창상') or has('봉합'): return '창상봉합술',0

    # ★v35 중증질환자 산정특례대상보장 = 산정특례 전용행. 파서가 담보명 3줄분할·OCR깨짐으로 조각냄
    #   실측조각(이성준KB 2026.07.09): ①'중증질환자（뇌혈관）'단독 ②'산정특례대상보장…중•증걸 환■자■（삼장）'
    if (has('중증질환자') or has('중증환자')) and has('뇌혈관') and no('수술'): return '산정특례뇌혈관',0
    if has('삼장') and no('수술'): return '산정특례심장',0   # OCR '삼장'=심장, 중증질환자 산특 조각 전용
    if (has('중증질환자') or has('중증환자')) and (has('심장') or has('삼장')) and no('수술'): return '산정특례심장',0
    # ★★★v227 (지점장 지시 2026.07.25, 영구): <b>'일반암직접치료' = 암수술 — 세부가입현황을 정본으로 따른다</b>.
    #   근거(이정화 AIA 건강+ 실측·검산 일치): 세부가입현황 AIA 건강+ 암 칸 = <b>일반암 1,000 / 암수술 1,000
    #   / 유사암 100 / 고액항암 8,000</b>. 별첨엔 `일반암직접치료 1,000`이 <b>2줄</b>인데 구 코드는
    #   ①매핑이 없어 [확인]큐로 빠지고 ②`_add`가 두 줄을 <b>합산 2,000</b>으로 만들었다.
    #   → <b>암수술 행 + 계약 내 대표(max) 1건</b>으로 처리하면 1,000이 되어 세부가입현황과 일치한다.
    #   전수 검산: 암수술 = 한화 750 + AIA 1,000 = <b>1,750</b> = 한장보장표 값.
    #   ★'암직접치료<b>입원일당</b>'은 위 일당 블록에서 이미 암일당으로 갔다(여기 안 온다).
    if has('암') and has('직접치료') and no('일당','입원','통원','방사선','약물','표적','양성자','세기','중입자','유사암'):
        return '암수술',0
    # ── 암 치료비 ── (지점장 2026.07.09 최종확정: '암주요치료비' 명시 > 하이클래스 > 유사암무시)
    if has('유사암') and has('주요치료'): return '__무시__',0   # ①유사암 주요치료비=무시(엑셀·PPT·설명지 전부)
    if has('암주요치료비') and no('유사암'): return '암주요치료비',0   # ★②담보명에 '암주요치료비' 있으면 하이클래스보다 우선→암주요치료비행 (하이클래스 암주요치료비형)
    if has('하이클래스'): return '하이클래스(암)',0   # ③암주요치료비 없는 하이클래스(항암약물형 등)→하이클래스(암)행. 2건이면 합산
    if has('주요치료') and no('순환계','2대','뇌','허혈','심장','심근','유사암','하이클래스'):   # ③하이클래스 없는 '병원+암주요치료비'→암주요치료비행. ★심장 추가(심장/순환계 주요치료비=2대주요치료비로, v38d)
        return '암주요치료비',0
    if has('고액항암') or (has('고액') and has('항암') and has('치료')): return '__무시__',0   # ★v30z 고액항암치료비=표적+양성자+세기조절+카티 합계값 → 무시(구성 치료비는 아래에서 각각 개별 매핑)
    # ★★★v215 (지점장 확정 2026.07.25, 영구): <b>괄호 안의 실제 치료법이 정본이다</b>.
    #   ①표적항암방사선치료비<b>(항암세기조절방사선)</b> = <b>세기조절치료</b>
    #   ②표적항암방사선치료비<b>(항암양성자방사선)</b> = <b>양성자치료</b>
    #   구 코드는 `has('표적')`이 세기조절·양성자 판정(아래)보다 <b>먼저</b> 걸려 둘 다 표적항암치료비로 뭉갰다(실측).
    #   ★세기조절·양성자 판정을 표적 <b>앞</b>으로 올린다.
    if has('세기조절'): return '세기조절치료',0
    if has('양성자'): return '양성자치료',0
    if has('표적'): return '표적항암치료비',0
    # ★v30h 암주요치료비 = 암특정치료비/암주요치료비/암(특정유사암포함)진단후(종합병원/상급종합병원)특정치료(지원금/비)
    #   구간밴드(연간1천~1억) 다줄·부위별 = 대표 1개(max, §8.2). 뇌·심·순환계·비급여·재활·통원·검사는 제외(각 전용행).
    if has('암') and (has('주요치료') or has('특정치료') or (has('진단후') and has('치료') and (has('종합병원') or has('특정'))))         and no('순환계','2대','뇌','허혈','심장','심근','비급여','하이클래스','재활','통원','입원일당','MRI','PET','초음파','검사','수술'):
        return '암주요치료비',0
    # 뇌혈관·허혈성심장 특정치료비 = 2대주요치료비(뇌·심 두 칸). ★v65 '뇌심주요치료비특약' 누락 수정(지점장 2026.07.15)
    if (has('뇌혈관') or has('허혈') or has('심장') or has('순환계') or has('뇌심') or (has('뇌') and has('심'))) and (has('특정치료') or has('주요치료')) and no('암','수술'):
        return '2대 주요치료비',0
    if has('치료지원금') or (has('진단후') and has('치료')): return None,0   # ★v30a 잔여 진단후 치료지원금(암·뇌·심 아닌) = 진단비 아님 → [확인]
    if has('유사암') and has('주요치료'): return None,0   # ★v30 유사암 주요치료비 = 전용행 없음 → [확인]
    # ★★★★★v248 (지점장 확정 2026.07.26 "유사암이라고 기재된 것만 해라"):
    #   구 v30a 규칙 <b>`갑상선암 + 진단 → 유사암`은 폐기</b>한다. 담보명에 '유사암'(또는 '소액암')이
    #   명시되지 않으면 <b>[확인]큐</b>로 보내 신인이 수기 확인한다(추측 금지).
    #   실측 — 메리츠 `갱신형 갑상선암(초기제외)진단비` 1,000이 유사암에 산입돼 3,000이 됐으나
    #   KB 전체보장현황 유사암은 1,000(`갱신형 유사암진단비`뿐)이다.
    if has('갑상선암') and has('진단') and no('유사암','소액암','주요치료','수술','일당'): return None,0
    if has('통합') and has('전이암') and no('주요치료','수술','통원','일당'): return '통합전이암',0   # ★v30z 통합전이암=개별담보·대표금액 1개(§8.2, PPT·보장설명지 반드시 반영)
    if has('전이암') and no('통합'): return '__무시__',0   # ★v30z 전이암진단비 단독=무시(지점장 2026.07.05)
    if has('암') and has('주요치료') and no('순환계','2대','유사암'): return '암주요치료비',0   # ★v30 '암주요치료' 명시가 비급여 수식어보다 우선
    if has('하이클래스'): return '하이클래스(암)',0
    if (has('비급여') or has('하이클래스')) and has('주요치료'): return '하이클래스(암)',0   # 비급여 주요치료비(암 미명시)=하이클래스(암)
    if has('중입자'): return '중입자치료비',0
    if has('양성자'): return '양성자치료',0
    if has('세기조절'): return '세기조절치료',0
    if has('항암') and (has('방사선') or has('약물')) and no('표적','호르몬'):   # ★v65 표적·호르몬 제외(지점장 2026.07.15): 표적항암약물허가/특정항암호르몬약물이 여기 흡수돼 항암방사선약물에 표적값(2천) 오입력되던 것 차단
        # ★v30: 특정부위·특정암 한정 변형(예 남성생식기관련암 3,000)은 기본 항암방사선과 별개 → [확인]
        if has('생식기') or has('전립선') or has('음경') or has('고환') or has('유방') or has('자궁') or has('갑상선'): return None,0
        return '항암방사선약물',0
    if has('카티') or has('CAR-T') or has('CART'): return '항암방사선약물',0
    # ★v197 지점장 확정 2026.07.23: 'N대특정암진단비'(16대·11대·5대 등)·'특정암진단비' = 고액암 행
    #   근거: 암진단비(유사암제외)와 별개 담보 → 일반암 산입 금지(등식1 위반)
    if has('특정암') and has('진단') and no('주요치료','특정치료','수술','일당','입원','방사선','약물','표적','양성자','세기','중입자','통원','보험료'):
        return '고액암',0
    if has('고액암'): return '고액암',0
    # ★v29q-2 '암입원일당(…유사암들…)' = 암입원일당 키워드 우선 → 암일당 (유사암 판정보다 먼저)
    if ('암입원일당' in n) or (has('암') and has('입원일당')): return '암일당',0
    # ★v29q-1 '암진단비(…유사암들…)' = 암진단비 키워드 우선 → 일반암 (괄호 유사암 구성 무시)
    if has('소아암') and no('제외'): return None,0   # ★v30q 다발성소아암 등 = 일반암과 별개 담보 → [확인](합산 금지, 지점장 2026.07.03)
    if re.search(r'암\s*진단비\s*[(（]', r) and no('유사암제외'): return '일반암',0
    # 유사암 — 단 '유사암제외'(유사암을 뺀 일반 암진단)는 일반암
    # ★★★★★v248 (지점장 확정 2026.07.26): <b>"유사암이라고 기재된 것만 해라"</b>
    #   → 유사암 행에는 담보명에 <b>'유사암'(또는 '소액암')이 명시된 담보만</b> 넣는다.
    #   구 v230의 <b>동의어 자동산입</b>(갑상선·갑상샘·기타피부·경계성·제자리·상피내·양성뇌종양)은 <b>폐기</b>.
    #   ★실측 근거(양*선 KB 3열): 메리츠 `갱신형 갑상선암(초기제외)진단비` 1,000 ·
    #     `갱신형 갑상선암 및 기타피부암의 전이암(림프절 등 전이제외)진단비` 1,000이 유사암에 산입돼
    #     <b>3,000</b>이 됐으나 KB 전체보장현황 유사암은 <b>1,000</b>(= `갱신형 유사암진단비`뿐)이다.
    #   ★★<b>영향 고지</b>: 이 규칙으로 이정화 우체국 `갑상샘암치료보험금`·`상피내암치료보험금`은
    #     <b>[확인]큐</b>로 간다(구 v230에선 유사암에 산입돼 한장표 900과 일치했다). 추측 대신 신인 수기 확인.
    if any(k in n for k in [_norm(x) for x in ['유사암','소액암']]) and no('유사암제외','유사암 제외'):
        # ★★★v207 (지점장 확정 2026.07.25, 양*선 메리츠 실측): '유사암(갑.기.경.제)'는 <b>진단비 전용 행</b>이다.
        #   글자만 보고 넣던 탓에 <b>수술비·치료비·일당</b>까지 산입돼 유사암이 1,250(=100+1,000+150)으로 부풀었다.
        #   실측 오류 2건 — '갱신형 갑상선기능항진증치료비' 100(갑상선 <b>기능</b>항진증 = 암이 아니다) ·
        #                  '갱신형 유사암수술비' 150(수술비는 진단비 행이 아니다).
        #   → 진단이 아닌 담보는 <b>[확인] 큐로 보낸다</b>(임의로 암수술 행에 넣으면 암수술 200→350으로 또 틀린다).
        if has('수술') or has('치료비') or has('일당') or has('입원') or has('통원') or has('주요치료') or has('기능항진') or has('기능저하'):
            return None, 0
        return '유사암(갑.기.경.제)',0
    if has('중대한') and has('암'): return '중대한 암',0
    if has('암') and has('진단') and no('고액','소액','표적','방사선','약물','수술','일당','양성자','세기','중입자','전이','뇌','보험료'):
        return '일반암',0
    # ★v30l ○○암보험(진단 표기 없는 암 주계약, 예 신한생활비주는암보험·종합암보험) → 일반암 합산. 유사/고액/통합/전이/중대한/치료·수술·일당 계열은 위에서 이미 분기
    if has('암보험') and no('유사','고액','소액','통합','전이','중대한','주요치료','특정치료','하이클래스','수술','일당','방사선','약물','표적','양성자','세기','중입자','보험료'):
        return '일반암',0
    if has('암') and has('입원'): return '암일당',0

    # ── 뇌혈관 ──
    if has('외상성') and has('뇌출혈'): return '외상성뇌출혈',0
    if has('뇌출혈'): return '뇌출혈진단비',0
    if has('중대한') and has('뇌졸'): return '중대한 뇌졸증',0
    if has('뇌졸'): return '뇌졸증진단비',0
    if has('산정특례') and has('뇌'): return '산정특례뇌혈관',0
    # ★2026.07.12 지점장 확정: '특정뇌혈관' = 뇌졸증
    if has('특정') and has('뇌혈관') and has('진단') and no('수술','주요치료','산정특례','혈전'): return '뇌졸증진단비',0
    # ★v30o 고정(전 회사, 지점장 2026.07.03): 뇌혈관진단비Ⅰ→뇌혈관진단비 / 뇌혈관진단비Ⅱ→뇌졸증. Ⅲ은 뇌혈관진단비.
    if has('뇌혈관') and has('진단') and no('수술','주요치료','산정특례','혈전'):
        _n=_rmn(raw)
        if _n==2: return '뇌졸증진단비',0
        if _n in (1,3): return '뇌혈관진단비',0
    if has('뇌혈관') and has('진단'): return '뇌혈관진단비',0
    if has('혈전용해') and has('뇌'): return '혈전용해치료비',0

    # ── 심장 ──
    if has('주요치료') and (has('순환계') or has('2대') or has('뇌혈관') or has('심뇌') or has('허혈') or has('심장')): return '2대 주요치료비',0   # 뇌혈관+허혈성/심장 주요치료비=순환계=2대주요치료비
    if has('중대한') and (has('심근') or has('급성심근')): return '중대한 급성심근',0
    if has('심근병증') or has('심근증'): return '심근병증',0
    if has('판막'): return '심장판막',0
    if has('급성심근'): return '급성심근경색',0
    # ★2026.07.12 지점장 확정: '특정허혈성' = 협심증 (v28 '허혈심장질환진단비→허혈성 진단비' 규칙보다 우선)
    if has('특정') and (has('허혈성') or has('허혈심장')) and has('진단') and not has('수술'): return '협심증',0
    if has('허혈성진단') or ((has('허혈성') or has('허혈심장')) and has('진단') and not has('수술')): return '허혈성 진단비',0   # ★v29t 허혈심장질환진단 포함
    # ★v40b KB '심장질환(특정)' 진단비: 특정Ⅱ=급성심근경색 / 특정Ⅰ=허혈성 진단비 (지침 §8.3.1 KB).
    #   OCR이 로마숫자를 Ⅱ/II/2 등으로 흘려 매칭 실패하던 버그 수정 → 세 표기 모두 인식.
    if has('심장질환') and has('특정') and has('진단') and no('수술','주요치료'):
        _rr=str(raw)
        _is2 = ('Ⅱ' in _rr) or ('특정 II' in _rr) or ('특정II' in _rr) or ('특정 2' in _rr) or ('특정2' in _rr) or ('(특정 II)' in _rr) or ('（특정 II）' in _rr)
        _is1 = ('Ⅰ' in _rr) or ('특정 I' in _rr) or ('특정I' in _rr) or ('특정 1' in _rr) or ('특정1' in _rr) or ('(특정 I)' in _rr) or ('（특정 I）' in _rr)
        if _is2 and not _is1: return '급성심근경색',0
        if _is1 and not _is2: return '허혈성 진단비',0
        return '급성심근경색',0   # 구분 불가 시 급성심근경색(보수적)
    if has('일당') and (has('허혈') or has('협심') or has('심부전') or has('부정맥') or has('빈맥') or has('뇌혈관') or has('심뇌')): return None,0   # ★v30b 질환별 입원일당 ≠ 진단비 → [확인] (조성래 허혈일당 오합산 수리)
    if has('협심'): return '협심증',0
    if has('허혈'): return '허혈성 진단비',0   # ★v29t §8.3: 허혈 단독=허혈성 진단비(구 협심증행 폐기)
    if has('심부전'): return '심부전',0
    if has('심내막') or has('심근염') or has('심장막') or has('심장염증'): return '염증',0
    if has('빈맥'): return '빈맥',0   # ★지점장 2026.07.05: 빈맥(I47·48)=master 40행 정식 사용(v30z6 '무행' 폐기). 빈맥≠부정맥(I49)
    # ★★★v217 (지점장 지시 2026.07.25, 영구): <b>'심장부정맥 고주파·냉각절제술보장' = 부정맥 행이 아니다</b>.
    #   지점장 원문 = <b>"부정맥은 부정맥이라고 적혀있다"</b> → 마스터 '부정맥'(42행)은 <b>진단비 전용</b>이고,
    #   <b>고주파절제술·냉각절제술(전극도자절제술)은 치료 시술 담보</b>라 진단비 행에 넣으면 안 된다 → <b>[확인]큐</b>.
    #   실측 오류: `심장부정맥고주파·냉각절제술보장`이 아래 `has('부정맥')`에 걸려 부정맥 진단비로 산입됐다.
    if has('부정맥') and (has('고주파') or has('절제') or has('냉각') or has('시술') or has('도자')):
        return None,0
    if has('부정맥'): return '부정맥',0
    if has('산정특례') and has('심'): return '산정특례심장',0
    if has('2대') and has('주요'): return '2대 주요치료비',0
    if has('혈전용해'): return '혈전용해치료비',0

    # ── 사망 ──
    if has('CI') and has('사망'): return '중대한CI적용',0
    if has('교통') and has('사망'): return '교통상해사망',0
    if (has('상해') or has('재해')) and has('사망'): return '상해사망',0
    if has('질병') and has('사망'): return '질병사망(80세)',0
    if has('일반사망') or (has('사망') and no('상해','질병','교통','재해','CI','암','운전','입원','수술')): return '일반사망',0

    # ── 후유장애 ──
    if has('화재') and (has('후유') or has('장해')): return None,0   # ★v29q-9 화재상해후유(3~100%)≠상해후유3%, 담보행 미기재→[확인] 큐
    # ★★★v104 영구지침(지점장 확정 2026.07.20): '無고도장해보장' = 80% 상해후유장해.
    #    구 로직은 '상해/재해/교통' 글자가 없어 질병후유80%로 잘못 갔다.
    if has('고도장해') and no('질병'): return '상해후유80%',0
    if has('후유') or has('장해') or has('장애'):
        # ★★★v222 (지점장 지시 2026.07.25, 영구): <b>후유장해 20% · 50%는 결과값 미기재 — 더하지도 말 것</b>.
        #   마스터 후유장해 행은 <b>3% · 80% 넷뿐</b>이고 20%·50% 전용행은 없다.
        #   구 코드는 '80'이 없으면 전부 <b>3% 행에 합산</b>해서 금액을 부풀렸다.
        #   실측(이정화): `일반상해50%이상후유장해생활자금` 4,800 · `교통상해50%이상후유장해(운전자)` 100이
        #   상해후유3%에 합산됐다. → <b>[확인]큐로 보내고 어느 행에도 넣지 않는다</b>.
        if re.search(r'(?<![\d.])(20|50)\s*%', n) and not re.search(r'(?<![\d.])80\s*%', n):
            return None,0
        # ★★v92 (장혜경 실측): '질병후유장해(80%미만)Ⅱ'가 '80' 글자 때문에 80%행으로 잘못 갔다.
        #   → '80%미만'/'80% 미만'이면 3% 행. (한장보장표 질병3% 100과 일치)
        _u80 = ('80%미만' in n.replace(' ','')) or has('미만')
        sev = '3' if _u80 else ('80' if ('80' in n or has('고도')) else '3')
        body = '상해' if (has('상해') or has('재해') or has('교통')) else '질병'
        return f'{body}후유{sev}%',0

    # ── 일당/입원 ──
    # ★v62 간호간병통합서비스(지점장 2026.07.15): 담보명에 '간호'+'간병'+'통합' 동시 존재 →
    #   간호통합병동 행. '간병인사용…일당비(간호간병통합서비스)' 형태가 아래 has('간병인')에
    #   먼저 걸려 간병인으로 오분류·누락되던 것을 차단(간호통합병동을 간병인보다 우선 판정).
    if has('간호') and has('간병') and has('통합'):
        if has('181'): return None,0        # 1-180일 기준만(181일이상 제외, v41 유지)
        return '간호통합병동',0
    # ★★v216 (지점장 지시 2026.07.25): <b>'간호'가 빠진 '간병통합서비스' 형태도 간호통합병동</b>.
    #   실측 누락: `간병통합서비스입원일당`이 위 조건(간호+간병+통합)에서 탈락하고
    #   아래 `has('간병인')`에도 안 걸려(‘간병인’이 아니라 ‘간병’) <b>[확인]큐로 사라졌다</b>.
    #   ★'간병인'이 들어간 담보는 제외한다(그건 간병인/간병인지원일당 행).
    # ★★★v220 (2026.07.25 긴급 축소 — v216 과잉 완화가 사고를 냈다):
    #   v216에서 `간병+통합+서비스`만 있으면 간호통합병동으로 보냈더니
    #   <b>`간병비통합서비스보장`·`통합간병서비스진단비`·`간병통합서비스수술비` 같은 진단비·수술비까지
    #   간호통합병동 행으로 들어갔다</b>(실측). 진단비가 들어오면 <b>2,000만원</b>이 일당 칸에 찍힌다.
    #   → <b>'입원' 또는 '일당'이 있는 담보만</b> 허용하고, 진단·수술·사망·치료비·간병인은 제외한다.
    if has('간병') and has('통합') and has('서비스') and (has('입원') or has('일당')) \
       and no('간병인','진단','수술','사망','치료비','후유','장해','보험료'):
        if has('181'): return None,0
        return '간호통합병동',0
    if has('간병인') and has('요양병원') and no('제외'): return None,0  # ★요양병원 포함형 미기재(지점장)
    # ★★★v215 (지점장 확정 2026.07.25, 영구): <b>'간병인사용'이 들어가면 '지원'보다 우선해서 '간병인' 행</b>이다.
    #   실측 오분류: '간병인사용지원일당' · '간병인사용비용지원' · '간병인사용 질병입원일당(간병인 지원)'이
    #   아래 `has('간병인') and has('지원')`에 먼저 걸려 <b>간병인지원일당 행에 20만·7만이 찍혔다</b>(지점장 지적).
    #   '간병인사용'은 실제 간병인을 쓴 날에 주는 <b>간병인 담보</b>이고, '간병인지원일당'은 별개 전용 담보다.
    if has('간병인') and has('사용'): return '간병인',0
    if has('간병인') and has('지원'): return '간병인지원일당',0   # ★v29w (지점장 2026.07.02) 간병인지원일당 전용행
    if has('간병인'): return '간병인',0
    if has('간호') and (has('통합') or has('간병')):
        if has('181'): return None,0        # ★v41 간호통합병동 = 1-180일 기준만(181일이상 제외)
        return '간호통합병동',0
    if has('1인실') and has('상급'): return '1인실 상급병원',0
    if has('1인실') and has('종합'): return '1인실 종합병원',0
    if has('중환자') and has('상해'): return '상해중환자실',0
    if has('중환자') and has('질병'): return '질병중환자실',0
    if (has('질병') or has('수술')) and has('일당') and has('수술'): return '질병수술일당',0
    # ★★★v209 단독행 원칙(지점장 확정 2026.07.25, 영구): <b>'질병종합병원일당' 행은 그 이름으로 된 단독 담보 전용</b>이다.
    #   <b>병실 등급 담보(2~3인실 · 2인실 · 3인실 등)는 이 행에 넣지 않는다</b> → [확인] 큐.
    #   마스터에 있는 병실 행은 <b>1인실 종합병원 · 1인실 상급병원 둘뿐</b>이고, 2~3인실 전용행은 만든 적이 없다.
    #   실측 오류(양*선 삼성 New내돈내삼): 상급 2~3인실 5 + 종합 2~3인실 5가 이 행에 들어가 10으로 찍혔다.
    if has('인실'): return None,0
    # ★★★v223 (지점장 지시 2026.07.25, 영구): <b>'상급병원질병입원일당' = 마스터 전용행 신설(54행)</b>.
    #   지점장 원문 = "상급병원질병입원일당 → 그건 상급병원질병입원일당 신설해야 한다".
    #   질병일당·질병종합병원일당과 <b>별개 담보</b>다. ★'인실' 담보는 위에서 이미 걸러졌다(1인실 2행 전용).
    if has('상급') and has('질병') and (has('일당') or has('입원')) and no('수술','중환자','간병','간호','진단'):
        return '상급병원질병입원일당',0
    if has('질병') and has('종합') and has('일당'): return '질병종합병원일당',0
    # ★★★v223 = v212 재적용(지점장 지시 2026.07.25, 영구): <b>'N대질병입원일당'(2대·3대·5대…) ·
    #   '특정질병입원일당'은 질병입원일당이 아니다 → 기재 금지</b>([확인]큐).
    #   지점장 원문 = "2대질병입원일당(1일이상) → 그건 질병일당이 아니므로 기재금지".
    #   실측 오류: 메리츠 `갱신형 2대질병입원일당(1일이상)` 3만이 질병일당으로 산입됐다.
    if (has('일당') or has('입원')) and re.search(r'\d+\s*대', n) and no('인실'):
        return None,0
    # ★v30k 교통상해입원일당 ≠ 상해입원일당(합산 금지). 질환·부위·교통 접두 변형은 base 아님 → [확인]
    # ★병원규모(상급종합/종합) 명시 = 개별 전용행 / 일반 질병입원일당(밴드) = 합산 (지점장 2026.07.05)
    # ★v223 `_dilqual`에 <b>'상급'</b> 추가 — 상급병원 일당은 위 전용행으로 갔고, 남은 변형이
    #   질병일당·상해일당에 섞이는 것을 막는다(v212 정본 재적용).
    _dilqual = ('교통','암','뇌','심','허혈','간','신장','폐','위','골절','화상','특정','재해외','종합','요양','중환자','수술','상급')
    if (has('상해일당') or has('상해입원일당') or has('재해일당') or has('재해입원일당')) and no(*_dilqual): return '상해일당',0   # ★재해=상해 동일(정본)
    if (has('질병일당') or has('질병입원일당')) and no(*_dilqual): return '질병일당',0   # 순수 질병(입원)일당만 합산
    # ★v29v (지점장 2026.07.02): 밴드형 '입원비(1일이상/180일한도)' = 입원일당
    if has('입원비') and (has('1일') or has('180일')) and no('실손','의료비','수술'):
        return ('상해일당' if (has('상해') or has('재해')) else '질병일당'),0   # 재해=상해(§v29v)
    # ★v30c AIG류 밴드 미표기 base 입원비 = 입원일당 (변형·질환한정 [확인])
    if has('입원비') and no('실손','의료비','수술','중환자','상급','종합','중증','특정','암','뇌','허혈','심','간질환','감염'):
        if has('질병'): return '질병일당',0
        if has('상해') or has('재해'): return '상해일당',0

    # ── 운전자 (지침 §운전자 매핑) ──
    #  벌금(대인)→대인 / 벌금(대물)→대물 / 처리지원금(중상해포함)→합의금 / 처리지원금(6주미만)→6주미만
    #  변호사→변호사 / 자동차(사고)부상보장·부상위로→자부상
    if has('6주'): return '6주미만',0
    if has('처리지원금') or has('형사합의') or has('합의금'): return '합의금',0
    if has('벌금') and has('대물'): return '대물',0
    if has('벌금') and no('화재','과실','치사','업무'): return '대인',0   # ★v29q-7 벌금담보 단독=대인. 과실치사·업무과실 벌금 변형은 이중합산 차단→[확인]
    if has('대인') and no('대물'): return '대인',0
    if has('대물'): return '대물',0
    if has('변호사'): return '변호사',0
    if has('자부상') or (has('자동차') and (has('부상') or has('자부상'))):
        # ★자부상=자동차 한정. 급수밴드 있으면 14급 포함 밴드만(1~3/1~7 제외).
        _band=re.search(r'(\d+)\s*~\s*(\d+)\s*급', r)
        if _band:
            if int(_band.group(2))>=14: return '자부상',0
            return None,0
        # ★v65 단일 급수(예 '(1급)'·'(14급)') = 14급만 자부상, 그 외 제외(지점장 2026.07.15).
        #   이전엔 밴드 없으면 무조건 자부상이라 급별 여러 줄(1~14급)이 전부 자부상→합산 수천원 오류였다.
        _single=re.search(r'(?<![\d~])(\d+)\s*급', r)
        if _single:
            return ('자부상',0) if int(_single.group(1))>=14 else (None,0)
        return '자부상',0        # 급 표기 자체가 없는 순수 자동차부상위로금 → 자부상

    # ── 골절/응급/독감/화상/깁스 ──
    if has('5대골절') and has('진단'): return '5대골절진단비',0
    # §골절: '치아/파절 제외' 명시된 것만 제외 행. 단독 골절진단비·치아포함은 포함 행.
    #   ★삼성형 '치아파절 (깨짐, 부러짐) 제외'처럼 중간에 딴말 껴도 잡도록 (치아 or 파절) + 제외 동반 판정
    if has('골절') and has('제외') and (has('치아') or has('파절')): return '골절(치아파절제외)',0
    if has('골절') and (has('치아포함') or has('파절포함') or (has('제외') is False and (has('치아') or has('파절')))): return '골절(치아파절포함)',0
    if has('골절') and has('진단'): return '골절(치아파절포함)',0
    if _norm(raw)=='골절' or has('골절') and no('수술','일당','입원','깁스','부목'): return '골절(치아파절포함)',0
    if (has('응급실') or (has('응급') and has('내원'))) and no('비응급'): return '응급실(응급)',0   # ★v29q-11 응급 단독, 비응급 합산 차단→[확인]
    if has('독감') or has('인플루엔자'): return '독감',0
    if has('화상') and (has('중증') or has('심재성') or has('중대한') or has('부식')): return '중증화상진단비',0
    if has('화상') and has('진단'): return '화상진단비',0
    if has('부목') or has('반깁스'): return '반깁스',0   # ★v29q-5 골절부목치료비=반깁스
    if has('깁스'): return '깁스진단비',0

    # ── 실손 ──
    if (has('실손') or has('입원의료비') or has('상해입원형') or has('질병입원형')) and has('입원'): return '입원',0
    if has('통원') and (has('실손') or has('외래') or has('의료비')) and no('주사','MRI','도수','체외','증식','비급여'): return '통원',0
    if has('처방조제') or has('약제비') or (has('약') and has('실손')): return '약값',0
    if has('일상생활') and has('배상') or has('일배책') or has('일상배상'): return '일상배상책임',0
    return None, 0

def resolve2(raw):
    """raw -> (std, jong). DMAP 정확매칭 우선, 없으면 키워드 사전엔진(resolve_kw)."""
    if raw in DMAP:
        v = DMAP[raw]
        return (v, get_종번호(raw))
    for k, v in DMAP.items():
        if k and k in raw and v: return (v, get_종번호(raw))
    # ★v29t (지점장 2026.07.02): 가족동승 부상치료비 = 자부상 아님 → [확인]
    if '가족동승' in raw: return (None, 0)
    # ★v29t: 방사선항암(소액암) 변형 = 합산 금지 → [확인] (수기 정본은 기본 방사선 100만 기재)
    if '방사선' in raw and '소액암' in raw and '제외' not in raw: return (None, 0)
    # ★v29t 부정어 처리: '(소액암제외)'·'(유사암제외)' 등 제외 문구를 지우고 키워드 매칭
    #   (예 '암치료자금_암(소액암제외)진단비특약' → 소액암 오탐으로 유사암행 오매핑되던 버그 차단)
    # ★v44 버그수정: '(치아파절제외)'까지 통째로 지워 골절 제외행→포함행으로 오배치되던 문제(정본 §8.7 위반).
    #    치아·파절이 들어간 제외 괄호는 보존한다. (유사암제외·소액암제외 등은 종전대로 제거)
    raw_kw = re.sub(r'[\(\[](?![^\)\]]*(?:치아|파절))[^\)\]]*제외[\)\]]', '', raw)
    _r = resolve_kw(raw_kw)
    if _r[0] is None and '재해' in raw_kw:
        _r = resolve_kw(raw_kw.replace('재해','상해'))   # ★v29v (지점장 2026.07.02): 재해=상해 동일 적용
    return _r

def resolve(raw):
    return resolve2(raw)[0]

def _dedup_std(raw):
    """★세부보충 dedup 전용 — build_excel의 뇌질환/심장질환Ⅰ·Ⅱ 매핑과 동일 해석.
    (resolve_kw엔 이 매핑이 없어 세부보충이 중복제거 실패→배증하던 양예서 버그 차단)"""
    _rn = re.sub(r'\s','',str(raw))
    if '진단' in _rn and '수술' not in _rn and '주요치료' not in _rn:
        if ('심장질환진단' in _rn) and ('허혈' not in _rn) and ('급성심근' not in _rn):
            _mn=_rmn(_rn)
            if _mn==2: return '급성심근경색'
            if _mn==1: return '허혈성 진단비'
        if '뇌질환진단' in _rn:
            _mn=_rmn(_rn)
            if _mn==2: return '뇌졸증진단비'
            if _mn==1: return '뇌혈관진단비'
    return resolve_kw(raw)[0]

NOFILL = PatternFill(fill_type=None)

# ★ SUM 수식 캐시값 채우기 — openpyxl은 수식만 저장(캐시 없음)→모바일/미리보기 공란.
#   저장 후 LibreOffice 재계산으로 값 주입(수식은 유지=법칙22).
def recalc_xlsx(path):
    import subprocess, shutil, tempfile
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice: return False
    try:
        outd = tempfile.mkdtemp()
        subprocess.run([soffice,'--headless','--norestore','--convert-to','xlsx','--outdir',outd,path],
                       timeout=90, capture_output=True)
        out = os.path.join(outd, os.path.splitext(os.path.basename(path))[0]+'.xlsx')
        if os.path.exists(out):
            shutil.copyfile(out, path); shutil.rmtree(outd, ignore_errors=True)
            return True
        shutil.rmtree(outd, ignore_errors=True)
        return False
    except Exception:
        return False

# ★v29u: LibreOffice 없는 환경(Railway)용 캐시 주입 — 합계 수식은 유지(§5)하고,
#   계산값을 파이썬으로 구해 시트 XML <v>에 직접 기록 → 폰·미리보기·보장설명지 모두 값 표시.
def _no_fullcalc(wb):
    """★v51(2026.07.13): 산출 엑셀에서 fullCalcOnLoad 플래그 제거.
    master.xlsx가 이 플래그를 물려주면 Excel이 파일을 열 때마다 전 수식을 강제 재계산한다
    (편집모드 진입 시 폰 Excel이 1분 이상 로딩). 끝열 =SUM 캐시값은 inject_sum_cache가 이미
    채워두므로 강제 재계산은 불필요한 부하일 뿐이다. 수식(§5)은 그대로 유지한다."""
    try:
        wb.calculation.fullCalcOnLoad = False
    except Exception:
        pass

def inject_sum_cache(path):
    import zipfile, shutil, tempfile
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb['보장분석']; last = ws.max_column
        vals = {}
        for r in range(2, ws.max_row+1):
            f = ws.cell(r,last).value
            if not (isinstance(f,str) and f.startswith('=')): continue
            nums = [ws.cell(r,c).value for c in range(3,last) if isinstance(ws.cell(r,c).value,(int,float))]
            s = sum(nums)
            if f.startswith('=MIN('):
                m = re.search(r',\s*(\d+)\s*\)\s*$', f); v = min(s, int(m.group(1))) if m else s
            elif f.startswith('=IF(COUNT'):
                v = max(nums) if nums else 0
            elif f.startswith('=IF(SUM'):
                v = 7 if s>0 else 0
            else:
                v = s
            vals[ws.cell(r,last).coordinate] = v
        if not vals: return False
        zin = zipfile.ZipFile(path,'r')
        # 보장분석 시트 XML 경로 확인 (workbook.xml 순서 + rels)
        wbxml = zin.read('xl/workbook.xml').decode('utf-8')
        rels  = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        m = re.search(r'<sheet[^>]*name="보장분석"[^>]*r:id="(rId\d+)"', wbxml) or re.search(r'<sheet[^>]*r:id="(rId\d+)"[^>]*name="보장분석"', wbxml)
        rid = m.group(1) if m else 'rId1'
        m2 = re.search(r'Id="'+rid+r'"[^>]*Target="([^"]+)"', rels)
        tgt = 'xl/'+m2.group(1).lstrip('/') if m2 else 'xl/worksheets/sheet1.xml'
        sx = zin.read(tgt).decode('utf-8')
        for coord, v in vals.items():
            vv = ('%d' % v) if float(v).is_integer() else repr(float(v))
            sx = re.sub(r'(<c r="'+coord+r'"[^>]*>)(<f[^>]*>[^<]*</f>)(?:<v>[^<]*</v>)?(</c>)', r'\1\2<v>'+vv+r'</v>\3', sx, count=1)
        tmp = path+'.tmp'
        zout = zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED)
        for it in zin.infolist():
            zout.writestr(it, sx.encode('utf-8') if it.filename==tgt else zin.read(it.filename))
        zout.close(); zin.close()
        shutil.move(tmp, path)
        return True
    except Exception as _e:
        print(f'[INJECT_CACHE_ERR] {_e}')
        return False

# ★ LLM 매핑 엔진 — 마스터 표준 담보명에 의미기반 매핑 (앱 자동화 핵심)
def load_std_dambo(ws):
    out=[]
    for r in range(6, ws.max_row+1):
        v=ws.cell(r,2).value
        if v and str(v).strip(): out.append(str(v).strip())
    return out

def llm_resolve(raw_names, std_list):
    """raw 담보명 -> {raw: {'std': 표준명 or None, 'jong': 0~5}}  (jong>0이면 종수술비)"""
    raw_names=[r for r in raw_names if r]
    if not raw_names: return {}
    key=os.environ.get('ANTHROPIC_API_KEY','')
    if not key:
        return {r:{'std':None,'jong':0,'note':''} for r in raw_names}  # 키 없음 -> 전부 [확인]
    rules=("한국 보험 보장분석. 아래 담보명들을 표준목록의 표준명에 의미기반 매핑.\n"
        "규칙:\n"
        "- 표준목록에 있는 것만 선택. 의미가 명확히 일치할 때만. 애매/해당없음=null.\n"
        "- 중입자방사선=중입자치료비, 양성자=양성자치료, 세기조절=세기조절치료, 표적항암약물=표적항암치료비\n"
        "- '상해 1~5종/1-5종 수술비'(_N종·(N종)·N종)=상해 종수술비(1-5종), jong=종번호. 질병도 동일=질병 종수술비(1-5종)\n"
        "- 허혈심장/협심증=협심증, 급성심근경색=급성심근경색, 부정맥=부정맥, 심장질환수술=심장수술비\n"
        "- 유사암(갑상선/기타피부/제자리/경계성)=유사암(갑.기.경.제). 특수치료 아닌 일반 암진단=일반암\n"
        "- '뇌혈관진단'(로마숫자 없음) 담보는 세부가입현황 표를 정본으로 재배치한다(AIA·라이나·AIG·우체국 등). 임의 하드코딩 금지.\n"
        "- 하이클래스=하이클래스(암), 표적항암약물허가=표적항암치료비(여러 건이면 가장 큰 1건만)\n"
        "- 유사암 진단금액은 가입연도 2020년 이하면 일반암의 1/10, 2021년 이상이면 1/5로 환산 기재\n"
        "- 뇌혈관수술=뇌혈관수술비, 항암방사선/약물치료비=항암방사선약물, 암수술=암수술\n"
        "- 화상 진단비='화상진단비'(화상 구분 행), 중대한화상·부식=중증화상진단비\n"
        "- 생명보험 종신 주계약/기본계약(사망보장)=일반사망\n"
        "- 운전자: 벌금(대인)=대인, 벌금(대물)=대물, 교통사고처리지원금(중상해)=합의금, 처리지원금(6주미만)=6주미만, 변호사비=변호사, 자동차(사고)부상보장=자부상\n"
        "- 표준목록에 자리 없는 담보(예 크론병·다발경화증·장기이식 등)=null (행 추가 금지)\n"
        "- 심장담보 질병코드 분류: 협심증(I20)=협심증, 급성심근경색(I21~22)=급성심근경색, "
        "부정맥(I47~49)=부정맥, 심부전(I50)=심부전, 심내막·심근·심장막염=염증. "
        "단 '특정Ⅰ/Ⅱ·특정심장' 등 상품별 정의는 약관마다 달라 본표 단정 금지→null.\n"
        "- note 규칙(엄수): 약관 없이 보장범위를 추측·일반론으로 단정하지 말 것. "
        "담보명에서 명백한 사실만 적고, 상품별 정의(특정Ⅰ/Ⅱ 등)나 불확실한 건 '약관 확인 필요'로만 기재. "
        "의학적 추정(예 '통상 허혈성'·'대개 ~포함') 금지.\n")
    prompt=rules+f"\n표준목록: {std_list}\n\n담보명: {raw_names}\n\nJSON만 출력(설명 금지): {{\"담보명\":{{\"std\":\"표준명 또는 null\",\"jong\":0,\"note\":\"보장범위 요약\"}}}}"
    try:
        r=httpx.post('https://api.anthropic.com/v1/messages',
            headers={'x-api-key':key,'anthropic-version':'2023-06-01','content-type':'application/json'},
            json={'model':'claude-haiku-4-5-20251001','max_tokens':8000,
                  'messages':[{'role':'user','content':prompt}]}, timeout=90)
        print(f'[LLM_RESOLVE] status={r.status_code} raw={len(raw_names)}')
        if r.status_code != 200:
            print(f'[LLM_RESOLVE_HTTP] {r.text[:400]}')
            return {x:{'std':None,'jong':0,'note':''} for x in raw_names}
        txt=''.join(b.get('text','') for b in r.json().get('content',[]) if b.get('type')=='text')
        txt=txt.strip().replace('```json','').replace('```','').strip()
        out=json.loads(txt)
        print(f'[LLM_RESOLVE] mapped={len(out)}')
        return {k:{'std':(v.get('std') if isinstance(v,dict) else None),
                   'jong':(v.get('jong',0) if isinstance(v,dict) else 0),
                   'note':(v.get('note','') if isinstance(v,dict) else '')} for k,v in out.items()}
    except Exception as e:
        print(f'[LLM_RESOLVE_ERR] {e}')
        return {x:{'std':None,'jong':0,'note':''} for x in raw_names}

def _fix_silson(contracts):
    """★★★★★실손 정본 §8.8 — <b>실손은 [별첨] 보험서비스(상품)별 보장현황이 답이다</b>(지점장 확정 2026.07.25, 영구).
    ★<b>소스 우선순위는 실손에도 그대로</b>: ①<b>[별첨] 상품별 보장현황 = 정본</b> ②세부가입현황 ③한장보장표(검산).
      한장보장표·세부가입현황이 별첨과 다르면 <b>별첨을 따른다</b>. 실측 확정 사례 —
      DB생명 CI종신(3세대) 별첨 `외래의료비 <b>25</b>` vs 한장표·세부내역 30 → <b>25가 정답</b>
      (3세대 실손 외래 한도가 25만원). 구 [확인] 대기 항목은 이로써 <b>확정·해소</b>.
    ★<b>입원·통원 = 별첨 명시값 최우선</b>(1세대 1억·5천·3천 다 유지, 절대 뭉개지 말 것).
      명시값이 없을 때만 세대·회사유형 기본값을 넣는다: 1세대 통원10 / 손보 통원25·약5 / 생보 통원20·약10 / 4·5세대 통원20.
    ★<b>약값은 예외 2건</b> — 1세대와 4·5세대는 외래+약제가 <b>통원 한 한도로 통합</b>이라
      약값 담보 자체가 없다 → <b>별첨에 값이 있어도 0으로 삭제</b>(세대 규칙이 별첨보다 우선하는 유일한 항목)."""
    for c in contracts:
        d=c.get('dambo',{})
        co=str(c.get('company','') or '')
        prod=str(c.get('product','') or '')
        cd=str(c.get('contract_date','') or '')
        is_saengbo = ('생명' in co or '라이프' in co) and ('화재' not in co and '손해' not in co and '손보' not in co and '해상' not in co)
        # 원본 담보키 → 표준(입원/통원/약값) 위치 찾기
        _kmap={}  # 표준명 → 원본키
        for _rk in list(d.keys()):
            _std,_=resolve_kw(_rk)
            if _std in ('입원','통원','약값') and _std not in _kmap:
                _kmap[_std]=_rk
        if '입원' not in _kmap: continue  # 실손 없음
        def _get(std):
            k=_kmap.get(std)
            if not k: return None
            try: return float(str(d[k]).replace(',',''))
            except: return None
        def _set(std,val):
            k=_kmap.get(std)
            if k: d[k]=val
            else: d[std]=val  # 없던 항목(약값 등)은 표준키로 신설
        ipw=_get('입원')
        if ipw is None: continue
        # ★입원 = 보장표 명시값 최우선(1세대 1억·5천·3천 다 있음, 절대 뭉개지 말 것).
        #   명시값이 있으면 그대로 유지. 통원·약값만 규칙 보정.
        _tw=_get('통원'); _yk=_get('약값')
        # ★4·5세대(2021.07~) = 통원+약 합쳐서 20 (약값 별도 없음)
        _ym=None
        _mm=re.search(r'(\d{4})\.(\d{2})',cd)
        if _mm: _ym=int(_mm.group(1))*100+int(_mm.group(2))
        _gen=silson_gen(cd, ipw, prod, _has_nonpay3(d))   # ★v250 3대비급여 하한(d=이 계약 dambo)
        if _gen in ('4세대','5세대') or (_ym and _ym>=202107):
            # ★★★★★v246 영구지침(지점장 확정 2026.07.25): <b>실손은 [별첨] 상품별 보장현황이 답이다</b>.
            #   구 코드는 `_tw<=20`이 아니면 <b>20으로 덮어썼다</b> → 별첨 명시값이 20 초과면 잘렸다.
            #   → <b>별첨 명시값이 있으면 그대로 쓰고</b>, 없을 때만 기본값 20을 넣는다.
            _set('통원', _tw if _tw else 20)
            _set('약값', 0)     # ★4·5세대는 통원+약제 통합 → 약값 담보 자체가 없다(별첨에 있어도 삭제)
            continue
        # ★★★v215 (지점장 확정 2026.07.25, 영구): <b>실손 1세대 = 입원 3,000 / 통원 10 / 약값 0</b>.
        #   1세대는 외래+약제가 <b>통원 한 한도로 통합</b>이라 약값 담보 자체가 없다.
        #   ★버그: `silson_gen`은 입원한도 3,000이면 <b>'1세대(구형)'</b>을 반환하는데 구 v41 코드는
        #   `== '1세대'`로 비교해 <b>항상 False</b>였다 → 아래 손보/생보 기본값으로 흘러
        #   <b>약값 5(손보)·10(생보)이 신설</b>됐고, 그 값이 그대로 엑셀→보장진단서 PPT까지 나갔다
        #   (지점장 지적 "간병인·실손 보장진단서 오류"). → `.startswith('1세대')`로 수정.
        if str(_gen).startswith('1세대'):
            if not _tw: _set('통원', 10)      # ★v215 1세대 통원 기본 10(명시값 있으면 그대로)
            _set('약값', 0)
            continue
        # 1~3세대: 통원·약값 명시값 우선, 없으면 회사유형 기본값.
        if is_saengbo:
            if not _tw: _set('통원',20)
            if not _yk: _set('약값',10)
        else:
            if not _tw: _set('통원',25)
            if not _yk: _set('약값',5)
    return contracts


def build_excel(data, out):
    _ci_diag = []   # ★v238 CI 진단표(확인사항 상시 출력) — 함수 최상단에서 확실히 초기화
    wb = openpyxl.load_workbook(TPL_XL)
    ws = wb['보장분석']
    client = data['client']; contracts = _fix_silson(data['contracts'])

    # 담보명 -> 행번호 맵 (A/B열 유지)
    nm2r = {}
    nm2r_norm = {}   # 공백무시 보조키 ('진 단 비' 등 라벨 변형 흡수)
    nm2r_multi = {}  # 동일 담보명이 여러 행(2대 주요치료비=뇌혈관+심장) → 모두 기재용
    for r in range(6, ws.max_row+1):
        v = ws.cell(r,2).value
        if v:
            k = str(v).strip()
            nm2r[k] = r
            nm2r_norm[re.sub(r'\s','',k)] = r
            nm2r_multi.setdefault(k, []).append(r)

    # ★ 데이터영역(C열~) 전체 초기화 — 옛 7계약 헤더·합계·SUM수식·슬래시골격 제거
    MAXC = 60  # 최대 50계약 + 여유
    for r in range(1, ws.max_row+1):
        for c in range(3, MAXC+1):
            cell = ws.cell(r,c)
            cell.value = None
            cell.fill = NOFILL
    ws.cell(1,1).value = f"{client} 보장진단"

    n_ct = len(contracts)

    # ★ LLM 배치 매핑 (앱 자동화): 전체 담보 1회 호출 -> 표준명/종번호
    std_list = load_std_dambo(ws)
    all_raw = sorted({raw for c in contracts for raw in c['dambo']})
    LLMMAP = llm_resolve(all_raw, std_list)
    unmapped = []  # (회사, 담보명, 금액) — 마스터 미수록/매핑실패 -> [확인]
    cancer_trace = []  # ★v30h 암 블록 기재 근거 — (회사, 원담보명, 기재행, 금액). 일반암 과다합산 즉시 추적
    surg_trace = []    # ★v30g 수술 블록 기재 근거 — (회사, 원담보명, 기재행/슬롯, 금액)
    raw_by_std = {}   # ★v39 워크시트 담보명 카피: 표준명→원본담보명(최댓값 담보 기준)
    heart_trace = []   # ★v29z (지점장 2026.07.03): 심장 블록 기재 근거 — (회사, 원담보명, 기재행들, 금액). '없는 값이 튀어나옴' 방지용 감사 로그
    silson_trace = []  # ★v29z: 실손 세대 판정 근거 — (회사, 가입일, 상품코드, 판정)

    def _in_sum(ct):
        """월보험료 합계 포함 여부. ★v129 정본: 잔여보험료 > 0 이면 포함(완납이면 0).
           잔여보험료를 못 읽었으면(None) 납입횟수 a>=b(완납)로 폴백."""
        _r = ct.get('remain')
        if _r is not None: return _r > 0
        return not _is_paid_up(ct.get('pay_period',''), ct.get('pay_count',''))

    for i, ct in enumerate(contracts):
        col = 3 + i
        gen  = ct['renewal'] == '갱신'
        paid = not _in_sum(ct)          # ★v199: 헤더 진녹(완납) 판정 = 합계 제외 판정과 동일 근거
        h = ws.cell(1, col)
        h.value = f"{ct['company']}\n{ct['product']}\n[{ct['renewal']}]"
        h.font = W; h.alignment = AL
        h.fill = FILL_GREEN if paid else (FILL_BLUE if gen else FILL_RED)
        pm = ct['premium']
        # ★★v199 지점장 확정 2026.07.23: 보험료 합계를 '=D2+E2+…'가 아니라 단일 '=SUM()'으로 만들기 위해
        #   완납 계약의 보험료 칸은 <b>텍스트</b>로 넣는다. 엑셀 SUM은 텍스트를 무시하므로
        #   금액은 화면에 그대로 보이면서 합계에서만 자동 제외된다(v129 정본 유지).
        if pm and paid:
            ws.cell(2,col).value = f'{pm:,} (완납)'
        else:
            ws.cell(2,col).value = pm if pm else None
        ws.cell(2,col).font = BL if gen else BK
        ws.cell(3,col).value = ct['contract_date']
        ws.cell(4,col).value = ct['expiry_date']
        ws.cell(5,col).value = f"{ct['pay_period']} ({ct['pay_count']})" if ct['pay_period'] else ''
        for r in [3,4,5]: ws.cell(r,col).font = BL if gen else BK

        dambo = ct['dambo']
        jong_acc = {'상해 종수술비(1-5종)':[0]*8, '질병 종수술비(1-5종)':[0]*8}   # ★v29v 8칸 수집 후 기재 시 5/8종 판정
        trio_acc = [0,0,0]   # ★v29y MRI/도수치료/비급여주사
        jong_blue = {'상해 종수술비(1-5종)':False, '질병 종수술비(1-5종)':False}

        # ★ CI/리빙케어/GI 본체 분해 (지점장 지시 2026.06.28): 주계약 최대=사망, 본체=사망의 80%/50%,
        #   본체를 중대한암·중대한뇌졸증·중대한급성심근에 동일 기재 / 사망 전액=일반사망 / 판별실패=주계약 [확인].
        # ★v246: 3열(KB)은 회사명 칸에 상품명이 붙어 오는 경우가 있다 → <b>회사명·상품명 둘 다</b> 본다.
        #   "이건 롯데나 KB나 동일하다"(지점장 확정) — 2열·3열이 같은 CI 규칙을 타야 한다.
        _is_ci = _isci_prod(ct.get('product')) or _isci_prod(ct.get('company'))
        _ci_fix = None      # ★v239 CI 본체 최종 기재 예약(담보 루프 뒤에서 적용)
        _cij = ct.get('ci_jugye') or []
        # ★★★v237 (지점장 지시 2026.07.25, 영구): <b>선지급률(50%형/80%형)을 세부가입현황(상세내역)에서 찾아낸다</b>.
        #   구 코드는 근거가 <b>별첨의 '주계약' 라벨(`ci_jugye`)뿐</b>이었다. 롯데(let:) 별첨엔 '주계약' 라벨이
        #   아예 없어(주계약을 `질병사망`/`상해사망`으로 씀) `_cij=[]` → <b>이 블록을 통째로 건너뛰었다</b>.
        #   그래서 CI 상품인데 선지급 분해가 하나도 안 됐다(v235에서 `_isci_prod`를 고쳐도 여기서 다시 막혔다).
        #   → <b>①별첨 주계약(1순위) ②세부가입현황 계약열(2순위)</b> 순으로 사망액·본체 후보를 확보한다.
        _sebu_ci = ct.get('ci_sebu') or {}
        _cl = ct.get('ci_lines') or {}
        _samang=0; _cand=[]
        # ★★★★★v239 판정 근거 우선순위(영구): ①별첨 '주계약' 라벨 → ②<b>별첨 줄 단위 사망/진단 원값</b> → ③세부가입현황 사망액
        #   ②가 핵심이다. 롯데 별첨은 주계약을 `질병사망`으로 쓰고 <b>추가특약 사망이 별도 줄</b>로 온다.
        #   dambo는 이를 합산하므로(3,000+2,000=5,000) 비율이 깨진다 → <b>줄 단위 개별값</b>으로 봐야 80%가 나온다.
        if _is_ci and _cij:
            _samang = max(_cij); _cand=[v for v in _cij if 0<v<_samang]
        elif _is_ci and (_cl.get('samang') or []):
            _samang = max(_cl['samang'])
            _cand   = [v for v in (_cl.get('cands') or []) if 0<v<_samang]
            print(f"[v239 CI·별첨줄] {ct.get('company')} 주계약사망={_samang:,} "
                  f"(사망줄 {sorted(set(_cl['samang']),reverse=True)}) 진단후보={sorted(set(_cand),reverse=True)[:8]}")
        elif _is_ci and _sebu_ci.get('samang'):
            _samang = _sebu_ci['samang']; _cand=[]
            print(f"[v239 CI·세부내역] {ct.get('company')} 사망={_samang:,} (담보군 합계는 본체로 쓰지 않음)")
        if _is_ci and _samang:
            _bonche=None; _pct=None; _samang_all=_samang; _cisrc=''
            _slset = set(_cl.get('samang') or []) | {_samang}
            # ★★★★★v241 <b>CI보험 찾기 4단계 체크리스트 — 영구지침(지점장 2026.07.25)</b>
            #   지점장 원문: "1)상세 세부내역에 체크 2)[별첨]보험서비스(상품)별보장현황에서 체크
            #   3)중대한*** 이 있는지 체크 4)똑같은 금액이 2개 이상 있다면 50%이라면 2배 계산이 사망보장,
            #   80%형이라면 수학으로 계산해서 사망보장금과 맞는지 체크.
            #   위의 3가지 다 체크해야 하는 <b>까다롭지만 보람된 CI보험 찾기</b>"
            #   → ①②는 소스 2개(`ci_sebu` 세부내역 · `ci_lines` 별첨)를 <b>둘 다</b> 확보해 교차확인,
            #     ③ '중대한OO' 담보가 있으면 <b>그것이 본체</b>(최우선),
            #     ④ 없으면 <b>동일 금액이 2개 이상</b>인 값을 본체 후보로 삼아
            #       <b>50%형이면 ×2, 80%형이면 ÷0.8 한 값이 별첨 사망보장금과 일치</b>하는지 검증한다.
            _jd=[v for v in (_cl.get('jungdae') or []) if 0<v<_samang]
            if _jd:                                   # ③ 중대한*** 최우선
                _bonche=max(_jd); _cisrc='③중대한OO담보'
                for _r2,_p2 in ((0.5,50),(0.8,80)):
                    if round(_bonche/_r2) in _slset: _pct=_p2; break
                if not _pct: _pct=50                  # 사망줄 매칭 없으면 50%형(지점장 확정)
                _samang = round(_bonche/(_pct/100.0))
            else:                                     # ④ 동일 금액 2개 이상 → 사망보장금 일치 검증
                _cnt={}
                for v in _cand: _cnt[v]=_cnt.get(v,0)+1
                _dup=[v for v,c in _cnt.items() if c>=2]
                for v in sorted(_dup, reverse=True):
                    for _r2,_p2 in ((0.5,50),(0.8,80)):
                        if round(v/_r2) in _slset:    # ★사망보장금과 수학적으로 일치할 때만 인정
                            _bonche=v; _pct=_p2; _samang=round(v/_r2); _cisrc='④동일금액2개이상'
                            break
                    if _bonche: break
                if not _bonche:                       # 폴백: 사망 대비 비율 매칭(±2%)
                    for _ratio,_p in ((0.8,80),(0.5,50)):
                        for v in sorted(set(_cand), key=lambda x:-_cand.count(x)):
                            if _samang and abs(v/_samang-_ratio) < 0.02:
                                _bonche=v; _pct=_p; _cisrc='폴백·비율매칭'; break
                        if _bonche: break
            if _bonche:
                print(f"[v241 CI 4단계] {ct.get('company')} {_cisrc} → 본체 {_bonche:,} ÷ {_pct}% "
                      f"= 사망보장 {_samang:,} · 별첨사망총액 {_samang_all:,} · 차액 {_samang_all-_samang:,}→질병사망(80세)"
                      + ("  ✓사망줄 일치" if _samang in _slset else "  (사망줄 미매칭)"))
            _pl=[]
            if _bonche:
                _pl=[f'중대한 암 {_bonche:,}', f'중대한 뇌졸증 {_bonche:,}', f'중대한 급성심근 {_bonche:,}']
            _ci_diag.append({'co':ct.get('company',''),'pd':ct.get('product',''),
                             'samang':_samang,'pct':_pct,'bonche':_bonche or 0,
                             'src':_cisrc or '판정불가',
                             'sebu':bool((ct.get('ci_sebu') or {}).get('samang')),
                             'byul':bool((ct.get('ci_lines') or {}).get('samang')),
                             'placed':len(_pl),'placed_txt':' · '.join(_pl)})
            if not _bonche:
                _sc = sorted(set(_sebu_ci.get('cands') or []), reverse=True)[:6]
                print(f"[v237 CI·확인큐] {ct.get('company')} 사망={_samang:,} — 50%/80% 근접 본체 없음 "
                      f"→ 선지급형이 아니거나(별도 진단비 특약) 상세내역 대조 필요. 세부가입현황 담보군값={_sc}")
                unmapped.append((col, ct.get('company',''), '[확인] CI 선지급률',
                                 0, f"세부가입현황 사망 {_samang:,} · 담보군값 {_sc} — 50%/80% 미해당. 상세내역에서 선지급형 대조 요망"))
            if _bonche:
                # ★v239: 담보 루프가 뒤에서 중대한OO 행에 별첨 원값을 <b>가산</b>하므로(실측 중대한 암 7,800=5,400+2,400)
                #   여기서 기재하면 덮어써진다 → <b>예약해두고 담보 루프 뒤에서 최종 기재</b>한다.
                _ci_fix = {'bonche':_bonche,'samang':_samang,'pct':_pct}
                # ★★★★★v242 뇌 축 판별: ①본체 금액과 <b>정확히 일치</b>하는 뇌 담보의 축 ②없으면 <b>뇌출혈 우선</b>
                #   (지점장: "무조건 뇌졸증이 아니라는거다 뇌졸증 or 뇌출혈이다")
                _br = (ct.get('ci_lines') or {}).get('brain') or []
                _ax = None
                for _a,_v in _br:
                    if _v == _bonche: _ax = _a; break
                if not _ax:
                    _ax = '뇌출혈' if any(a=='뇌출혈' for a,_ in _br) else ('뇌졸증' if _br else None)
                ct['ci_brain'] = {'axis':_ax, 'rows':_br}
                # ★★★v249 (2026.07.26 한정환 KB 3열 실측): 확인사항 CI 진단표의 '중대한OO 배치' 문구가
                #   <b>'중대한 뇌졸증'으로 하드코딩</b>돼 있었다(2606행 `_pl`). 본표(엑셀 34행)는 v242로
                #   축을 따라 <b>중대한 뇌출혈</b>에 기재되는데 <b>진단표만 뇌졸증</b>이라 산출물끼리 어긋났다.
                #   실측 — 신한 본체 2,000·DB생명 본체 2,400 둘 다 축=뇌출혈인데 진단표는 '중대한 뇌졸증' 출력.
                #   `_pl` 생성이 축 판별보다 <b>앞</b>이라 생긴 순서 결함 → 여기서 배치 문구를 최종 갱신한다.
                if _ci_diag:
                    _bxl = '중대한 뇌출혈' if _ax == '뇌출혈' else '중대한 뇌졸증'
                    _ci_diag[-1]['placed_txt'] = ' · '.join(
                        [f'중대한 암 {_bonche:,}', f'{_bxl} {_bonche:,}', f'중대한 급성심근 {_bonche:,}'])
                print(f"[v242 CI뇌축] {ct.get('company')} 본체 {_bonche:,} → <b>중대한 {_ax or '뇌졸증(기본)'}</b>"
                      f"  (별첨 뇌담보 {_br})")
                # ★★★★★v239 이중계산 차단: 별첨 '주계약' 라벨 경로(`ci_jugye`)에서만 dambo에 사망을 주입한다.
                #   롯데 줄단위 경로(`ci_lines`)는 <b>이미 dambo에 '질병사망'이 들어 있어</b> 여기서 또 더하면
                #   일반사망 3,000+5,000=8,000 · 상해사망 26,100(한장표 23,100 초과)로 부푼다 — 실측 확인.
                if _cij:
                    dambo['일반사망(종신주계약)']=dambo.get('일반사망(종신주계약)',0)+_samang
                    dambo['상해사망(종신주계약)']=dambo.get('상해사망(종신주계약)',0)+_samang   # §8.1 종신 1:1
                _rci=nm2r.get('중대한CI적용')   # 사망−본체=선지급 후 잔여 사망보험금(80%형=20%잔여).
                if _rci:
                    ws.cell(_rci,col).value=_samang-_bonche; ws.cell(_rci,col).font = BL if gen else BK
                # ★v29t (지점장 확정 2026.07.02, 김진구 정본): CI추가보장특약 = 급성심근 초과분 → 최대 1건 급성심근경색 행, 잔여 [확인]
                _cex=ct.get('ci_extra') or []
                if _cex and 'CI추가보장특약' in dambo:
                    _mx=max(_cex)
                    dambo.pop('CI추가보장특약', None)
                    _rgs=nm2r.get('급성심근경색')   # 셀 직접 기재(§8.4 CI 재매핑 회피 — 초과분은 '일반' 급성심근경색 행)
                    if _rgs:
                        _ex0=ws.cell(_rgs,col).value
                        ws.cell(_rgs,col).value=(_ex0+_mx) if isinstance(_ex0,(int,float)) else _mx
                        ws.cell(_rgs,col).font = BL if gen else BK
                    _lv=sum(_cex)-_mx
                    if _lv>0: unmapped.append((col, ct['company'], 'CI추가보장특약(잔여)', _lv, '급성심근 초과분 외 잔여 → 약관 확인'))
                dambo.pop('주계약', None)
            else:
                unmapped.append((col, ct['company'], f'주계약(CI 80/50%판별실패 {_cij})', _samang, 'CI 본체비율 불명 → 좌측표 수기'))
                dambo.pop('주계약', None)

        for raw, amt in dambo.items():
            # ★ 심장 묶음담보 6사 정본 매핑(2026.06.29). I20→협심증 / 허혈성칸=단독전용 / 순환계=전체5 / 급성심근=묶음제외 / 빈맥 포함(★지점장 7/1)·심근병증 제외.
            _rn = re.sub(r'\s','',raw)
            _heart_bundle = None
            _co = ct.get('company','')
            if ('진단' in _rn and '수술' not in _rn and '주요치료' not in _rn
                and '산정특례' not in _rn and '혈전' not in _rn):
                # ★★★v225 (2026.07.25 이정화 실측): <b>산정특례·혈전용해는 심장 묶음이 아니다</b>.
                #   실사고 = `중증질환자심장질환산정특례대상진단비<b>Ⅱ</b>(연간1회한)` 의 <b>로마숫자 Ⅱ</b> 때문에
                #   `_rmn()`이 `_t=2`를 반환해 DB '특정Ⅱ' 묶음으로 오인 → <b>급성심근경색 행에 1,000이 배정</b>되고
                #   <b>산정특례심장은 0</b>이 됐다(급성심근경색 8,000 = 정답 7,000 + 오배정 1,000).
                #   `_HB` 후처리 테이블에는 이미 산정특례 제외가 있었는데 <b>이 인라인 블록에만 빠져 있었다</b>.
                # ★v30o 고정(메리츠, 지점장 2026.07.03): 심장질환진단비Ⅰ→허혈성 진단비 / 심장질환진단비Ⅱ→급성심근경색
                if ('심장질환진단' in _rn) and ('허혈' not in _rn) and ('급성심근' not in _rn):
                    # ★양예서/메리츠 어린이: 심장질환진단비Ⅱ→급성심근경색 / Ⅰ→허혈성 진단비 (별첨값 앵커: Ⅰ=600 허혈성, Ⅱ=3000 급성심근)
                    _mn=_rmn(_rn)
                    if _mn==2: _heart_bundle=['급성심근경색']
                    elif _mn==1: _heart_bundle=['허혈성 진단비']
                # ★뇌질환진단비Ⅰ/Ⅱ (메리츠 어린이 등): Ⅱ→뇌졸증(넓음) / Ⅰ→뇌혈관진단비
                if _heart_bundle is None and ('뇌질환진단' in _rn):
                    _mn=_rmn(_rn)
                    if _mn==2: _heart_bundle=['뇌졸증진단비']
                    elif _mn==1: _heart_bundle=['뇌혈관진단비']
                # ★v30z4 성인병진단금(생보·AIA·AIG·라이나·우체국 등) = 급성심근경색(100% 확정) + 뇌졸증/뇌출혈(세부가입 판별).
                #   지점장 반복 확정: [확인] HOLD 폐기. 뇌축 = 계약에 뇌출혈 담보 있으면 뇌출혈, 없으면 뇌졸증(세부가입 뇌혈관 표기도 뇌졸증계로 해석).
                if _heart_bundle is None and ('성인병' in _rn):
                    _brain = '뇌출혈진단비' if any('뇌출혈' in str(_k) for _k in dambo.keys()) else '뇌졸증진단비'
                    _heart_bundle = ['급성심근경색', _brain]
                # ★v29w 심장 범위 재점검(지점장 2026.07.02, 6사 정본 대조):
                # DB 순환계 5종(중증) = 급성심근경색 + 뇌졸중
                if _heart_bundle is None and '순환계' in _rn and '5종' in _rn:
                    _heart_bundle = ['급성심근경색','뇌졸증진단비']
                # DB 순환계 4종 = 협심증·심부전(+빈맥, 심근병증 [확인])
                elif '순환계' in _rn and '4종' in _rn:
                    _heart_bundle = ['협심증','심부전']
                # DB 순환계 3종 = 염증·부정맥
                elif '순환계' in _rn and '3종' in _rn:
                    _heart_bundle = ['염증','부정맥']
                # ★순환계3대(허혈성심장 I20~25 + 뇌혈관 I60~69 + 말초 I70~79) — 삼성·DB 등. 말초=전용행無(누락 감수)
                elif '순환계' in _rn:
                    if ('DB' in _co) or ('디비' in _co):
                        _heart_bundle = ['급성심근경색','빈맥','부정맥','심부전']   # DB 순환계3대=심장정지I46.0·부정맥I47~49·심부전I50
                    else:
                        _heart_bundle = ['협심증','급성심근경색','허혈성 진단비','뇌혈관진단비']
                # ===== BARUM 10사 질병코드 분류표 정본(2026.07.05 지점장 확정): 특정Ⅰ/Ⅱ 라벨=회사마다 다름 → 회사별 표대로 =====
                elif any(_k in _rn for _k in ('심혈관','심장','허혈','부정맥','빈맥','심부전','심근병','판막','협심','전도','방실')):
                    _t=_rmn(_rn)
                    _i49excl=('제외' in _rn) and (('I49' in _rn) or ('부정맥' in _rn))   # ★(기타심장부정맥제외)=Ⅰ에서 I49 뺀 묶음(부정맥 담보 아님)
                    _i49=(not _i49excl) and (('I49' in _rn) or ('기타부정맥' in _rn) or ('기타심장부정맥' in _rn))
                    # 흥국·롯데: 특정Ⅰ=급성심근 / 특정Ⅱ=협심증+허혈+염증 / 롯데 15대=판막·심근병·빈맥·심부전
                    if ('흥국' in _co) or ('롯데' in _co):
                        if _i49excl: _heart_bundle=['협심증','허혈성 진단비','빈맥','심부전']   # 흥국 특정심혈관질환(기타심장부정맥제외)=협심·허혈·빈맥·심부전(별표70)
                        elif _i49: _heart_bundle=['부정맥']
                        elif '심근병' in _rn: _heart_bundle=['심근병증']
                        elif '15대' in _rn: _heart_bundle=['심장판막','심근병증','빈맥','심부전']
                        elif ('방실' in _rn) or ('전도' in _rn): pass   # 전용행無→[확인]
                        elif ('주요' in _rn and ('염증' in _rn or '심장염' in _rn)) or ('심낭' in _rn): _heart_bundle=['염증']
                        elif _t==1: _heart_bundle=['급성심근경색']
                        elif _t==2: _heart_bundle=['협심증','허혈성 진단비','염증']
                    # ★DB(정본 재수정): 특정Ⅰ=협심증·허혈·염증 / 특정Ⅱ=급성심근 / 특정Ⅲ=판막·빈맥·심부전 / 심근병증
                    elif ('DB' in _co) or ('디비' in _co):
                        if _t==2: _heart_bundle=['급성심근경색']
                        elif _t==3: _heart_bundle=['심장판막','빈맥','심부전']
                        elif '심근병' in _rn: _heart_bundle=['심근병증']
                        elif _i49: _heart_bundle=['부정맥']
                        elif _t==1: _heart_bundle=['협심증','허혈성 진단비','염증']
                    # 한화·NH농협: Ⅰ=협심증+허혈+빈맥+부정맥+심부전 / Ⅱ=급성심근 / (I49제외)=부정맥 뺀 묶음 / 심근병증
                    elif ('한화' in _co) or ('농협' in _co) or ('NH' in _co):
                        if _i49excl: _heart_bundle=['협심증','허혈성 진단비','빈맥','심부전']   # Ⅰ에서 I49(부정맥) 제외 묶음
                        elif _t==2: _heart_bundle=['급성심근경색']
                        elif '심근병' in _rn: _heart_bundle=['심근병증']
                        elif ('주요' in _rn and ('염증' in _rn or '심장염' in _rn)): _heart_bundle=['염증']
                        elif _i49: _heart_bundle=['부정맥']
                        elif '특정질환' in _rn: _heart_bundle=['협심증','허혈성 진단비','빈맥','심부전']   # 한화 심혈관특정질환=Ⅰ에서 I49제외
                        elif _t==1: _heart_bundle=['협심증','허혈성 진단비','빈맥','부정맥','심부전']
                    # KB: 특정Ⅰ=협심증+허혈+빈맥+심부전(염증X·부정맥X) / Ⅱ=급성심근 / 심장판막=판막+염증 / I49=부정맥(빈맥X)
                    elif ('KB' in _co) or ('케이비' in _co):
                        if _t==2: _heart_bundle=['급성심근경색']
                        elif '심근병' in _rn: _heart_bundle=['심근병증']
                        elif '판막' in _rn: _heart_bundle=['심장판막','염증']
                        elif _i49: _heart_bundle=['부정맥']
                        elif _t==1 or ('확대' in _rn and '심장' in _rn) or ('특정심장' in _rn): _heart_bundle=['협심증','허혈성 진단비','빈맥','심부전']
                    # 현대(정본 재수정, 6가지): 허혈성심장=협심증+허혈 / 특정허혈=급성심근 / 특정Ⅰ=빈맥+심부전 / 특정Ⅱ=급성심근 / 주요염증 / 특정2대+I49=부정맥
                    elif '현대' in _co:
                        if '특정허혈' in _rn: _heart_bundle=['급성심근경색']
                        elif ('허혈성심장' in _rn) or ('허혈심장' in _rn): _heart_bundle=['협심증','허혈성 진단비']
                        elif '심근병' in _rn: _heart_bundle=['심근병증']
                        elif ('주요' in _rn and ('염증' in _rn or '심장염' in _rn)): _heart_bundle=['염증']
                        elif ('특정2대' in _rn) or ('방실' in _rn) or ('전도' in _rn) or _i49: _heart_bundle=['부정맥']   # 특정2대+기타부정맥(I49) 병합→부정맥(전도장애 전용행無)
                        elif _t==2: _heart_bundle=['급성심근경색']   # ★현대 특정Ⅱ=급성심근경색(정본 재수정)
                        elif _t==1 or '심혈관' in _rn: _heart_bundle=['빈맥','심부전']
                    # 삼성·메리츠: 허혈성심장질환 6가지 → 급성심근+협심증+허혈성 (메리츠는 기존 심장질환진단Ⅰ/Ⅱ와 병존)
                    elif ('삼성' in _co) or ('메리츠' in _co):
                        # ★★단독담보 원칙(지점장 확정 2026.07.14, 최상위 · 2026.07.25 재확정): 회사담보명이 '허혈성심장질환진단비'
                        #   단독이면 어느 회사든 분해 금지 → '허혈성 진단비' 행 단독. 묶음 수식어가 붙은 것만 분해.
                        # ★★★v206 (2026.07.25 양*선 KB리포트 실측 회귀수정): 판정이 `^...$` 완전일치라
                        #   회사담보명 앞에 <b>'갱신형 '</b>이 붙은 '갱신형 허혈성심장질환진단비'가 _solo=False로 떨어져
                        #   급성심근+협심증+허혈성 3행으로 <b>분해</b>됐다(실측: 협심증 1,000 신규발생 · 급성심근 2,000→3,000).
                        #   → 접두 수식어(갱신형·비갱신형·정액·실손·무배당·[건강] 등)를 떼고 판정한다.
                        _rn_core = re.sub(r'^(?:갱신형|비갱신형|무배당|정액|실손|\[[^\]]*\])+', '', re.sub(r'\s', '', _rn))
                        _solo = bool(re.match(r'^허혈(성)?심장질환진단(비)?$', _rn_core))
                        if _solo: pass
                        elif ('허혈성심장' in _rn) or ('허혈심장' in _rn): _heart_bundle=['급성심근경색','협심증','허혈성 진단비']
            if _heart_bundle:
                for _bt in _heart_bundle:
                    _br = nm2r.get(_bt)
                    if _br:
                        _ex = ws.cell(_br,col).value
                        ws.cell(_br,col).value = (_ex+amt) if isinstance(_ex,(int,float)) else amt
                        ws.cell(_br,col).font = BL if gen else BK
                heart_trace.append((ct['company'], raw, ' · '.join(_heart_bundle), amt))   # ★v29z 근거 기록
                continue
            # ★ 우선순위 역전: 확정 규칙(resolve2) 먼저 → 못 잡은 것만 Haiku(llm_resolve).
            #   Haiku가 간병인·암주요치료비·하이클래스 등 확정담보를 가로채 누락시키던 문제 차단.
            std, jong = resolve2(raw)
            jong = jong or get_종번호(raw)
            if not std:                       # 규칙이 못 잡은 것만 LLM 폴백
                # ★v30m 수술·일당은 resolve_kw(+DMAP)가 최종 판정. resolve_kw가 [확인](None)으로 보낸 변형을
                #   Haiku가 질병/상해수술비·상해/질병일당 base 행으로 되끌어와 합산하던 문제 차단 → 변형은 그대로 [확인].
                if re.search(r'수술|일당', raw):
                    m = {}
                else:
                    m = LLMMAP.get(raw) or {}
                    std = m.get('std')
                    if not jong: jong = m.get('jong', 0) or 0
            else:
                m = {}
            if std and _isci_prod(ct['product']):
                std = {'일반암':'중대한 암','뇌졸증진단비':'중대한 뇌졸증','급성심근경색':'중대한 급성심근'}.get(std, std)
            elif std and not _isci_prod(ct['product']):
                # ★2026.07.12 지점장 확정: 상품명에 CI/GI/리빙케어가 없으면 진짜 CI가 아니다.
                #   '중대한 암' → 일반암진단비로 산입 / 그 외 '중대한OO'는 가짜 → 전부 무시(기재 안 함).
                if std == '중대한 암':
                    std = '일반암'
                elif std in ('중대한 뇌졸증', '중대한 급성심근', '중대한CI적용', '중대한 뇌출혈'):
                    unmapped.append((col, ct['company'], raw, amt, 'CI/GI/리빙케어 상품 아님 → 중대한OO 무시'))
                    continue
            # ★★★v220 영구 가드(2026.07.25 실사고): <b>일당 행에 진단비 금액이 들어가면 그건 일당이 아니다</b>.
            #   실사고 = 간호통합병동 행에 <b>2,000(2천만원)</b>이 찍혔다. 일당은 하루당 지급액이라
            #   실무 최대가 <b>30~50만원</b>이고 100만원을 넘는 일당 담보는 존재하지 않는다.
            #   → 일당 계열 행에 <b>100 초과</b> 금액이 오면 매핑 오류로 보고 [확인]큐로 보낸다(조용한 오출고 차단).
            _DAILY = ('질병일당','상해일당','간병인','간병인지원일당','간호통합병동','질병종합병원일당','상급병원질병입원일당',
                      '1인실 상급병원','1인실 종합병원','질병중환자실','상해중환자실',
                      '질병수술일당','상해수술일당','암일당')
            if std in _DAILY and isinstance(amt,(int,float)) and amt > 100:
                unmapped.append((col, ct['company'], raw, amt,
                                 f'[확인] 일당 행에 100만원 초과({amt}) — 진단비·수술비 오매핑 의심'))
                continue
            if std in ('골절(치아파절포함)','골절(치아파절제외)','화상진단비') and amt>=100:
                unmapped.append((col, ct['company'], raw, amt, '등급별 100만↑ 제외'))  # 등급별 → 합산·기재 안 함
                continue
            if std=='합의금' and amt>25000:   # 합의금 최대 2.5억, 초과는 불가 → [확인]
                unmapped.append((col, ct['company'], raw, amt, '합의금 2.5억 초과(불가)'))
                continue
            if std=='입원':
                # ★★v91 수정(지점장 확정 2026.07.19 · 장혜경 실데이터로 확인):
                #   구 규칙 '입원은 무조건 5,000 고정'이 <b>구실손의 실제 한도를 뭉갰다</b>.
                #   실측: DB 0604_TM(2006년형 1세대) 질병입원의료비 <b>500</b> → 5,000으로 부풀려짐.
                #   한장보장표(실손입원 상해 5,100 / 질병 5,500)와 불일치 = 등식1 위반.
                #   → <b>별첨 명시값을 그대로 쓴다.</b> 금액을 못 읽었을 때만 5,000을 기본값으로 넣는다.
                if not amt: amt=5000
            # ★v34 암주요치료비 10,000 강제 폐기(지점장 2026.07.09): 실제 가입금액 사용. 하이클래스는 별도행 합산.
            blue = gen or ('갱신' in raw)      # ★ 담보명에 (갱신) 표시 -> 파랑
            # 수술비 1~5종 -> 종별 슬래시 누적
            if std == '종수술비공통' and 1 <= jong <= 5:   # ★v29q-12 상해/질병 미표기 → 상해·질병 양쪽 동일 기재
                for _k in ('상해 종수술비(1-5종)','질병 종수술비(1-5종)'):
                    jong_acc[_k][jong-1] += amt
                    if blue: jong_blue[_k] = True
                surg_trace.append((ct['company'], raw, f'상해·질병 종수술 양쪽 {jong}종 슬롯', amt))   # ★v30g
                continue
            if std in jong_acc and 1 <= jong <= 5:
                jong_acc[std][jong-1] += amt
                if blue: jong_blue[std] = True
                surg_trace.append((ct['company'], raw, f'{std} {jong}종 슬롯', amt))   # ★v30g
                continue
            # ★v29y (지점장 2026.07.02): MRI·도수치료·비급여주사 = 'MRI/도수치료/비급여주사' 한 행 슬래시(1-5종 방식)
            if std in ('MRI','도수치료','비급여주사'):
                _ti={'MRI':0,'도수치료':1,'비급여주사':2}[std]
                trio_acc[_ti]=max(trio_acc[_ti],amt)   # 실손 계열=중복합산 금지, 대표 최댓값
                continue
            if std == '__무시__':            # ★v30z 지점장 무시지정 담보(전이암진단비·고액항암치료비) = 완전 드롭, [확인]에도 미노출
                continue
            r = nm2r.get(std)
            if r is None and std=='n대수술비': r = nm2r.get('120대수술비')   # ★v30c std↔행명 별칭(1XX대 대표행)
            if r is None and std:             # 공백무시 재매칭 (화상 '진 단 비' 등)
                r = nm2r_norm.get(re.sub(r'\s','', std))
            if not std or r is None:          # 마스터 미수록/매핑실패 -> [확인]
                unmapped.append((col, ct['company'], raw, amt, m.get('note','') or ''))
                continue
            # 2대 주요치료비는 뇌혈관·심장 두 칸 모두 기재(동일 담보, 양쪽 표기). 그 외는 단일 행.
            # ★★★v216 (지점장 지시 2026.07.25, 영구): <b>혈전용해치료비 누락 수정</b>.
            #   마스터에 '혈전용해치료비' 행이 <b>2개</b>다 — <b>뇌 블록 37행 · 심장 블록 50행</b>.
            #   그런데 `nm2r[담보명]=행`은 <b>뒤에 나온 50행이 앞의 37행을 덮어써</b> 뇌 37행에는
            #   <b>어떤 값도 영원히 들어가지 않았다</b>(구조적 누락). 게다가 실측에서 메리츠
            #   '뇌혈관혈전용해치료비 200' + '급성심근경색증혈전용해치료비 200'이 <b>같은 50행에 400으로 합산</b>됐다.
            #   → <b>담보명의 축으로 갈라 배정한다</b>: 뇌 축=37행 / 심장 축(심근·심장·허혈·관상동맥·심혈관)=50행 /
            #   <b>축 미표기 또는 심뇌 동시 표기 = 양쪽 행에 각 100% 동일 금액</b>(2대 주요치료비·묶음담보 공통원칙 §8.3.1과 동일).
            if std == '2대 주요치료비':
                target_rows = nm2r_multi.get(std, [r])
            elif std == '혈전용해치료비':
                _hjall = nm2r_multi.get(std, [r])
                _hjn = _norm(raw)
                _hjb = ('뇌' in _hjn)
                _hjh = any(_k in _hjn for _k in ('심근','심장','허혈','관상동맥','심혈관'))
                if _hjb and not _hjh:   target_rows = [_hjall[0]]     # 뇌 전용 → 뇌 블록 행
                elif _hjh and not _hjb: target_rows = [_hjall[-1]]    # 심장 전용 → 심장 블록 행
                else:                   target_rows = _hjall          # 축 미표기·심뇌동시 → 양쪽 각 100%
            else:
                target_rows = [r]
            for tr in target_rows:
                existing = ws.cell(tr,col).value
                # ★★★v215 (지점장 확정 2026.07.25, 영구): <b>간병인지원일당 = 대표(max) 1건</b>.
                #   근거: 간병인지원 질병입원일당 3 + 간병인지원 상해입원일당 3 은 <b>질병·상해 택일 지급</b>이라
                #   <b>합산 6이 아니라 3</b>이다(둘 중 하나만 기재). 간병인·간호통합병동·1인실과 동일 처리.
                # ★항암방사선약물도 대표(max) 1건 — 항암약물치료비 / 항암방사선치료비 /
                #   항암방사선약물치료비 / 항암약물방사선치료비는 <b>이름만 다른 같은 담보</b>다(합산 금지).
                _rep1 = std in ('표적항암치료비','다빈치로봇수술비','n대수술비','입원','통원','약값','약','간병인','간병인지원일당','창상봉합술','항암방사선약물','암수술','중입자치료비','암주요치료비','통합전이암','간호통합병동','합의금','1인실 상급병원','1인실 종합병원')   # ★v198 합의금=대표1개 / ★v208 1인실 / ★v215 간병인지원일당=택일 대표(max)
                _rep1 = _rep1 or ('통합' in raw and std in ('일반암','유사암(갑.기.경.제)','통합전이암'))   # ★v30a §8.2 통합 계열=대표금액 1개
                if _rep1 and isinstance(existing,(int,float)):
                    ws.cell(tr,col).value = max(existing, amt)   # 표적·n대·창상봉합=대표 최댓값1건(★v29q-6) / 실손=중복합산 안함(한도)
                else:
                    ws.cell(tr,col).value = (existing+amt) if isinstance(existing,(int,float)) else amt
                # 실손(입원/통원/약값)·일상배상책임은 갱신·비갱신 무관 항상 파랑
                # ★★★v210 (지점장 확정 2026.07.25, 영구): <b>간병인 · 간호통합병동 2가지는 '항상 파랑' 강제 폐기</b>.
                #   보험료 · 가입년일 · 만기일자 · 총납입기간(=계약 갱신 판정) 또는 담보명의 <b>[갱신] 표기</b>에 따라
                #   갱신=파랑 / 비갱신=검정으로 <b>일반 담보와 동일하게</b> 칠한다(구 v139 '간병인 계열 3행 무조건 파랑' 폐기).
                ws.cell(tr,col).font = BL if (blue or std in ('입원','통원','약값','약','일상배상책임')) else BK
                # ★v39 워크시트용 원본담보명 수집(그 표준명 중 최댓값 담보의 raw 1개)
                _WS_STD = ('암주요치료비','하이클래스(암)','2대 주요치료비','산정특례뇌혈관','산정특례심장','일반암','뇌혈관진단비','뇌졸증진단비','급성심근경색','허혈성 진단비')
                if std in _WS_STD:
                    _prev = raw_by_std.get(std)
                    if _prev is None or amt >= _prev[1]:
                        raw_by_std[std] = (str(raw).strip(), amt)
                if std in {'협심증','심부전','빈맥','염증','부정맥','심근병증','심장판막','산정특례심장','2대 주요치료비','허혈성 진단비','급성심근경색','중대한 급성심근','혈전용해치료비','심장수술비','허혈성수술비'}:
                    heart_trace.append((ct['company'], raw, std, amt))   # ★v29z 심장 단독 기재 근거
                if '수술' in str(std) or std == '창상봉합술':
                    surg_trace.append((ct['company'], raw, std, amt))   # ★v30g 수술 기재 근거
                if std in {'일반암','유사암(갑.기.경.제)','통합전이암','고액암','중대한 암','암주요치료비','하이클래스(암)','암수술','암일당'}:
                    cancer_trace.append((ct['company'], raw, std, amt))   # ★v30h 암 기재 근거

        for nm, vals in jong_acc.items():     # 종수술비 슬래시 기재(§6)
            if any(vals):
                # ★v29v (지점장 2026.07.02): 6~8종 값이 있으면 그 계약의 종수술은 8단계 → (1-8종) 행에 8칸 슬래시,
                #   아니면 기존대로 (1-5종) 행에 5칸 슬래시.
                if any(vals[5:]):
                    tgt=nm.replace('(1-5종)','(1-8종)'); use=vals
                else:
                    tgt=nm; use=vals[:5]
                r = nm2r.get(tgt) or nm2r.get(nm)
                if r:
                    ws.cell(r,col).value = '/'.join(str(x) for x in use)
                    ws.cell(r,col).font = BL if (gen or jong_blue[nm]) else BK

        if any(trio_acc):   # ★v29y MRI/도수/주사 슬래시 기재(실손 계열=항상 파랑)
            _rt=nm2r.get('MRI/도수치료/비급여주사')
            if _rt:
                ws.cell(_rt,col).value='/'.join(str(x) for x in trio_acc)
                ws.cell(_rt,col).font=BL

        # ★★★★★v239 CI 본체 최종 기재 — 담보 루프가 끝난 뒤에 한다.
        #   정본: 본체(=사망액 × 선지급률)를 <b>중대한 암·중대한 뇌졸증·중대한 급성심근 3행에 동일 금액</b>으로 기재하고,
        #   <b>초과분은 일반 담보 행</b>에 남긴다. 중대한CI적용 = 사망 − 본체.
        #   (지점장 지적 "각 중대한의 금액이 다 틀리다" = 구 코드가 별첨 담보 원값을 그대로 옮겨 3행이 제각각이었다)
        if _ci_fix:
            _bc=_ci_fix['bonche']; _sm=_ci_fix['samang']
            # ★★★★★v242 영구지침(지점장 지시 2026.07.25): <b>CI 본체의 뇌 축은 '뇌졸증'이 아니라 '뇌졸증 or 뇌출혈'</b>이다.
            #   지점장 원문 — "정답은 <b>중대한뇌출혈</b>이다. 왜냐면 사망이 3천에 중대한뇌출혈은 2400 + 추가로 뇌출혈 가입이고.
            #   <b>신한도 뇌출혈</b>이다. … <b>무조건 뇌졸증이 아니라는거다 뇌졸증 or 뇌출혈이다.</b>"
            #   ★실측 근거 — DB 별첨 `뇌출혈진단 2,400`이 <b>본체와 정확히 일치</b>(나머지 1,000+1,000은 추가 가입분),
            #     신한은 뇌 담보가 `뇌출혈진단`뿐. 구 코드는 무조건 '중대한 뇌졸증'에 넣어 <b>축이 통째로 틀렸다</b>.
            #   ★판별: 그 계약 별첨의 <b>뇌 담보 이름</b>을 본다 — ①본체 금액과 일치하는 담보의 축 ②없으면 뇌출혈 우선.
            _braw = ct.get('ci_brain') or {}
            _bax = '중대한 뇌출혈' if _braw.get('axis')=='뇌출혈' else '중대한 뇌졸증'
            _bovf = '뇌출혈진단비' if _braw.get('axis')=='뇌출혈' else '뇌졸증진단비'
            for _nm,_ovf in (('중대한 암','일반암'),(_bax,_bovf),('중대한 급성심근','급성심근경색')):
                _r=nm2r.get(_nm); _ro=nm2r.get(_ovf)
                if not _r: continue
                # ★★★★★v242: <b>중대한 행 + 일반 행을 합쳐서</b> 본체만큼 중대한 행에, 나머지를 일반 행에 배분한다.
                #   구 코드는 <b>중대한 행의 원값만</b> 봤다 → 뇌출혈처럼 값이 <b>일반 행(뇌출혈진단비)에만</b> 있으면
                #   `원값 0 ≤ 본체`라 스킵되어 <b>중대한 뇌출혈이 0으로 남았다</b>(지점장 지적 "이번에도 틀렸다").
                #   ★합계를 보존하므로 한장보장표(중대한OO + 일반 합산)는 어느 쪽이든 불변이다.
                _v1=ws.cell(_r,col).value;  _v1=_v1 if isinstance(_v1,(int,float)) else 0
                _v2=ws.cell(_ro,col).value if _ro else 0; _v2=_v2 if isinstance(_v2,(int,float)) else 0
                _tot=_v1+_v2
                # ★★★★★v253 영구지침(지점장 지시 2026.07.26 "오늘 무조건 CI는 문제 없어야 한다"):
                #   <b>합계가 0이어도 본체를 중대한OO 행에 기재한다</b>.
                #   구 코드는 `_tot<=0`이면 continue라 <b>중대한OO 3행이 통째로 0</b>이 됐다.
                #   ★왜 0이 되나 — 교보생명 큰사랑CI처럼 <b>별첨 담보명이 '주계약'·'○○특약'뿐</b>인 생보 CI는
                #     암·뇌·심 담보가 종류별로 dambo에 없다(4p 세부가입현황에만 있다). 그래서 뺄 일반행이
                #     없는데 정본 배분이 "합계에서 나눈다"라서 <b>나눌 합계가 0</b> → 스킵 → 기재 0.
                #   ★실측(김O구 교보): 1차 별첨 주계약 4,000(=일반사망)×80% = 2차 세부가입현황 질병사망
                #     4,000×80% = 본체 3,200. 두 경로 일치인데 <b>중대한 암·뇌졸증·급성심근이 전부 0</b>이었다.
                #   ★합계가 있으면 기존 배분(합계 보존) 그대로 — 신한·DB생명 회귀 없음.
                # ★★★★★v254 영구지침(지점장 지시 2026.07.26): <b>별첨 '주계약' 라벨 경로(교보 등)는
                #   CI 본체와 일반 담보가 별개다</b> → 중대한행에 <b>본체를 기재하고 일반행은 손대지 않는다</b>.
                #   지점장 원문 "니가 올린건 CI + 일반 담보다 / 교보의 일반 급성심근 담보 정답".
                #   ★실측(김O구 교보): 급성심근 4,200 = <b>CI 본체 3,200 + 교보 일반 급성심근 1,000</b>.
                #     구 코드는 합계에서 본체를 빼는 배분이라 <b>일반 1,000을 지워</b> 3,200이 됐다.
                #   ★반대로 별첨 <b>담보명</b>으로 CI가 표시된 계약(신한 `뇌출혈진단`·DB생명)은 담보값에
                #     본체가 <b>포함</b>돼 있으므로 기존 배분(합계 보존)을 유지한다 — 회귀 없음.
                if _cij:
                    ws.cell(_r,col).value=_bc; ws.cell(_r,col).font = BL if gen else BK
                    print(f"[v254 CI가산] {ct.get('company')} {_nm} ← 본체 {_bc:,} · {_ovf} {_v2:,} 유지(일반 담보 별개)")
                    continue
                if _tot<=0:
                    ws.cell(_r,col).value=_bc; ws.cell(_r,col).font = BL if gen else BK
                    print(f"[v253 CI기재] {ct.get('company')} {_nm} ← 본체 {_bc:,} (일반행 값 0)")
                    continue
                # ★★★★★v253: 중대한 행 = <b>본체 그대로</b>(구 코드 `min(_bc,_tot)`은 합계가 본체보다
                #   작을 때 <b>본체가 잘려나갔다</b>). 정본 = "중대한OO 3행에 <b>동일 금액</b> 기재".
                #   ★실측(김O구 교보): 급성심근 합계 1,000 < 본체 3,200 → 1,000만 기재되어 한장표 4,200에
                #     2,200이 비었다. 본체 3,200을 기재하면 다른 계약분 1,000과 합쳐 <b>4,200</b>이 맞는다.
                #   ★합계 ≥ 본체인 경우는 `min(_bc,_tot)==_bc`라 <b>기존과 동일</b> — 신한·DB생명 회귀 없음.
                _put=_bc
                ws.cell(_r,col).value=_put; ws.cell(_r,col).font = BL if gen else BK
                if _ro:
                    _rest=max(0,_tot-_put)
                    ws.cell(_ro,col).value=(_rest if _rest>0 else None)
                    if _rest>0: ws.cell(_ro,col).font = BL if gen else BK
            # ★★★v242: <b>축이 아닌 쪽의 '중대한' 뇌 행은 일반 행으로 되돌린다</b>.
            #   CI 상품이라 resolve가 `뇌졸중진단`을 '중대한 뇌졸증'으로 리네임했는데, 축이 뇌출혈로 확정되면
            #   그 값은 <b>CI 본체가 아니라 별도 일반 담보</b>다(실측 DB 뇌졸중진단 1,000). 합계는 보존된다.
            _other = ('중대한 뇌졸증','뇌졸증진단비') if _bax=='중대한 뇌출혈' else ('중대한 뇌출혈','뇌출혈진단비')
            _ro2=nm2r.get(_other[0]); _rn2=nm2r.get(_other[1])
            if _ro2 and _rn2:
                _vo=ws.cell(_ro2,col).value
                if isinstance(_vo,(int,float)) and _vo:
                    _vn=ws.cell(_rn2,col).value
                    ws.cell(_rn2,col).value=(_vn if isinstance(_vn,(int,float)) else 0)+_vo
                    ws.cell(_rn2,col).font = BL if gen else BK
                    ws.cell(_ro2,col).value=None
                    print(f"[v242 축정리] {ct.get('company')} {_other[0]} {_vo:,} → {_other[1]} (축={_bax})")
            _rci=nm2r.get('중대한CI적용')
            if _rci:
                ws.cell(_rci,col).value=_sm-_bc; ws.cell(_rci,col).font = BL if gen else BK
            print(f"[v239 CI본체] {ct.get('company')} 사망{_sm:,} × {_ci_fix['pct']}% = 본체 {_bc:,} → 중대한OO 3행 동일기재 · 잔여 {_sm-_bc:,}")

        # ★ §8 생보 종신(만기 9999): 일반사망(종신) + 상해사망 1:1 복제
        if ct['expiry_date'].startswith('9999'):
            r_il = nm2r.get('일반사망'); r_sh = nm2r.get('상해사망'); r_jb = nm2r.get('질병사망(80세)')
            # ★★★★★v239 (지점장 지적 2026.07.25 "엑셀에 사망이 안 잡혀있다"): 정본 §8.1은
            #   <b>"일반사망 = 생명보험사 만기 9999(종신)으로 표기된 사망"</b>인데, 구 코드는
            #   담보명이 <b>'일반사망'</b>일 때만 6행에 넣고 <b>'질병사망'</b>은 질병사망(80세) 8행에 뒀다.
            #   → CI 종신보험(신한 라이프케어CI·DB CI종신)은 별첨 담보명이 <b>'질병사망'</b>이라
            #   <b>일반사망 행이 통째로 비었고</b>, 그 결과 선지급률 계산 근거가 엑셀에 없었다.
            #   ★생보 + 만기 9999면 질병사망(80세) 값을 <b>일반사망(종신) 행으로 이동</b>한다.
            #   (질병사망 합계는 '일반사망 + 질병사망(80세)'로 보므로 <b>한장보장표 총액은 불변</b> — 실측 확인)
            _life = any(k in (ct.get('company') or '') for k in ('생명','라이프','AIA','메트라이프','우체국','공제'))
            if _life and r_il and r_jb:
                _vjb = ws.cell(r_jb,col).value
                if isinstance(_vjb,(int,float)) and _vjb:
                    # ★★★★★v240(지점장 지시): CI 계약이면 <b>주계약 사망만 일반사망(종신)</b>으로 옮기고,
                    #   <b>차액(추가 특약 사망)은 질병사망(80세) 행에 그대로 남긴다</b>.
                    #   실측 신한 — 별첨 사망 10,000 중 주계약 4,000(본체 2,000÷50%) → 일반사망 4,000 / 질병사망 6,000.
                    #   (CI가 아니거나 주계약 판별 실패면 종전대로 전액 이동 — 한장표 총액은 어느 쪽이든 불변)
                    _mv = _vjb
                    if _ci_fix and 0 < _ci_fix.get('samang',0) < _vjb:
                        _mv = _ci_fix['samang']
                    _v0 = ws.cell(r_il,col).value
                    ws.cell(r_il,col).value = (_v0 if isinstance(_v0,(int,float)) else 0) + _mv
                    ws.cell(r_il,col).font = BL if gen else BK
                    ws.cell(r_jb,col).value = (_vjb-_mv) if (_vjb-_mv) > 0 else None
                    print(f"[v240 종신사망] {ct.get('company')} 별첨사망 {_vjb:,} → 일반사망(종신) {_mv:,}"
                          + (f" · 질병사망(80세) {_vjb-_mv:,}(추가특약)" if _vjb-_mv>0 else ""))
            v = ws.cell(r_il,col).value if r_il else None
            if isinstance(v,(int,float)) and r_sh and not isinstance(ws.cell(r_sh,col).value,(int,float)):
                ws.cell(r_sh,col).value = v
                ws.cell(r_sh,col).font = BL if gen else BK

        # ★ 실손 통원/약값 디폴트 (지점장 2026.06.28): ①별첨 명시값 최우선 → ②입원3,000 구형=통원10·약5
        #   → ③2021.06 이전: 손보 통원25·약5 / 생보 통원20·약10 → ④4세대(2021.07~): 통원20·약0(통원포함).
        _rip=nm2r.get('입원'); _rtw=nm2r.get('통원'); _ryk=nm2r.get('약값')
        _ipv=ws.cell(_rip,col).value if _rip else None
        if isinstance(_ipv,(int,float)) and _ipv:   # 이 계약에 실손(입원) 존재
            _life=any(k in (ct['company'] or '') for k in ('생명','라이프','AIA','메트라이프','우체국','공제'))
            def _ym(d):
                try: return int(str(d)[:4])*100+int(str(d)[5:7])
                except: return 0
            _g4=(silson_gen(ct.get('contract_date',''), None, ct.get('product',''), _has_nonpay3(ct.get('dambo'))) in ('4세대','5세대'))   # ★v29v 상품코드 반영
            _guhy=(_ipv==3000)                            # 입원한도 3,000=구형
            _twc=ws.cell(_rtw,col).value if _rtw else None
            _ykc=ws.cell(_ryk,col).value if _ryk else None
            # ★v215: 통원 디폴트 판정에도 1세대(가입일 기준)를 포함한다. 구 코드는 입원한도 3,000(_guhy)일
            #   때만 10을 넣어, 가입일이 2009.09 이전인데 입원한도가 3,000이 아닌 1세대는 25/20이 됐다.
            _g1a=str(silson_gen(ct.get('contract_date',''), _ipv, ct.get('product',''), _has_nonpay3(ct.get('dambo')))).startswith('1세대')
            if _rtw and not isinstance(_twc,(int,float)):  # ① 별첨 통원 없을 때만 디폴트
                _twd = 10 if (_guhy or _g1a) else (20 if _g4 else (20 if _life else 25))
                ws.cell(_rtw,col).value=_twd; ws.cell(_rtw,col).font=BL
            # ★★★v215 (지점장 확정 2026.07.25, 영구): <b>실손 1세대 = 입원 3,000 / 통원 10 / 약값 0</b>.
            #   1세대는 외래+약제가 <b>통원 한 한도로 통합</b>돼 있어 <b>약값 담보 자체가 없다</b> → 약값 칸은 비운다.
            #   ★버그2건 실측: ①`silson_gen`은 입원한도 3,000이면 <b>'1세대(구형)'</b>을 반환하는데
            #     구 코드는 `== '1세대'`로 비교해 <b>_g1이 항상 False</b>였다 → `_guhy` 분기로 빠져 <b>약값 5가 찍혔다</b>
            #     (지점장 지적 "보장진단서 오류"의 원인). → `.startswith('1세대')`로 수정.
            #     ②구 코드는 '별첨에 약값이 없을 때만' 0으로 뒀다 → 별첨에 약값이 인쇄돼 있으면 그대로 들어갔다.
            #     → 1세대는 <b>별첨 명시값이 있어도 약값 칸을 지운다</b>(강제).
            _g1=str(silson_gen(ct.get('contract_date',''), _ipv, ct.get('product',''), _has_nonpay3(ct.get('dambo')))).startswith('1세대')   # ★v215 (구 v41 == '1세대' 버그)
            if _ryk and _g1:                               # ★v215 1세대=약값 없음(강제 미기재)
                if isinstance(_ykc,(int,float)) and _ykc:
                    silson_trace.append((ct['company'], ct.get('contract_date',''), '1세대 약값삭제', f'{_ykc}→0'))
                ws.cell(_ryk,col).value=None
            elif _ryk and not isinstance(_ykc,(int,float)):  # ① 별첨 약값 없을 때만 디폴트
                _ykd = 0 if _g4 else (10 if _life else 5)
                if _ykd: ws.cell(_ryk,col).value=_ykd; ws.cell(_ryk,col).font=BL   # 4세대 약0=미기재
            # ★ 실손 세대 자동판별 → 헤더에 라벨 기재
            _sg = silson_gen(ct.get('contract_date',''), _ipv, ct.get('product',''), _has_nonpay3(ct.get('dambo')))
            _pm0=re.search(r'(?<!\d)(0[9]|1[0-9]|2[0-6])(0[1-9]|1[0-2])(?!\d)', str(ct.get('product','')))
            silson_trace.append((ct['company'], ct.get('contract_date',''), (_pm0.group(0) if _pm0 else '없음'), _sg or '판정불가'))   # ★v29z 세대 근거
            if _sg:
                _hc = ws.cell(1,col)
                if _hc.value and _sg not in str(_hc.value):
                    _hc.value = str(_hc.value) + f'\n({_sg} 실손)'

    # ★v29t (지점장 확정 2026.07.02): CI 존재 시 '중대한CI적용' 행 = CI 잔여액 + 비CI 계약의 일반사망 동일액 —
    #   CI 적용/미적용 각각의 총 사망액이 양쪽 행에서 가로합산되도록.
    # ★★★v208 1인실 상급병원 가산(지점장 확정 2026.07.25, 영구):
    #   <b>1인실 상급병원 = 상급종합 1인실 + 종합 1인실</b>(상급종합에 입원하면 종합병원이하 담보도 함께 나온다).
    #   <b>1인실 종합병원 = 종합 1인실만</b>. 질병·상해는 택일 지급이라 담보 기재 단계에서 대표(max)로 잡힌다.
    #   실측(양*선 삼성 New내돈내삼): 상급 20 · 종합 10 → 상급행 30 / 종합행 10.
    #   ★계약 루프 <b>바깥</b>에서 1회만 돌린다(루프 안에 넣으면 계약 수만큼 중복 가산 — 실측 60 오류).
    _r1s = nm2r.get('1인실 상급병원'); _r1j = nm2r.get('1인실 종합병원')
    if _r1s and _r1j:
        for _c in range(3, 3 + n_ct):
            _vs = ws.cell(_r1s, _c).value; _vj = ws.cell(_r1j, _c).value
            if isinstance(_vs, (int, float)) and isinstance(_vj, (int, float)):
                ws.cell(_r1s, _c).value = _vs + _vj

    _rci_all=None; _ril_all=None; _rjb_all=None
    for _rr in range(6, ws.max_row+1):
        _b=str(ws.cell(_rr,2).value or '').strip()
        if _b=='중대한CI적용': _rci_all=_rr
        if _b=='일반사망': _ril_all=_rr
        if _b=='질병사망(80세)': _rjb_all=_rr
    _has_ci=any(_isci_prod(c.get('product')) for c in contracts)
    # ★★★★★v245 영구지침(지점장 확정 2026.07.25): <b>비CI 계약의 사망은 '질병사망(80세)' 행에 넣는다</b>.
    #   지점장 원문 = "(비CI 일반사망이 중대한CI적용에 찍히는 v29t 규칙) <b>그건 질병사망(80)에 넣어라</b>".
    #   구 v29t는 비CI 계약의 일반사망을 <b>'중대한CI적용'에 복사</b>했다(실측 메트라이프 6,000).
    #   → <b>'일반사망(종신)' 행은 CI 주계약 사망 전용</b>, 비CI 종신 사망은 질병사망(80세)로 옮긴다.
    #   ★한장보장표 질병사망 = 일반사망 + 질병사망(80세) 합이므로 <b>총액은 불변</b>(실측 22,000 유지).
    if _has_ci and _ril_all and _rjb_all:
        for _ix,_c in enumerate(contracts):
            _cl=3+_ix
            if _isci_prod(_c.get('product')): continue
            _ilv=ws.cell(_ril_all,_cl).value
            if isinstance(_ilv,(int,float)) and _ilv:
                _j0=ws.cell(_rjb_all,_cl).value
                ws.cell(_rjb_all,_cl).value=(_j0 if isinstance(_j0,(int,float)) else 0)+_ilv
                ws.cell(_rjb_all,_cl).font=ws.cell(_ril_all,_cl).font.copy()
                ws.cell(_ril_all,_cl).value=None
                if _rci_all: ws.cell(_rci_all,_cl).value=None   # ★구 v29t 복사분 제거
                print(f"[v245 비CI사망] {_c.get('company')} 일반사망 {_ilv:,} → 질병사망(80세) (중대한CI적용 복사 제거)")

    # ★ 합계 = 항상 표 맨 끝 열. 가로 SUM 수식(법칙22, 하드코딩 금지).
    last_col = 3 + n_ct
    # ★★★v230 (지점장 지시 2026.07.25, 영구): <b>유사암 자동유도(일반암×10%)는 완전 폐기</b>.
    #   지점장 원문 = <b>"유사암 적힌 것만 넣어라"</b>. 별첨에 유사암 담보가 없으면 <b>그 계약은 공란</b>이다 —
    #   일반암 금액으로 유추해 넣지 않는다. 구 v30q 자동유도가 없는 담보를 만들어냈다.
    #   실측(이정화): 한화 일반암 6,000 → 유사암 <b>600 자동생성</b> · 메리츠0804 1,000 → <b>100 자동생성</b>
    #   = 합계 1,450(정답 900). 자동유도 제거 + 명시 담보만 산입 → <b>900</b>으로 한장보장표와 일치.
    #   ★구 v213 '명시액이 하나도 없을 때만 유도' 게이트 방식도 함께 폐기(지점장 'no').

    first_L = get_column_letter(3)
    last_ct_L = get_column_letter(last_col-1) if n_ct>0 else first_L
    hc = ws.cell(1, last_col)
    hc.value = '합계'; hc.font = W; hc.fill = FILL_SUM; hc.alignment = AL
    # 보험료 합계 = 숫자만 표기(§3): 수식 아닌 계산된 숫자값. 글자 검정(흰바탕)
    # ★v128 지점장 확정 2026.07.21: 월보험료 합계 = 잔여보험료 > 0(납입 진행중) 계약만.
    #   납입완료(회차 a>=b) 계약은 계약열에는 남기되 합계에서는 뺀다.
    #   실측 근거(배학술 롯데): 12건 전액 711,218 → 교보 120/120·AIA 180/180 제외 → 605,618 = 리포트 일치.
    if n_ct>0:
        # ★★★v199 지점장 확정 2026.07.23: 보험료 합계도 '=D2+E2+…'가 아니라 <b>단일 =SUM()</b>.
        #   근거: 셀을 복사·이동하면 a+b+ 수식은 참조가 어긋나 합계가 하나도 안 맞는다.
        #   완납 계약은 보험료 칸을 텍스트로 넣어(위 루프) SUM이 자동으로 건너뛴다 = v129 정본 유지.
        #   ★캐시: inject_sum_cache가 숫자 셀만 합산해 같은 값을 박으므로 설명서·PPT가 0원으로 읽지 않는다.
        ws.cell(2, last_col).value = f'=SUM(C2:{last_ct_L}2)'
        ws.cell(2, last_col).font = BK

    for r in range(6, ws.max_row+1):
        slash_t=[0]*8; slash_n=0; is_slash=False; has_num=False   # ★v29v 1-8종·v29y 트리오: 실제 칸수 따름
        for col in range(3, last_col):
            v = ws.cell(r,col).value
            if isinstance(v,(int,float)): has_num=True
            elif isinstance(v,str) and '/' in v:
                is_slash = True
                _ps=v.split('/')[:8]
                slash_n=max(slash_n,len(_ps))
                for k,p in enumerate(_ps):
                    try: slash_t[k] += int(p)
                    except: pass
        sc = ws.cell(r, last_col)
        if is_slash and any(slash_t):
            sc.value = '/'.join(str(x) for x in slash_t[:(slash_n or 5)]); sc.font = BK   # 슬래시 행은 §3 SUM 예외
        else:
            # ★v29t: §5·v29c(2) 원복 — 합계는 동적 =SUM 수식(하드코딩 금지). 사용자가 값을 추가해도 자동 합산.
            #   저장 후 recalc_xlsx가 캐시값 주입 → 폰·미리보기에서도 숫자 표시(수식 유지).
            _rng = f'C{r}:{last_ct_L}{r}'
            _bnm=str(ws.cell(r,2).value).strip()
            # ★v91: 구 '입원 5,000 캡' 폐기(지점장 2026.07.19). 실손 다건이면 합이 5,000을 넘는다.
            #   실측 장혜경 = 현대 5,000 + DB 500 = 5,500 (한장보장표 질병 5,500과 일치).
            if _bnm=='입원': sc.value = f'=SUM({_rng})'
            elif _bnm=='자부상': sc.value = f'=MIN(SUM({_rng}),80)'          # ★지점장 2026.07.02: 자부상 최대 80만 캡
            elif _bnm=='120대수술비':                                       # ★v30k n대수술비=계약별 값 가로 슬래시(합산·최댓값 금지, 지점장 2026.07.03)
                _nd=[ws.cell(r,c).value for c in range(3,last_col)]
                _nd=[str(int(x)) for x in _nd if isinstance(x,(int,float)) and x>0]
                sc.value = '/'.join(_nd) if _nd else f'=SUM({_rng})'
            elif _bnm in ('간병인','중입자치료비'): sc.value = f'=IF(COUNT({_rng})=0,0,MAX({_rng}))'  # ★v30d 간병인·중입자=전 계약 대표 최댓값 1건
            elif _bnm=='간호통합병동': sc.value = f'=IF(COUNT({_rng})=0,0,MAX({_rng}))'   # ★v41 1-180일 최댓값 1건
            else: sc.value = f'=SUM({_rng})'
            sc.font = BK

    ws.column_dimensions['B'].width = 22
    for c in range(3, last_col+1):
        ws.column_dimensions[get_column_letter(c)].width = 12

    # ★ 테두리: A(구분)~끝열(합계) 전체 격자 직접 그림 + 구분(키워드)마다 굵은 구분선.
    #   (마스터 A·B 테두리가 중간행에서 끊겨 '선 없음' 발생 → 전부 새로 그림)
    _thin = Side(style='thin', color='FF000000'); _med = Side(style='medium', color='FF000000')   # ★v29t: 6자리 색은 알파00(투명) 저장돼 일부 뷰어에서 선 사라짐 → FF 필수
    # 구분(그룹) 끝행 동적 계산: A열에 값 있는 행=그룹 시작 → 다음 시작-1 = 그룹 끝
    g_starts = [r for r in range(6, ws.max_row+1) if ws.cell(r,1).value not in (None,'')]
    g_end = set()
    for k, s in enumerate(g_starts):
        e = (g_starts[k+1]-1) if k+1 < len(g_starts) else ws.max_row
        g_end.add(e)
    # 수술비 블록 내부 구분: 질병수술비 행 위에 굵은 선(상해 수술 ↔ 질병 수술)
    row_top_med = set()
    for r in range(6, ws.max_row+1):
        if str(ws.cell(r,2).value).strip() == '질병수술비':
            row_top_med.add(r)
    for r in range(1, ws.max_row+1):
        for c in range(1, last_col+1):
            left   = _med if c == 1 else _thin
            right  = _med if c == last_col else _thin
            top    = _med if (r in (1, 6) or r in row_top_med) else _thin
            # 헤더 5행 + 각 구분 끝행 = 굵은 가로 구분선
            bottom = _med if (r in (1,2,3,4,5) or r in g_end) else _thin
            ws.cell(r,c).border = Border(left=left, right=right, top=top, bottom=bottom)
    # ★ 숫자 콤마: 보험료·담보값·합계 SUM 전부 #,##0. (날짜·납입기간·슬래시 행은 텍스트라 제외)
    for r in range(1, ws.max_row+1):
        for c in range(2, last_col+1):
            v = ws.cell(r,c).value
            if isinstance(v,(int,float)) or (isinstance(v,str) and v.startswith('=')):
                ws.cell(r,c).number_format = '#,##0'

    # ★ 합계 이후 잔재 열 삭제 (§3: 합계 = 맨 끝 열)
    if ws.max_column > last_col:
        ws.delete_cols(last_col+1, ws.max_column - last_col)

    # ── 확인사항 시트: LLM 매핑 실패 담보 노출(자가진단, §10) ──
    for _sn in ('📋확인사항','확인사항'):
        if _sn in wb.sheetnames: del wb[_sn]
    ws2 = wb.create_sheet('확인사항')   # ★v41 이모지·외부하이퍼링크 제거(엑셀 '편집사용' 지연 원인)
    ws2.cell(1,1, f'{client} · 자동분석 {datetime.datetime.now():%Y.%m.%d}')
    ws2.cell(3,1,'계약수'); ws2.cell(3,2,n_ct)
    ws2.cell(4,1,'월보험료합계'); ws2.cell(4,2,f'{sum(c["premium"] for c in contracts):,}원')
    # ★★v186 (지점장 2026.07.22): AIA/AIG/라이나(우체국) 계약이 있으면 <b>엑셀 확인사항 시트 최상단</b>에
    #   "뇌 범위 부분 꼭 체크" 경고를 굵은 빨강으로 박는다. 3군데(엑셀·진단서 1p·7p 워크시트) 동일 문구.
    try:
        _wc = [f for f in ('AIA','AIG','라이나','우체국')
               if any(f in str(c.get('company','')).replace(' ','') for c in contracts)]
        if _wc:
            # ★v187 회사명은 <b>항상 3개 고정</b>(AIA / AIG / 라이나) — 보유분만 나열하지 않는다.
            _wcell = ws2.cell(2,1, '★ AIA / AIG / 라이나생명은 "뇌 범위 부분" 꼭 체크하기 '
                                   '(세부가입현황에서 뇌혈관 / 뇌졸증 / 뇌출혈 축 대조)')
            _wcell.font = Font(bold=True, size=13, color='C0392B')
    except Exception: pass
    ws2.cell(6,1,'[확인] 자동매핑 실패 담보 (마스터 미수록 또는 약관 확인 후 수기 기재)')
    ws2.cell(7,1,'회사'); ws2.cell(7,2,'담보명'); ws2.cell(7,3,'금액(만원)'); ws2.cell(7,4,'보장범위(참고)'); ws2.cell(7,5,'약관검색')
    LINKF = Font(color='0000FF', underline='single')
    # ★v148 (지점장 확정 2026.07.21): 흥국화재 10억통장(리셋월렛II)은 <b>엑셀 전면 기재금지</b>.
    #   본표에서 뺐어도 [확인] 큐(확인사항 시트)에 남으면 설계사가 수기 기재하게 되므로 여기서도 제외한다.
    #   이 담보는 보장진단서 7p 카드 전용이다.
    unmapped = [u for u in unmapped
                if not (('리셋월렛' in re.sub(r'\s','',str(u[2]))) or ('리셋월랫' in re.sub(r'\s','',str(u[2]))))]
    for k,(col,comp,raw,amt,note) in enumerate(unmapped):
        rr = 8+k
        ws2.cell(rr,1,comp); ws2.cell(rr,2,raw); ws2.cell(rr,3,amt); ws2.cell(rr,4,note)
        prod = contracts[col-3]['product'] if 0<=col-3<len(contracts) else ''
        prod_key = re.sub(r'[\(\)\[\]ⅠⅡⅢ_]', ' ', prod)[:18].strip()
        q = f"{comp} {prod_key} {raw[:12]} 약관 보장내용"
        # ★v41 hyperlink 객체 금지 → 평문 URL(엑셀이 열 때 외부링크 검증 안 함 = 편집사용 즉시)
        ws2.cell(rr,5, "https://search.naver.com/search.naver?query=" + urllib.parse.quote(q))
    ws2.column_dimensions['B'].width = 34; ws2.column_dimensions['D'].width = 40; ws2.column_dimensions['E'].width = 12
    # ★v29z 근거 감사 로그 — '없는 값' 논쟁 즉시 검증용
    _rr = 9 + len(unmapped)
    if silson_trace:
        _rr += 2; ws2.cell(_rr,1,'[근거] 실손 세대 판정 (가입일 vs 상품코드 — 상품코드 우선)')
        _rr += 1; ws2.cell(_rr,1,'회사'); ws2.cell(_rr,2,'가입일'); ws2.cell(_rr,3,'상품코드(YYMM)'); ws2.cell(_rr,4,'판정')
        for (_c,_d,_p,_g) in silson_trace:
            _rr += 1; ws2.cell(_rr,1,_c); ws2.cell(_rr,2,_d); ws2.cell(_rr,3,_p); ws2.cell(_rr,4,_g)
    if heart_trace:
        _rr += 2; ws2.cell(_rr,1,'[근거] 심장 블록 기재 내역 (원 담보명 → 기재 행) — 별첨 원문 그대로')
        _rr += 1; ws2.cell(_rr,1,'회사'); ws2.cell(_rr,2,'별첨 원 담보명'); ws2.cell(_rr,3,'기재 행'); ws2.cell(_rr,4,'금액(만원)')
        for (_c,_raw,_rows,_a) in heart_trace:
            _rr += 1; ws2.cell(_rr,1,_c); ws2.cell(_rr,2,str(_raw)[:60]); ws2.cell(_rr,3,_rows); ws2.cell(_rr,4,_a)
    if surg_trace:   # ★v30g 수술 블록 근거 — 종수술 슬롯 이상치 즉시 추적용
        _rr += 2; ws2.cell(_rr,1,'[근거] 수술 블록 기재 내역 (원 담보명 → 기재 행/슬롯) — 별첨 원문 그대로')
        _rr += 1; ws2.cell(_rr,1,'회사'); ws2.cell(_rr,2,'별첨 원 담보명'); ws2.cell(_rr,3,'기재 행/슬롯'); ws2.cell(_rr,4,'금액(만원)')
        for (_c,_raw,_rows,_a) in surg_trace:
            _rr += 1; ws2.cell(_rr,1,_c); ws2.cell(_rr,2,str(_raw)[:60]); ws2.cell(_rr,3,_rows); ws2.cell(_rr,4,_a)
    # ★★★★★ v238 CI 진단표 — <b>CI 계약이 하나라도 있으면 무조건 출력</b>(영구).
    #   오늘(2026.07.25) 사고: CI 2건이 통째로 누락됐는데 <b>산출물 어디에도 흔적이 없어</b> 지점장이
    #   직접 발견해야 했다(확인사항에 CI 근거표가 없었다). → 이제 CI 계약·사망액·본체·선지급률·
    #   중대한OO 배치 결과를 <b>표로 항상 남긴다</b>. 배치가 0이면 최상단에 [경고] 행을 찍는다.
    _cid = _ci_diag or []
    if _cid:
        _zero = all((d.get('placed') or 0) == 0 for d in _cid)
        _rr += 2
        ws2.cell(_rr,1, '[경고] CI 계약이 있으나 중대한OO 배치가 0건 — 선지급 분해 실패, 즉시 확인' if _zero
                       else '[근거] CI 판정·선지급률 내역 (CI 계약은 항상 이 표를 확인할 것)')
        if _zero:
            try: ws2.cell(_rr,1).font = Font(bold=True, color='C00000')
            except Exception: pass
        _rr += 1
        for _i,_h in enumerate(('회사','상품명','①세부내역','②별첨','판정근거','사망보장(만원)','선지급률','본체(만원)','중대한OO 배치'),1):
            ws2.cell(_rr,_i,_h)
        for d in _cid:
            _rr += 1
            ws2.cell(_rr,1,d.get('co','')); ws2.cell(_rr,2,str(d.get('pd',''))[:40])
            ws2.cell(_rr,3,'O' if d.get('sebu') else 'X')
            ws2.cell(_rr,4,'O' if d.get('byul') else 'X')
            ws2.cell(_rr,5,d.get('src',''))
            ws2.cell(_rr,6,d.get('samang') or 0)
            ws2.cell(_rr,7,(f"{d['pct']}%형" if d.get('pct') else '[확인] 50%/80% 미해당'))
            ws2.cell(_rr,8,d.get('bonche') or 0)
            ws2.cell(_rr,9,d.get('placed_txt') or '없음 — [확인]')
        _rr += 1
        ws2.cell(_rr,1,'※ CI 4단계 체크: ①상세 세부내역 ②[별첨] 보장현황 ③중대한*** 유무 ④동일금액 2개 이상 → 50%면 ×2 / 80%면 ÷0.8 이 사망보장금과 맞는지 검증. 선지급률은 50%·80% 두 가지뿐 — 추측 금지.')
    if cancer_trace:   # ★v30h 암 블록 근거 — 일반암 과다·통합암 중복 즉시 추적
        _rr += 2; ws2.cell(_rr,1,'[근거] 암 블록 기재 내역 (원 담보명 → 기재 행) — 별첨 원문 그대로')
        _rr += 1; ws2.cell(_rr,1,'회사'); ws2.cell(_rr,2,'별첨 원 담보명'); ws2.cell(_rr,3,'기재 행'); ws2.cell(_rr,4,'금액(만원)')
        for (_c,_raw,_rows,_a) in cancer_trace:
            _rr += 1; ws2.cell(_rr,1,_c); ws2.cell(_rr,2,str(_raw)[:60]); ws2.cell(_rr,3,_rows); ws2.cell(_rr,4,_a)
    # ★v39 워크시트 담보명 카피: 원본담보명을 숨김 시트 _dambo_raw 에 저장 (등식·기존시트 무손상)
    try:
        if '_dambo_raw' in wb.sheetnames: del wb['_dambo_raw']
        _rs = wb.create_sheet('_dambo_raw'); _rs.sheet_state='hidden'
        _rs.cell(1,1,'std'); _rs.cell(1,2,'raw'); _rs.cell(1,3,'amt')
        for _i,(_std,(_rw,_am)) in enumerate(raw_by_std.items(), start=2):
            _rs.cell(_i,1,_std); _rs.cell(_i,2,_rw); _rs.cell(_i,3,_am)
    except Exception:
        pass
    _no_fullcalc(wb)          # ★v51 편집모드 강제 재계산 방지(수식은 유지)
    wb.save(out)
    _force_nocalc_xml(out)    # ★v124 저장 후 XML에 직접 못박음(3중 방어)
    return unmapped


def _force_nocalc_xml(path):
    """★★★v124(2026.07.21) 편집모드 속도 — 3중 방어의 마지막 단계.
       세 번 재발한 이유: 원천 master.xlsx가 fullCalcOnLoad="1"이라 코드가 매번 덮는 구조였다.
       v124에서 ①master.xlsx 자체를 "0"으로 교정 ②_no_fullcalc 유지 ③저장된 XML에 직접 못박음.
       셋 중 둘이 사라져도 폰 Excel '편집 사용'이 느려지지 않는다. 절대 제거 금지."""
    import zipfile, shutil, tempfile, os as _os
    try:
        zin = zipfile.ZipFile(path, 'r')
        tmp = path + '.nc'
        zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'xl/workbook.xml':
                x = data.decode('utf-8')
                if 'fullCalcOnLoad' in x:
                    x = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="0"', x)
                elif '<calcPr' in x:
                    x = re.sub(r'(<calcPr\b)', r'\1 fullCalcOnLoad="0"', x, count=1)
                else:
                    x = x.replace('</workbook>', '<calcPr fullCalcOnLoad="0"/></workbook>')
                data = x.encode('utf-8')
            zout.writestr(it, data)
        zin.close(); zout.close()
        shutil.move(tmp, path)
        return True
    except Exception:
        try:
            if _os.path.exists(path + '.nc'): _os.unlink(path + '.nc')
        except Exception:
            pass
        return False

def read_excel_totals(path):
    """완성 엑셀에서 담보명->합계 읽음. 등식2: PPT는 이것만 본다.
       ★★★v218 (지점장 지시 2026.07.25, 영구): <b>PPT 합계는 엑셀 끝열과 '반드시' 같아야 한다</b>.
       구 코드는 캐시가 없으면 <b>무조건 데이터셀 단순 SUM</b>으로 폴백했는데, 엑셀 끝열은 행마다
       수식이 다르다 → <b>대표(max)·캡 담보가 전부 어긋났다</b>(실측 불일치 5건):
         간병인 MAX 20 → PPT 30 / 간호통합병동 MAX 7 → PPT 12 / 중입자 MAX 300 → PPT 400 /
         자부상 MIN(,80) 80 → PPT 100 / 120대수술비 '1000/500' → PPT 1500.
       → <b>폴백도 끝열 수식 종류를 그대로 따른다</b>(inject_sum_cache와 동일 규칙).
       ★<b>동명 담보 행이 2개 이상</b>(혈전용해치료비 = 뇌 37행·심장 50행)이면 dict 키가 충돌해
       <b>뒤엣것이 앞엣것을 덮어쓴다</b> → PPT엔 한 칸뿐이므로 <b>두 행 중 대표(max)</b>로 병합한다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path)            # ★v218 수식 원문 판독용(data_only=False)
    ws = wb['보장분석']; wsf = wbf['보장분석']; last = ws.max_column
    out = {}; sq=[0]*5; ss=[0]*5; splits={}   # ★v219 splits[담보]=(갱신합, 비갱신합) — 엑셀 글자색 근거

    def _fallback(r):
        """끝열 캐시가 없을 때 — 엑셀 끝열 수식과 <b>같은 규칙</b>으로 계산한다."""
        nums = [ws.cell(r,c).value for c in range(3,last) if isinstance(ws.cell(r,c).value,(int,float))]
        s = sum(nums)
        f = wsf.cell(r,last).value
        if isinstance(f,str) and f.startswith('='):
            if f.startswith('=MIN('):
                mm = re.search(r',\s*(\d+)\s*\)\s*$', f)
                return min(s, int(mm.group(1))) if mm else s
            if f.startswith('=IF(COUNT'):   return max(nums) if nums else 0   # 대표(max) 담보
            if f.startswith('=IF(SUM'):     return 7 if s>0 else 0
        return s

    def _put(nm, val):
        """★v218 동명 행 충돌 방지 — 이미 있으면 큰 값 유지(대표)."""
        if not val: return
        prev = out.get(nm)
        out[nm] = max(prev, val) if isinstance(prev,(int,float)) else val

    for r in range(6, ws.max_row+1):
        nm = ws.cell(r,2).value
        if not nm: continue
        nm = str(nm).strip()
        endv = ws.cell(r,last).value
        # ★★★v219 (지점장 지시 2026.07.25, 영구): <b>PPT 갱신/비갱신 색은 엑셀 글자색이 원천이다</b>.
        #   구 코드는 PPT가 raw dambo로 갱신합/비갱신합을 <b>따로 다시 계산</b>했다. 그런데 엑셀 끝열은
        #   대표(max)·캡·[확인] 제외 때문에 raw 분할합과 <b>거의 항상 다르다</b> → 구 코드는 그럴 때
        #   <b>'큰 쪽으로 전부 몰아버려'</b>(`if gs>=ns: gs,ns=_T,0`) <b>갱신↔비갱신이 통째로 뒤바뀌었다</b>.
        #   실측 뒤바뀜: 갱신 100(엑셀 산입) + 비갱신 200([확인] 제외) → ns가 커서 <b>검정</b>으로 찍히지만
        #   엑셀의 100은 <b>갱신 계약 값이라 파랑</b>이어야 한다. 역방향도 동일하게 발생.
        #   → <b>엑셀 데이터셀의 글자색(파랑 0070C0 = 갱신)을 그대로 읽어</b> 분할한다. 이중 계산 폐기.
        _cells=[]
        for c in range(3, last):
            _v = ws.cell(r,c).value
            if not isinstance(_v,(int,float)) or not _v: continue
            try: _rgb = str(wsf.cell(r,c).font.color.rgb or '')
            except Exception: _rgb = ''
            _cells.append((_v, _rgb.upper().endswith('0070C0')))
        _f0 = wsf.cell(r,last).value
        if isinstance(_f0,str) and _f0.startswith('=IF(COUNT') and _cells:
            # ★대표(max) 행 — 끝열 값을 만든 <b>최댓값 셀 하나의 색</b>이 정답이다.
            #   합으로 나누면 갱신 8+9(=17)가 비갱신 15를 눌러 <b>색이 뒤집힌다</b>(구 코드 실패 지점).
            _mx = max(_cells, key=lambda x: x[0])
            _gs, _ns = (_mx[0], 0) if _mx[1] else (0, _mx[0])
        else:
            _gs = sum(v for v,b in _cells if b)
            _ns = sum(v for v,b in _cells if not b)
        if _gs or _ns:
            _pv0 = splits.get(nm,(0,0))
            splits[nm] = (max(_pv0[0], _gs), max(_pv0[1], _ns))
        # 수술비 1~5종: 끝열 슬래시 문자열(수식 아님, 항상 존재)
        if nm == '상해 종수술비(1-5종)' and isinstance(endv,str) and '/' in endv:
            for k,p in enumerate(endv.split('/')[:5]):
                try: ss[k]=int(p)
                except: pass
            continue
        if nm == '질병 종수술비(1-5종)' and isinstance(endv,str) and '/' in endv:
            for k,p in enumerate(endv.split('/')[:5]):
                try: sq[k]=int(p)
                except: pass
            continue
        if nm == 'MRI/도수치료/비급여주사' and isinstance(endv,str) and '/' in endv:   # ★v29y 트리오 분해
            _ps=endv.split('/')
            for _k,_std in enumerate(('MRI','도수치료','비급여주사')):
                try: out[_std]=int(_ps[_k])
                except: pass
            continue
        # ★★v218 120대수술비(n대수술비) = 끝열이 <b>계약별 가로 슬래시 문자열</b>이다(합산 금지, v30k 정본).
        #   구 코드는 이 문자열을 못 읽어 폴백 단순합산 1,500을 PPT에 넣었다(엑셀 '1000/500'과 불일치).
        #   → 슬래시를 쪼개 <b>대표(max)</b>를 PPT 값으로 쓴다.
        if isinstance(endv,str) and '/' in endv:
            _pv=[]
            for _p in endv.split('/'):
                try: _pv.append(int(str(_p).strip().replace(',','')))
                except: pass
            if _pv: _put(nm, max(_pv))
            continue
        # 숫자 합계: 끝열 캐시값 있으면 사용, 없으면 끝열 수식과 동일 규칙으로 계산(★v218)
        _put(nm, endv if isinstance(endv,(int,float)) and endv else _fallback(r))
    return out, sq, ss, splits

def build_ppt(data, out, totals=None, surg_q=None, surg_s=None, splits=None):
    if not os.path.exists(TPL_PPT): return False
    prs = Presentation(TPL_PPT)
    sl = prs.slides[0]
    by = {sh.name:sh for sh in sl.shapes if sh.has_text_frame}
    client = data['client']; contracts = data['contracts']
    now = datetime.datetime.now()

    totals = totals if totals is not None else {}
    surg_q = surg_q if surg_q is not None else [0]*5
    surg_s = surg_s if surg_s is not None else [0]*5
    if not totals:   # 폴백: 전달 없으면 옛 방식(등식2 불가시)
        for ct in contracts:
            for raw,amt in ct['dambo'].items():
                std = resolve(raw)
                if std: totals[std]=totals.get(std,0)+amt

    # ★PPT 색: 하나라도 갱신=파랑 / 전부 비갱신=검정 / 실손 항상 파랑 (미가입은 값 미기재라 해당없음)
    _BLUE=RGBColor(0x00,0x00,0xFF); _BLACK=RGBColor(0x00,0x00,0x00)
    _silson={'입원','통원','약값','약','MRI','도수치료','비급여주사','일상배상책임'}  # ★v210 간병인·간호통합병동 강제 파랑 폐기(엑셀과 동일 규칙) — 구 v139 3행 무조건 파랑 폐기
    # 담보별 '최대 기여 계약'의 갱신여부로 색 결정 → 합산 시 전부 파랑 쏠림 방지(엑셀 혼합과 일치)
    _dom={}  # std -> (max_amt, gen)
    for ct in contracts:
        _gen = (ct.get('renewal','')=='갱신')
        for raw,amt in ct.get('dambo',{}).items():
            if not amt: continue
            st=resolve(raw)
            if not st: continue
            if st not in _dom or amt>_dom[st][0]: _dom[st]=(amt,_gen)
    def pcol(std):
        if std in _silson: return _BLUE
        d=_dom.get(std)
        return _BLUE if (d and d[1]) else _BLACK
    # ★담보별 갱신합/비갱신합 (분할 표기용)
    # ★★★v219 (지점장 지시 2026.07.25, 영구): <b>갱신합/비갱신합은 엑셀 글자색에서 그대로 가져온다</b>.
    #   구 코드는 raw dambo로 재계산해 엑셀 끝열과 어긋났고, 어긋나면 큰 쪽으로 몰아버려
    #   <b>갱신↔비갱신이 통째로 뒤바뀌었다</b>. splits=read_excel_totals가 준 (갱신합, 비갱신합).
    _gensum={}; _nonsum={}
    if splits:
        for _st,(_g,_n) in splits.items():
            if _g: _gensum[_st]=_g
            if _n: _nonsum[_st]=_n
    else:   # 폴백(엑셀 없이 호출된 경우) — 구 방식
        for ct in contracts:
            _gen=(ct.get('renewal','')=='갱신')
            for raw,amt in ct.get('dambo',{}).items():
                if not amt: continue
                st=resolve(raw)
                if not st: continue
                tgt=_gensum if (_gen or '갱신' in raw) else _nonsum
                tgt[st]=tgt.get(st,0)+amt
    # ★v30n PPT 골절 = 골절(치아파절포함) + 골절(치아파절제외) 합산. 엑셀은 두 행 분리 유지, PPT만 하나로 합산 표기(지점장 2026.07.03)
    for _b in (_gensum, _nonsum, totals):
        _b['골절합산PPT'] = (_b.get('골절(치아파절포함)',0) or 0) + (_b.get('골절(치아파절제외)',0) or 0)
    def _seg(run0, segs):
        run0.text=segs[0][0]
        if segs[0][1] is not None:
            try: run0.font.color.rgb=segs[0][1]
            except: pass
        prev=run0._r
        for txt,col in segs[1:]:
            new=_copy.deepcopy(run0._r)
            t=new.find(_qn('a:t'))
            if t is not None: t.text=txt
            prev.addnext(new); prev=new
            nr=_Run(new, run0._parent)
            if col is not None:
                try: nr.font.color.rgb=col
                except: pass
    def pv(box,pi,ri,std,prefix=': ',suffix=''):
        # 한 칸 분할: 갱신합=파랑 / 비갱신합=검정 (둘 다면 'gen / non'), 실손=파랑 합계
        # ★★★v221: 구 코드는 박스·문단·run을 못 찾으면 <b>아무 말 없이 return</b>했다.
        #   그래서 'PPT 상해 수술 라인이 통째로 비어 있다' 같은 사고가 <b>로그 한 줄 없이</b> 나갔다.
        #   → 실패 사유를 반드시 찍는다(배포 후 Railway 로그로 즉시 원인 확정 가능).
        if box not in by:
            print(f'[PPT_MISS] 박스없음 box={box} std={std} — 템플릿 ppt_form.pptx의 도형 이름 확인 필요'); return
        tf=by[box].text_frame
        if pi>=len(tf.paragraphs):
            print(f'[PPT_MISS] 문단없음 box={box} p{pi} (문단수 {len(tf.paragraphs)}) std={std}'); return
        p=tf.paragraphs[pi]
        if ri>=len(p.runs):
            print(f'[PPT_MISS] run없음 box={box} p{pi} r{ri} (run수 {len(p.runs)}) std={std}'); return
        gs=_gensum.get(std,0); ns=_nonsum.get(std,0)
        if std in _silson:
            _v = totals.get(std,0)            # ★실손=완성 엑셀값(입원5천캡·통원디폴트 반영). _gensum 원본합산(상해+질병=1만) 사용 안 함
            if not _v: return
            segs=[(f'{prefix}{_v:,}{suffix}', _BLUE)]
            _seg(p.runs[ri], segs); return
        # ★v30f 등식1 (지점장 승인 2026.07.03): PPT 표기 총액의 정본 = 완성 엑셀 끝열.
        #   대표 1건·MAX·캡·[확인] 분리로 끝열 ≠ raw 분할합이면 끝열 값 단색 표기(색=우세 성분측).
        # ★★v219: gs·ns가 이미 <b>엑셀 셀에서 온 값</b>이므로 원칙적으로 gs+ns == 끝열이다.
        #   다만 끝열이 대표(max)·캡 수식인 행은 합보다 작다 → <b>큰 쪽으로 몰지 말고</b>
        #   <b>어느 쪽 성분이 끝열 값을 만들었는지</b>로 판정한다(색 뒤바뀜 차단).
        _T = totals.get(std, None)
        if isinstance(_T,(int,float)) and _T>0 and (gs+ns)!=_T:
            if   gs and not ns: gs,ns=int(_T),0
            elif ns and not gs: gs,ns=0,int(_T)
            elif gs>=_T and ns< _T: gs,ns=int(_T),0      # 끝열값을 만든 쪽 = 갱신
            elif ns>=_T and gs< _T: gs,ns=0,int(_T)      # 끝열값을 만든 쪽 = 비갱신
            else:                                        # 둘 다 기여(=SUM 계열) → 비율 보존
                _tot=gs+ns
                gs=int(round(_T*gs/_tot)); ns=int(_T)-gs
        if not gs and not ns: return
        if gs and ns:
            segs=[(f'{prefix}{gs:,}', _BLUE),(f'+{ns:,}{suffix}', _BLACK)]
        elif gs:
            segs=[(f'{prefix}{gs:,}{suffix}', _BLUE)]
        else:
            segs=[(f'{prefix}{ns:,}{suffix}', _BLACK)]
        _seg(p.runs[ri], segs)
    def _setcol(run,std):
        try: run.font.color.rgb=pcol(std)
        except: pass
    _last_std=[None]
    def g(nm):
        _last_std[0]=nm
        return totals.get(nm,0)
    def r_set(box,pi,ri,val,std='__USE_LAST__'):
        use = _last_std[0] if std=='__USE_LAST__' else std
        # ★v221 조용한 실패 금지 — 못 찍었으면 사유를 로그로 남긴다.
        if box not in by:
            print(f'[PPT_MISS] 박스없음 box={box} val={val}')
        else:
            tf=by[box].text_frame
            if pi>=len(tf.paragraphs):
                print(f'[PPT_MISS] 문단없음 box={box} p{pi} (문단수 {len(tf.paragraphs)}) val={val}')
            else:
                p=tf.paragraphs[pi]
                if ri>=len(p.runs):
                    print(f'[PPT_MISS] run없음 box={box} p{pi} r{ri} (run수 {len(p.runs)}) val={val}')
                else:
                    p.runs[ri].text=val
                    if use: _setcol(p.runs[ri], use)
        _last_std[0]=None

    for b in ['TextBox 49','TextBox 56']:
        if b in by: by[b].text_frame.word_wrap=False

    by['TextBox 21'].text_frame.word_wrap=False
    by['TextBox 21'].text_frame.auto_size=MSO_AUTO_SIZE.NONE  # 도형 고정(이름 길이에 따라 박스 이동·크기변경 방지)
    by['TextBox 21'].text_frame.paragraphs[0].runs[0].text=f'{client} 님의 보장'
    by['TextBox 21'].text_frame.paragraphs[0].runs[1].text='(전)'
    # 날짜를 한 박스(TextBox 36)로 통합, 35·29는 비움
    if 'TextBox 36' in by and 'TextBox 29' in by:
        try: by['TextBox 36'].width = by['TextBox 29'].left + by['TextBox 29'].width - by['TextBox 36'].left
        except: pass
    by['TextBox 36'].text_frame.paragraphs[0].runs[0].text=f'{now.year}년 {now.month:02d}월 {now.day:02d}일 기준'
    for _eb in ('TextBox 35','TextBox 29'):
        if _eb in by:
            for _pp in by[_eb].text_frame.paragraphs:
                for _rr in _pp.runs: _rr.text=''
    # 상단 헤더(이름+날짜) 전부 18pt·도형 고정
    for _hb in ('TextBox 21','TextBox 36','TextBox 35','TextBox 29'):
        if _hb in by:
            by[_hb].text_frame.word_wrap=False
            try: by[_hb].text_frame.auto_size=MSO_AUTO_SIZE.NONE
            except: pass
            for _pp in by[_hb].text_frame.paragraphs:
                try: _pp.alignment = PP_ALIGN.CENTER      # ★v41 이름·날짜 우측쏠림 → 가운데
                except: pass
                for _rr in _pp.runs:
                    try: _rr.font.size=Pt(18)
                    except: pass

    # ★v48(지점장 2026.07.13): 제목(고객명)+날짜를 한 덩어리로 슬라이드 가운데 배치
    try:
        _t, _d = by.get('TextBox 21'), by.get('TextBox 36')
        if _t is not None and _d is not None:
            _SW = prs.slide_width
            _GAP = 100000
            _tot = _t.width + _GAP + _d.width
            _lf = int((_SW - _tot) / 2)
            _t.left = _lf
            _d.left = _lf + _t.width + _GAP
            _d.top = _t.top
    except Exception:
        pass

    if g('질병사망(80세)'): pv('TextBox 10',2,2,'질병사망(80세)',prefix=': ',suffix='')
    if g('상해사망'): pv('TextBox 11',0,1,'상해사망',prefix=': ',suffix='')
    종신_d=0
    for ct in contracts:
        if '종신' in ct['renewal']:
            for raw,v in ct['dambo'].items():
                if resolve(raw)=='상해사망': 종신_d+=v
    if 종신_d: r_set('TextBox 10',3,1,f': {종신_d:,}')

    if g('상해후유3%'): pv('TextBox 8',2,1,'상해후유3%',prefix='3% : ',suffix='')
    if g('질병후유3%'): pv('TextBox 8',0,1,'질병후유3%',prefix='3% : ',suffix='')
    if g('상해후유80%'): pv('TextBox 8',3,1,'상해후유80%',prefix='80% : ',suffix='')
    if g('질병후유80%'): pv('TextBox 8',1,1,'질병후유80%',prefix='80% : ',suffix='')

    if g('뇌혈관진단비'): pv('TextBox 46',0,0,'뇌혈관진단비',prefix='뇌혈관\n',suffix='')
    if g('뇌졸증진단비'): pv('TextBox 47',0,0,'뇌졸증진단비',prefix='뇌졸증\n',suffix='')
    if g('뇌출혈진단비'): pv('TextBox 48',0,0,'뇌출혈진단비',prefix='뇌출혈\n',suffix='')
    if g('산정특례뇌혈관'): pv('TextBox 49',0,3,'산정특례뇌혈관',prefix=': ',suffix='')
    if g('혈전용해치료비'): pv('TextBox 49',1,1,'혈전용해치료비',prefix=': ',suffix='')
    if g('2대 주요치료비'): pv('TextBox 49',2,2,'2대 주요치료비',prefix=': ',suffix='')   # 뇌혈관쪽 2대주요치료비

    # ★ 심장 표기(설명서와 동일 8종): 1줄 협심증/심부전/염증/빈맥 · 2줄 부정맥/심근병증/심장판막. 값 있는 것만. 급성심근·허혈성 별도칸.
    if 'TextBox 심장4종' in by:
        _h4=by['TextBox 심장4종'].text_frame
        _hp=totals.get('협심증',0); _sf=totals.get('심부전',0); _ym=totals.get('염증',0); _bm=totals.get('빈맥',0)
        _bj=totals.get('부정맥',0); _mbz=totals.get('심근병증',0); _pmz=totals.get('심장판막',0)
        _names=[n for n,v in [('협심증',_hp),('심부전',_sf),('염증',_ym),('빈맥',_bm)] if v]   # ★지점장 2026.07.05 빈맥 복원(40행 정식)
        _amt=max(_hp,_sf,_ym,_bm)
        if _names and len(_h4.paragraphs[0].runs)>=2:
            _h4.paragraphs[0].runs[0].text='/'.join(_names)+' '
            _h4.paragraphs[0].runs[1].text=f'{_amt:,}' if _amt else ''
        elif len(_h4.paragraphs[0].runs)>=1:
            _h4.paragraphs[0].runs[0].text=''
        # 2줄 = 부정맥·심근병증·심장판막(★지점장 2026.07.05: 설명서와 동일하게 심근병증·판막 추가)
        _names2=[n for n,v in [('부정맥',_bj),('심근병증',_mbz),('심장판막',_pmz)] if v]
        _amt2=max(_bj,_mbz,_pmz)
        if len(_h4.paragraphs)>1 and len(_h4.paragraphs[1].runs)>=2:
            if _names2:
                _h4.paragraphs[1].runs[0].text='/'.join(_names2)+' '
                _h4.paragraphs[1].runs[1].text=f'{_amt2:,}' if _amt2 else ''
            else:
                _h4.paragraphs[1].runs[0].text=''
                _h4.paragraphs[1].runs[1].text=''
    # ★허혈성 진단비 값 채움(TextBox 54) — 설명서와 동일하게(★지점장 2026.07.05, 기존 미채움 버그 수정)
    if 'TextBox 54' in by:
        _hv=totals.get('허혈성 진단비',0)
        _t54=by['TextBox 54'].text_frame
        if _t54.paragraphs[0].runs:
            _t54.paragraphs[0].runs[0].text = (f'허혈성 : {_hv:,}' if _hv else '허혈성')
    if g('급성심근경색'): pv('TextBox 55',0,0,'급성심근경색',prefix='급성심근\n',suffix='')
    if g('산정특례심장'): pv('TextBox 56',0,3,'산정특례심장',prefix=': ',suffix='')
    if g('2대 주요치료비'): pv('TextBox 56',2,2,'2대 주요치료비',prefix=': ',suffix='')   # 심장쪽 2대주요치료비

    if g('일반암'): pv('TextBox 14',0,1,'일반암',prefix=': ',suffix='')
    if g('유사암(갑.기.경.제)'): pv('TextBox 14',1,2,'유사암(갑.기.경.제)',prefix=': ',suffix='')
    if g('항암방사선약물'): pv('TextBox 14',4,1,'항암방사선약물',prefix=': ',suffix=' / ')
    if g('표적항암치료비'): pv('TextBox 14',5,1,'표적항암치료비',prefix=': ',suffix=' / ')
    if g('세기조절치료'): pv('TextBox 14',5,4,'세기조절치료',prefix=': ',suffix='')
    if g('양성자치료'): pv('TextBox 14',5,5,'양성자치료',prefix=': ',suffix='')
    if g('다빈치로봇수술비'): pv('TextBox 14',7,1,'다빈치로봇수술비',prefix=': ',suffix='')
    # 상급병원 암주요치료비 / 하이클래스 (TextBox 57)
    if 'TextBox 57' in by: by['TextBox 57'].text_frame.word_wrap=False
    if g('암주요치료비'): pv('TextBox 57',0,2,'암주요치료비',prefix=': ',suffix='')
    if g('하이클래스(암)'): pv('TextBox 57',1,2,'하이클래스(암)',prefix=': ',suffix='')

    if g('질병수술비'): pv('TextBox 17',0,1,'질병수술비',prefix=': ',suffix='')
    if any(surg_q): r_set('TextBox 17',3,0,f'({"/".join(str(x) for x in surg_q)})','질병 종수술비(1-5종)'); r_set('TextBox 17',3,2,'',None)
    if g('뇌혈관수술비'): pv('TextBox 17',5,1,'뇌혈관수술비',prefix=': ',suffix='')
    if g('심장수술비'): pv('TextBox 17',7,1,'심장수술비',prefix=': ',suffix='')
    if g('상해수술비'): pv('TextBox 19',0,1,'상해수술비',prefix=': ',suffix='')
    if any(surg_s): r_set('TextBox 19',3,0,f'({"/".join(str(x) for x in surg_s)})','상해 종수술비(1-5종)'); r_set('TextBox 19',3,2,'',None)
    if g('골절수술비'): pv('TextBox 19',4,1,'골절수술비',prefix=': ',suffix='')

    _ys=totals.get('양성자치료',0); _sgj=totals.get('세기조절치료',0)   # ★v29v (지점장 2026.07.02) 양성자·세기조절 → 암 박스
    if _ys or _sgj:
        try:
            _p14=by['TextBox 14'].text_frame.paragraphs[5]
            _t = (f'{_ys:,}/{_sgj:,}' if (_ys and _sgj) else f'{(_ys or _sgj):,}')
            _p14.runs[-1].text=': '+_t
        except Exception: pass
    실손_cts=[ct for ct in contracts
        if any('실손' in k or '입원의료비' in k for k in ct['dambo']) and ct['contract_date']]
    실손가입일=min((c['contract_date'] for c in 실손_cts), default='___________')
    _실손상품=next((c.get('product','') for c in 실손_cts if c['contract_date']==실손가입일), '')
    _np3=any(_has_nonpay3(c.get('dambo')) for c in 실손_cts)   # ★v250 3대비급여 하한
    _sg=silson_gen(실손가입일, totals.get('입원'), _실손상품, _np3)   # ★실손 세대 자동판별(상품명 연도코드 반영)
    by['TextBox 59'].text_frame.word_wrap=False
    by['TextBox 59'].text_frame.paragraphs[0].runs[0].text='실손'+(f' {_sg}' if _sg else '')
    by['TextBox 59'].text_frame.paragraphs[1].runs[0].text='('
    by['TextBox 59'].text_frame.paragraphs[1].runs[1].text='가입일:'
    by['TextBox 59'].text_frame.paragraphs[1].runs[2].text=f'{실손가입일})'
    for r in by['TextBox 59'].text_frame.paragraphs[1].runs: r.font.size=Pt(10)  # ★v50 '다10'
    if g('입원'): pv('TextBox 6',0,1,'입원',prefix=': ',suffix='')
    if g('통원'): pv('TextBox 6',1,1,'통원',prefix=': ',suffix=' / ')
    if g('약값'): pv('TextBox 6',1,3,'약값',prefix=': ',suffix='')   # ★v29t 등식1: 약값 PPT 누락 수리
    if g('MRI'): pv('TextBox 6',2,0,'MRI',prefix='MRI : ',suffix='')
    if g('도수치료'): pv('TextBox 6',3,1,'도수치료',prefix=': ',suffix='')
    if g('비급여주사'): pv('TextBox 6',4,1,'비급여주사',prefix=': ',suffix='')

    if g('골절합산PPT'): pv('TextBox 7',0,1,'골절합산PPT',prefix=': ',suffix='')   # ★v30n 엑셀 골절 두 행 합산 표기
    if g('화상진단비'): pv('TextBox 7',2,1,'화상진단비',prefix=': ',suffix='')
    if g('깁스진단비'): pv('TextBox 7',5,1,'깁스진단비',prefix=': ',suffix='')
    if g('응급실(응급)'): pv('TextBox 7',6,1,'응급실(응급)',prefix=': ',suffix='')
    if g('일상배상책임'): pv('TextBox 5',0,1,'일상배상책임',prefix=': ',suffix='')
    if g('대인'): pv('TextBox 9',0,1,'대인',prefix=': ',suffix='')
    if g('대물'): pv('TextBox 9',1,1,'대물',prefix=': ',suffix='')
    if g('합의금'): pv('TextBox 9',2,1,'합의금',prefix=': ',suffix='')
    if g('6주미만'): pv('TextBox 9',3,2,'6주미만',prefix=': ',suffix='')
    if g('변호사'): pv('TextBox 9',4,1,'변호사',prefix=': ',suffix='')
    if g('자부상'): pv('TextBox 9',5,2,'자부상',prefix=': ',suffix='')
    if g('질병일당'): pv('TextBox 22',0,1,'질병일당',prefix=': ',suffix=' / ')
    if g('상해일당'): pv('TextBox 22',1,1,'상해일당',prefix=': ',suffix=' / ')
    if g('1인실 상급병원'): pv('TextBox 22',3,2,'1인실 상급병원',prefix=': ',suffix='')
    if g('1인실 종합병원'): pv('TextBox 22',4,2,'1인실 종합병원',prefix=': ',suffix='')
    if g('간병인'): pv('TextBox 22',7,1,'간병인',prefix=': ',suffix=' / ')
    if g('간호통합병동'): pv('TextBox 22',8,2,'간호통합병동',prefix=': ',suffix='')
    if g('크라운'): pv('TextBox 13',0,1,'크라운',prefix=': ',suffix='')
    if g('임플란트'): pv('TextBox 13',1,1,'임플란트',prefix=': ',suffix='')

    # ── 누락 슬롯 보충 (엑셀 합계 끌어오기) ──
    if g('중입자치료비'): pv('TextBox 14',3,2,'중입자치료비',prefix=': ',suffix='')
    if g('5대골절진단비'): pv('TextBox 7',1,3,'5대골절진단비',prefix=': ',suffix='')
    if g('중증화상진단비'): pv('TextBox 7',3,1,'중증화상진단비',prefix=': ',suffix='')
    if g('허혈성수술비'): pv('TextBox 17',6,2,'허혈성수술비',prefix=': ',suffix='')
    if g('5대골절수술비'): pv('TextBox 19',5,3,'5대골절수술비',prefix=': ',suffix='')
    if g('화상수술비'): pv('TextBox 19',6,1,'화상수술비',prefix=': ',suffix='')
    if g('창상봉합술'): pv('TextBox 19',8,2,'창상봉합술',prefix=': ',suffix='')
    if g('질병중환자실'): pv('TextBox 22',2,2,'질병중환자실',prefix=': ',suffix=' / ')
    if g('상해중환자실'): pv('TextBox 22',2,5,'상해중환자실',prefix=': ',suffix='')
    if g('1인실 상급병원'): pv('TextBox 22',3,2,'1인실 상급병원',prefix=': ',suffix='')
    if g('1인실 종합병원'): pv('TextBox 22',4,2,'1인실 종합병원',prefix=': ',suffix='')

    # ★v29t CI 담보값 노란 배경(§8.4·§11): 중대한 계열을 해당 칸에 표기 + 값 run만 노랑 하이라이트
    from pptx.oxml.ns import qn as _ciqn
    import copy as _cicopy
    from pptx.text.text import _Run as _ciRunCls
    _CIHL_AFTER=[_ciqn('a:uLnTx'),_ciqn('a:uLn'),_ciqn('a:uFillTx'),_ciqn('a:uFill'),_ciqn('a:latin'),_ciqn('a:ea'),
               _ciqn('a:cs'),_ciqn('a:sym'),_ciqn('a:hlinkClick'),_ciqn('a:hlinkMouseOver'),_ciqn('a:rtl'),_ciqn('a:extLst')]
    def _hl_yellow(run):
        rPr=run._r.get_or_add_rPr()
        for old in rPr.findall(_ciqn('a:highlight')): rPr.remove(old)
        hl=rPr.makeelement(_ciqn('a:highlight'),{}); hl.append(rPr.makeelement(_ciqn('a:srgbClr'),{'val':'FFFF00'}))
        ins=None
        for ch in rPr:
            if ch.tag in _CIHL_AFTER: ins=ch; break
        if ins is not None: ins.addprevious(hl)
        else: rPr.append(hl)
    def _ci_run(box,pidx,std,sep):
        v=totals.get(std,0)
        if not v or box not in by: return
        tf=by[box].text_frame
        if pidx>=len(tf.paragraphs): return
        p=tf.paragraphs[pidx]
        if not p.runs: return
        base=p.runs[-1]
        nr_el=_cicopy.deepcopy(base._r); base._r.addnext(nr_el)
        nr=_ciRunCls(nr_el,p); nr.text=f'{sep}{v:,}'
        _hl_yellow(nr)
    def _ci_split(box,label,ci_std,extra_std):
        # ★v29t: 라벨줄 + [CI값(노랑)] + [+일반값] 을 별도 run으로 구성 — 개행 포함 run의 하이라이트 미표시(파워포인트) 방지
        civ=totals.get(ci_std,0)
        if not civ or box not in by: return
        tf=by[box].text_frame; p=tf.paragraphs[0]
        if not p.runs: return
        base=p.runs[0]
        for _r in list(p.runs[1:]): _r._r.getparent().remove(_r._r)
        base.text=f'{label}\n'
        el1=_cicopy.deepcopy(base._r); base._r.addnext(el1)
        r1=_ciRunCls(el1,p); r1.text=f'{civ:,}'; _hl_yellow(r1)
        exv=totals.get(extra_std,0)
        if exv:
            el2=_cicopy.deepcopy(base._r); el1.addnext(el2)
            r2=_ciRunCls(el2,p); r2.text=f'+{exv:,}'
            try: r2.font.color.rgb=(_BLUE if _gensum.get(extra_std) else _BLACK)
            except: pass
    # ★★★★★v243(지점장 지시 2026.07.25): <b>보장분석지 PPT 뇌 칸도 축을 따라간다</b>.
    #   구 코드는 `중대한 뇌졸증`만 봐서, 축이 뇌출혈인 CI(신한·DB 실측)는 <b>PPT에 아무것도 안 찍혔다</b>.
    #   → 끝열 합계에 <b>중대한 뇌출혈</b>이 있으면 그 축으로 라벨·값을 바꾼다.
    if (totals.get('중대한 뇌출혈',0) or 0) > (totals.get('중대한 뇌졸증',0) or 0):
        _ci_split('TextBox 47','뇌출혈','중대한 뇌출혈','뇌출혈진단비')
    else:
        _ci_split('TextBox 47','뇌졸증','중대한 뇌졸증','뇌졸증진단비')
    _ci_split('TextBox 55','급성심근','중대한 급성심근','급성심근경색')
    _ci_run('TextBox 14',0,'중대한 암','+')
    _ci_run('TextBox 10',3,'중대한CI적용','+')
    _autofit_ppt(by)
    prs.save(out); return True


# ★v50(지점장 '다10'): 제목·날짜만 예외(18pt). 실손박스(59)도 10pt 대상으로 편입.
_HEADER_BOXES={'TextBox 21','TextBox 36','TextBox 35','TextBox 29'}
_SURGERY_BOXES={'TextBox 17','TextBox 19'}   # ★v29t: 질병수술·상해수술 9.0pt 고정(지점장 2026.07.02), 1~5종 줄만 축소 허용
def _autofit_ppt(by):
    """겹침·단락내림 방지(§11): 값박스 word_wrap off + 최장 단락 기준 박스 단위 축소.
    수술 박스 2개는 8.9pt 고정, '1~5종' 제목줄·슬래시 괄호줄만 축소 허용."""
    for _bn, sh in by.items():
        if _bn in _HEADER_BOXES: continue
        tf = sh.text_frame
        try:
            tf.word_wrap = False
            w_in = sh.width / 914400.0
        except: continue
        if _bn in _SURGERY_BOXES:
            # ★수술비 폰트(지점장 규정 2026.07.07): 1-5종 슬래시 줄만 6pt, 나머지 수술 줄은 9pt 고정(축소 금지)
            for p in tf.paragraphs:
                ptxt=''.join(r.text for r in p.runs)
                _sz = 6.0 if ('/' in ptxt) else 10.0  # ★v50: 슬래시(1-5종)만 6pt, 그 외 10pt
                for r in p.runs:
                    if r.text:
                        try: r.font.size = Pt(_sz)
                        except: pass
            continue
        runs_all = [r for p in tf.paragraphs for r in p.runs if r.text]
        if not runs_all: continue
        # ★v50 정본(지점장 2026.07.13): 값 폰트는 전부 10pt 고정.
        #   - v50(2026.07.13): 값·라벨 전부 10pt(지점장 '다10'). 예외=제목·날짜(18pt)·수술 1~5종(6pt).
        #   - 6pt는 수술 1~5종 슬래시 줄에만 허용(위 _SURGERY_BOXES 분기).
        for r in runs_all:
            try:
                cur = r.font.size.pt if r.font.size else 9.0
                if cur < 18.0 and cur != 10.0:
                    r.font.size = Pt(10)
            except: pass


def build_chiryo(data, out, totals=None, unmapped=None):
    """치료비 정리 폼: 고객명/날짜 + [확인](AI 미매핑) 항목을 카테고리별로 채움.
    추측 금지 — 박스 라벨이 명시한 'AI가 못 채운 항목'에 실제 미매핑 목록만 주입."""
    if not os.path.exists(TPL_TX): return False
    prs = Presentation(TPL_TX); sl = prs.slides[0]
    by = {sh.name:sh for sh in sl.shapes if sh.has_text_frame}
    client = data['client']; now = datetime.datetime.now()
    def first_run_set(box, text):
        if box not in by: return
        tf = by[box].text_frame
        if tf.paragraphs and tf.paragraphs[0].runs:
            tf.paragraphs[0].runs[0].text = text
    if 'TextBox 21' in by:
        by['TextBox 21'].text_frame.word_wrap=False
        by['TextBox 21'].text_frame.auto_size=MSO_AUTO_SIZE.NONE  # 도형 고정
        rs=by['TextBox 21'].text_frame.paragraphs[0].runs
        if rs: rs[0].text=f'{client} 님의 보장'
        if len(rs)>1: rs[1].text='(전)'
    first_run_set('TextBox 36', f'{now.year}년')
    first_run_set('TextBox 35', f'{now.month:02d}월')
    first_run_set('TextBox 29', f'{now.day:02d}일 기준')
    # [확인] 미매핑 항목 → 회사별 묶어 본문 박스에 기재(있을 때만)
    unmapped = unmapped or []
    if unmapped:
        lines = [f"{comp} {raw}: {amt:,}" for (col,comp,raw,amt,note) in unmapped]
        blob = '\n'.join(lines[:20])
        for box in ['TextBox 25','TextBox 32','TextBox 37','TextBox 51']:
            if box in by:
                tf=by[box].text_frame
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = '⚠ AI 미매핑(별첨 직접확인):\n'+blob
                break
    _autofit_ppt(by)
    prs.save(out); return True

def make_summary(data):
    contracts=data['contracts']; cust=data['client']
    total_premium=sum(ct['premium'] for ct in contracts)
    갱신수=sum(1 for ct in contracts if ct['renewal']=='갱신')
    lines=[f"<b>👤 {cust} 고객님 분석 완료</b>","",
           f"📋 <b>계약 현황</b>",
           f"  • 총 계약 수: <b>{len(contracts)}건</b>",
           f"  • 갱신형: {갱신수}건 / 비갱신형: {len(contracts)-갱신수}건",
           f"  • 월 보험료 합계: <b>{total_premium:,}원</b>","","🏢 <b>가입 회사</b>"]
    for ct in contracts:
        tag='🔵갱신' if ct['renewal']=='갱신' else '🔴비갱신' if '비갱신' in ct['renewal'] else '🟢완납'
        lines.append(f"  • {ct['company']} [{tag}] {ct['premium']:,}원")
    totals={}
    for ct in contracts:
        for raw,amt in ct['dambo'].items():
            std=resolve(raw)
            if std: totals[std]=totals.get(std,0)+amt
    key=[('일반암','🎗암진단'),('뇌혈관진단비','🧠뇌혈관'),('협심증','❤️허혈성'),
         ('급성심근경색','❤️급성심근'),('상해사망','💀상해사망'),('질병사망(80세)','💀질병사망'),('입원','🏥실손')]
    found=[(lbl,totals[k]) for k,lbl in key if k in totals and totals[k]>0]
    if found:
        lines+=["","🔑 <b>주요 담보 합계 (만원)</b>"]
        for lbl,amt in found: lines.append(f"  • {lbl}: <b>{amt:,}만원</b>")
    return '<br>'.join(lines)

INDEX_HTML = r'''<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0">
<title>MAKEONE 보장설명서</title>
<style>
:root{--bg:#0c0d10;--panel:#15171c;--line:#2a2d34;--acc:#7C3AED;--acc2:#A78BFA;--ink:#EAECEF;--mute:#929aa6;--green:#4ADE80;--blue:#5B9BFF}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:'Pretendard','Noto Sans KR',sans-serif;line-height:1.55}
#gate{position:fixed;inset:0;z-index:100;background:var(--bg);display:flex;flex-direction:column;align-items:center;justify-content:center;padding:30px 26px;text-align:center}
#gate .kick{font-size:14px;font-weight:800;letter-spacing:.45em;color:var(--acc);margin-bottom:14px}
#gate h1{font-size:30px;font-weight:800;margin-bottom:14px;color:var(--acc2)}
#gate .s{font-size:14px;color:var(--mute);margin-bottom:38px}
#gate .pw{width:100%;max-width:420px;background:#1a1c22;border:1px solid var(--line);border-radius:14px;padding:18px 20px;font-size:17px;color:var(--ink);text-align:center;letter-spacing:.3em;outline:none}
#gate .pw:focus{border-color:var(--acc)}
#gate .go{width:100%;max-width:420px;margin-top:14px;border:none;border-radius:14px;padding:18px;font-size:17px;font-weight:800;color:#fff;background:var(--acc);cursor:pointer}
#gate .err{color:var(--acc2);font-size:13px;font-weight:700;margin-top:14px;min-height:18px}
.shake{animation:sh .35s}@keyframes sh{0%,100%{transform:translateX(0)}25%{transform:translateX(-8px)}75%{transform:translateX(8px)}}
.app{max-width:520px;margin:0 auto;height:100vh;display:none;flex-direction:column}
header{padding:14px 18px;border-bottom:1px solid var(--line);background:linear-gradient(135deg,#17131f,#0d0e11 60%,#1a1426);display:flex;align-items:center;gap:10px}
.logo{width:32px;height:32px;border-radius:9px;border:1px solid var(--acc);display:flex;align-items:center;justify-content:center;font-size:16px}
h1{font-size:14px;font-weight:800}h1 b{color:var(--acc2)}.sub{font-size:10px;color:var(--mute)}
.chat{flex:1;overflow-y:auto;padding:16px 12px;display:flex;flex-direction:column;gap:12px}
.msg{max-width:90%;font-size:13px}
.me{align-self:flex-end;background:rgba(124,58,237,.14);border:1px solid rgba(124,58,237,.32);border-radius:14px 14px 4px 14px;padding:9px 13px}
.bot{align-self:flex-start;background:var(--panel);border:1px solid var(--line);border-radius:14px 14px 14px 4px;padding:11px 14px;width:100%}
.file-cards{display:flex;flex-direction:column;gap:8px;margin-top:10px}
.file-card{display:flex;align-items:center;gap:11px;border-radius:12px;padding:11px 13px}
.file-card.xl{background:rgba(74,222,128,.06);border:1px solid rgba(74,222,128,.3)}
.file-card.pt{background:rgba(91,155,255,.06);border:1px solid rgba(91,155,255,.3)}
.file-card .ic{font-size:22px}.file-card .nm{flex:1;font-size:12.5px;font-weight:700}
.file-card .dl{font-size:11px;font-weight:800;padding:5px 11px;border-radius:8px}
.file-card.xl .dl{color:var(--green);background:rgba(74,222,128,.12)}
.file-card.pt .dl{color:var(--blue);background:rgba(91,155,255,.12)}
.summary-box{background:#1a1f2a;border:1px solid #2a3040;border-radius:10px;padding:12px 14px;margin-top:10px;font-size:12px;line-height:1.7}
.err{color:#ffb4b4;font-size:12px}
.spin{width:22px;height:22px;border:3px solid var(--line);border-top-color:var(--acc);border-radius:50%;animation:sp .8s linear infinite;display:inline-block;vertical-align:middle}
@keyframes sp{to{transform:rotate(360deg)}}
.bar{padding:12px;border-top:1px solid var(--line);display:flex;gap:9px;background:var(--bg)}
.up{flex:1;border:1.5px dashed rgba(124,58,237,.5);border-radius:12px;padding:13px;text-align:center;font-size:13px;font-weight:700;cursor:pointer;color:var(--acc2)}
.send{border:none;border-radius:12px;padding:0 20px;font-weight:800;font-size:14px;background:var(--acc);color:#fff;cursor:pointer}
.send:disabled{opacity:.4}
.qbar{padding:8px 12px;border-top:1px solid var(--line);display:none;gap:8px;background:var(--bg)}
.qinput{flex:1;background:#1a1c22;border:1px solid var(--line);border-radius:10px;padding:10px 14px;font-size:13px;color:var(--ink);outline:none}
.qinput:focus{border-color:var(--acc)}.qinput::placeholder{color:var(--mute)}
.qbtn{border:none;border-radius:10px;padding:0 16px;font-weight:800;font-size:13px;background:#2a2d34;color:var(--ink);cursor:pointer}
.qbtn:hover{background:#3a3d44}.qbtn:disabled{opacity:.4}
.qlbl{font-size:10px;color:var(--mute);padding:2px 12px;display:none}
footer{text-align:center;font-size:10px;color:var(--mute);padding:8px}footer b{color:var(--acc2)}
</style></head><body>
<div id="gate">
  <div class="kick">MAKEONE</div><h1>MAKEONE 보장설명서</h1>
  <div class="s">접속 비밀번호를 입력하세요</div>
  <input id="pw" class="pw" type="password" inputmode="numeric" placeholder="비밀번호" autocomplete="off">
  <button id="go" class="go">접속</button><div id="gerr" class="err"></div>
</div>
<div class="app" id="app">
  <header><div class="logo">📋</div><div><h1>MAKEONE <b>보장설명서</b></h1>
    <div class="sub">보장분석 리포트 PDF 1개 → 엑셀+PPT 개별 다운로드 · 최은혜 지점장</div></div></header>
  <div class="chat" id="chat">
    <div class="msg bot">채널에서 받은 <b>보장분석 리포트 PDF 원본</b> 1개를 올려주세요. 엑셀·PPT를 각각 드려요.<br><br>
      <span style="font-size:11px;color:var(--mute)">※ 받은 PDF를 <b>그대로</b> 올리세요. 인쇄·재스캔·OCR 변환하면 금액이 깨져 분석이 틀어집니다.<br>
      ※ 롯데(let:) · KB · 메리츠 리포트 모두 원본 PDF 그대로 인식합니다.</span></div>
  </div>
  <div class="bar">
    <label class="up" id="upp">📑 <span id="upplabel">보장분석 PDF 선택</span></label>
    <label class="up" id="up">📄 <span id="uplabel">TXT (구방식)</span></label>
    <button class="send" id="send" disabled>분석</button>
  </div>
  <div class="qlbl" id="qlbl">📋 분석된 보장분석지에 대해 질문하세요</div>
  <div class="qbar" id="qbar">
    <input class="qinput" id="qinput" placeholder="예: 심장 담보 왜 빠졌어요?" autocomplete="off">
    <button class="qbtn" id="qbtn">질문</button>
  </div>
  <footer>미래를 <b>바르게</b> 설계합니다 · BARUM <b>v32-ocrpdf</b></footer>
</div>
<input type="file" id="fi" accept=".txt,text/plain" style="display:none">
<input type="file" id="fp" accept=".pdf,application/pdf" style="display:none">
<script>
const $=s=>document.querySelector(s);let ACCESS='';
async function unlock(){const v=$("#pw").value;$("#gerr").textContent="확인 중…";
  try{const r=await fetch("/check",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({pw:v})});
    const j=await r.json();if(j.ok){ACCESS=v;$("#gerr").textContent="";$("#gate").style.display="none";$("#app").style.display="flex";}else{fail();}}
  catch(e){$("#gerr").textContent="서버 연결 실패";}}
function fail(){$("#gerr").textContent="비밀번호가 올바르지 않습니다.";$("#gate").classList.add("shake");setTimeout(()=>$("#gate").classList.remove("shake"),350);$("#pw").value="";$("#pw").focus();}
$("#go").onclick=unlock;$("#pw").addEventListener("keydown",e=>{if(e.key==="Enter")unlock();});window.addEventListener("load",()=>$("#pw").focus());
const chat=$("#chat");let file=null;let pdfFile=null;
function _syncSend(){$("#send").disabled=!(file||pdfFile);}
$("#up").onclick=()=>$("#fi").click();
$("#upp").onclick=()=>$("#fp").click();
$("#fi").onchange=e=>{file=e.target.files[0]||null;$("#uplabel").textContent=file?file.name:"TXT (구방식)";_syncSend();};
$("#fp").onchange=e=>{pdfFile=e.target.files[0]||null;$("#upplabel").textContent=pdfFile?pdfFile.name:"보장분석 PDF 선택";_syncSend();};
function esc(s){return String(s==null?"":s).replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));}
function add(html,cls){const d=document.createElement("div");d.className="msg "+cls;d.innerHTML=html;chat.appendChild(d);chat.scrollTop=chat.scrollHeight;return d;}
function b64toBlob(b64,mime){const bin=atob(b64);const arr=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);return new Blob([arr],{type:mime});}
function dl(blob,fname){const u=URL.createObjectURL(blob);const a=document.createElement("a");a.href=u;a.download=fname;document.body.appendChild(a);a.click();document.body.removeChild(a);setTimeout(()=>URL.revokeObjectURL(u),3000);}
const XLMIME="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PTMIME="application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PDFMIME="application/pdf";
let savedFiles={};
function reDL(k){const f=savedFiles[k];if(f&&f.b64){dl(b64toBlob(f.b64,f.mime),f.name);}}
$("#send").onclick=async()=>{
  if(!file&&!pdfFile)return;add("📄 "+esc((file?file.name:"")+(file&&pdfFile?" + ":"")+(pdfFile?pdfFile.name:"")),"me");
  $("#send").disabled=true;$("#up").style.opacity=.5;
  const loading=add('<div style="display:flex;align-items:center;gap:11px"><span class="spin"></span><div style="flex:1"><div id="ldmsg" style="font-weight:800">📄 TXT 파싱 중…</div><div id="ldtime" style="font-size:11px;color:var(--mute);margin-top:2px">0초 · 기다려 주세요</div></div></div>',"bot");
  const t0=Date.now();const steps=["📄 TXT 파싱 중…","🔎 담보 추출 중…","📊 엑셀 생성 중…","🖼 PPT 채우는 중…","✅ 완성 중…"];let si=0;
  const timer=setInterval(()=>{si=Math.min(si+1,steps.length-1);const s=Math.floor((Date.now()-t0)/1000);const tm=document.getElementById("ldtime");const mm=document.getElementById("ldmsg");if(tm)tm.textContent=s+"초 경과";if(mm)mm.textContent=steps[si];},8000);
  const fd=new FormData();
  if(file&&pdfFile){fd.append("file",file);fd.append("file2",pdfFile);}
  else if(file){fd.append("file",file);}
  else{fd.append("file",pdfFile);}
  fd.append("pw",ACCESS);
  let j=null;
  try{
    const r=await fetch("/analyze",{method:"POST",body:fd});clearInterval(timer);loading.remove();
    j=await r.json();
    if(!j.ok){
      /* ★v94: '⚠ 실패'만 뜨고 원인을 알 수 없던 문제 — 서버가 보내주는 trace를 화면에 같이 찍는다. */
      var _m = esc(j.error||"실패(서버가 오류 문구를 못 보냄)");
      var _t = j.trace ? String(j.trace) : "";
      if(_t){ var _tail=_t.split("\n").slice(-8).join("\n");
              _m += '<br><span style="font-size:11px;opacity:.85;white-space:pre-wrap">'+esc(_tail)+'</span>'; }
      add('<span class="err">⚠ '+_m+'</span>',"bot");
    }
    else{
      savedFiles={};
      const xlBlob=b64toBlob(j.xlsx_b64,XLMIME);
      dl(xlBlob,j.xlsx_name);
      savedFiles.xlsx={b64:j.xlsx_b64,name:j.xlsx_name,mime:XLMIME};
      let ptCard='';
      if(j.pptx_b64){
        const ptBlob=b64toBlob(j.pptx_b64,PTMIME);
        setTimeout(()=>dl(ptBlob,j.pptx_name),800);
        savedFiles.pptx={b64:j.pptx_b64,name:j.pptx_name,mime:PTMIME};
        ptCard=`<div class="file-card pt" onclick="reDL('pptx')" style="cursor:pointer"><span class="ic">📊</span><span class="nm">${esc(j.pptx_name)}<br><span style="font-size:10px;color:var(--mute)">보장분석 PPT</span></span><span class="dl">💾 다시저장</span></div>`;}
      if(j.chiryo_b64){
        const txBlob=b64toBlob(j.chiryo_b64,PTMIME);
        setTimeout(()=>dl(txBlob,j.chiryo_name),1600);
        savedFiles.chiryo={b64:j.chiryo_b64,name:j.chiryo_name,mime:PTMIME};
        ptCard+=`<div class="file-card pt" onclick="reDL('chiryo')" style="cursor:pointer"><span class="ic">🩺</span><span class="nm">${esc(j.chiryo_name)}<br><span style="font-size:10px;color:var(--mute)">치료비 정리 PPT</span></span><span class="dl">💾 다시저장</span></div>`;}
      if(j.report_b64){
        const rpBlob=b64toBlob(j.report_b64,PDFMIME);
        setTimeout(()=>dl(rpBlob,j.report_name),2400);
        savedFiles.report={b64:j.report_b64,name:j.report_name,mime:PDFMIME};
        ptCard+=`<div class="file-card pt" onclick="reDL('report')" style="cursor:pointer"><span class="ic">📄</span><span class="nm">${esc(j.report_name)}<br><span style="font-size:10px;color:var(--mute)">보장설명지 PDF</span></span><span class="dl">💾 다시저장</span></div>`;}
      if(j.report_pptx_b64){
        const rpxBlob=b64toBlob(j.report_pptx_b64,PTMIME);
        setTimeout(()=>dl(rpxBlob,j.report_pptx_name),3000);
        savedFiles.reportpptx={b64:j.report_pptx_b64,name:j.report_pptx_name,mime:PTMIME};
        ptCard+=`<div class="file-card pt" onclick="reDL('reportpptx')" style="cursor:pointer"><span class="ic">📋</span><span class="nm">${esc(j.report_pptx_name)}<br><span style="font-size:10px;color:var(--mute)">보장진단서 PPT (편집가능)</span></span><span class="dl">💾 다시저장</span></div>`;}
      add('<b>✅ 분석 완료!</b> <span style="font-size:11px;color:var(--mute)">(카드 누르면 다시 저장)</span><div class="summary-box">'+j.summary+'</div><div class="file-cards">'+
        `<div class="file-card xl" onclick="reDL('xlsx')" style="cursor:pointer"><span class="ic">📗</span><span class="nm">${esc(j.xlsx_name)}<br><span style="font-size:10px;color:var(--mute)">보장진단 엑셀</span></span><span class="dl">💾 다시저장</span></div>`+ptCard+'</div>',"bot");}
  }catch(e){clearInterval(timer);loading.remove();add('<span class="err">오류: '+esc(e.message)+'</span>',"bot");}
  if(j&&j.data){analysisData=j.data;document.getElementById("qbar").style.display="flex";document.getElementById("qlbl").style.display="block";}
  file=null;$("#uplabel").textContent="TXT (구방식)";$("#send").disabled=true;$("#fi").value="";$("#up").style.opacity=1;
  if(j&&j.report_error){add('<span class="err">⚠ 보장설명지 PDF 생성 실패: '+esc(j.report_error)+'</span>',"bot");}
  if(j&&j.report_pptx_error){add('<span class="err">⚠ 보장진단서 PPT 생성 실패: '+esc(j.report_pptx_error)+'</span>',"bot");}
  if(j&&j.ok){add('다음 고객 TXT를 올리면 이어서 분석합니다.',"bot");}
};
let analysisData=null;
function askAI(){
  const q=document.getElementById("qinput").value.trim();
  if(!q||!analysisData)return;
  add("💬 "+esc(q),"me");
  document.getElementById("qinput").value="";
  document.getElementById("qbtn").disabled=true;
  const loading=add('<span class="spin"></span> 분석 중…',"bot");
  fetch("/ask",{method:"POST",headers:{"Content-Type":"application/json"},
    body:JSON.stringify({pw:ACCESS,question:q,data:analysisData})})
  .then(r=>r.json()).then(j=>{
    loading.remove();
    add(j.ok?esc(j.answer):'<span class="err">⚠ '+esc(j.error||"오류")+'</span>',"bot");
    document.getElementById("qbtn").disabled=false;
  }).catch(e=>{loading.remove();add('<span class="err">오류: '+esc(e.message)+'</span>',"bot");document.getElementById("qbtn").disabled=false;});
}
document.addEventListener("DOMContentLoaded",function(){
  document.getElementById("qinput").addEventListener("keydown",function(e){if(e.key==="Enter")askAI();});
  document.getElementById("qbtn").onclick=askAI;
});
</script>
<script>if("serviceWorker" in navigator){navigator.serviceWorker.getRegistrations().then(function(rs){rs.forEach(function(r){r.unregister();});}).catch(function(){});}</script></body></html>'''

@app.get('/health')
def health():
    _cib = ci_selftest()   # ★v238 CI 자가진단 — 실패하면 즉시 노출
    return {'ok':True,'version':'v254-cijugye-20260726',
            'ci_selftest': ('PASS %d/%d' % (len(_CI_SELFTEST)-len(_cib), len(_CI_SELFTEST))) if not _cib else ('FAIL: '+' | '.join(_cib[:6]))}

# ★★v101 진단 엔드포인트(2026.07.20): 폰에서 링크 한 번만 눌러
#   Railway 컨테이너에 pdftotext(poppler)가 실제로 살아있는지 확인한다.
#   'KB(PDF)는 죽고 롯데(txt)는 산다'의 원인을 서버 로그 없이 확정하기 위함.
@app.get('/diag')
def diag():
    import subprocess, shutil
    out = {'version': 'v254-cijugye-20260726'}
    out['pdftotext_path'] = shutil.which('pdftotext') or '없음(★범인)'
    try:
        r = subprocess.run(['pdftotext', '-v'], capture_output=True, text=True, timeout=20)
        out['pdftotext_ver'] = ((r.stderr or '') + (r.stdout or '')).strip().split('\n')[0][:80]
    except Exception as e:
        out['pdftotext_ver'] = 'ERR ' + str(e)[:80]
    for d in ('/usr/share/poppler', '/usr/share/poppler/cMap'):
        try: out['poppler_data:' + d] = os.path.isdir(d)
        except Exception: out['poppler_data:' + d] = 'ERR'
    out['api_key'] = bool(os.environ.get('ANTHROPIC_API_KEY', ''))
    out['pdf2image'] = _mod_ok('pdf2image')
    out['weasyprint'] = _mod_ok('weasyprint')
    return out

def _mod_ok(m):
    try:
        __import__(m); return True
    except Exception as e:
        return 'ERR ' + str(e)[:60]

@app.get('/',response_class=HTMLResponse)
def home(): return INDEX_HTML

@app.post('/check')
async def check_pw(body:dict): return {'ok':body.get('pw')==PW}

@app.post('/analyze')
async def analyze(file:UploadFile=File(...), file2:UploadFile=File(None), pw:str=Form('')):
    if pw!=PW: return JSONResponse({'ok':False,'error':'비밀번호 오류'})
    # ★v30z5 입력 확장: .txt + .pdf 둘 다 수용. 두 파일 동시 업로드 시 OR·병합.
    _files=[f for f in (file, file2) if f is not None]
    _txt_f=next((f for f in _files if (f.filename or '').lower().endswith('.txt')), None)
    _pdf_f=next((f for f in _files if (f.filename or '').lower().endswith('.pdf')), None)
    if not _txt_f and not _pdf_f:
        return JSONResponse({'ok':False,'error':'TXT 또는 PDF 파일이 필요합니다.'})

    txt=''; fname=(_txt_f or _pdf_f).filename
    if _txt_f:
        raw=await _txt_f.read()
        for enc in ['utf-8','cp949','euc-kr']:
            try: txt=raw.decode(enc); break
            except: pass
        else: txt=raw.decode('utf-8',errors='ignore')

    # ★OCR PDF 우선(2026.07.07 지점장 정답): PDF 있으면 pdftotext 직독을 주 소스. 깨지면 txt 폴백.
    src_note=''
    try:
        _txt_data = parse_txt(txt, fname) if txt.strip() else None
    except Exception:
        _txt_data=None
    _pdf_data=None; pdf_txt=''; _img_pdf_nokey=False; _img_prod=''
    if _pdf_f:
        pdf_bytes=await _pdf_f.read()
        # ★v60 이미지 PDF 진단: 텍스트레이어 직독이 0글자면 = 이미지 전용 PDF(글자 없음).
        #   'Microsoft Print To PDF'로 다시 저장하면 글자층이 통째 이미지가 돼 직독 불가.
        #   이때 유일한 경로는 비전 OCR(API키 필요) — 키 없으면 명확히 알린다.
        try:
            import subprocess as _sp, tempfile as _tf
            with _tf.NamedTemporaryFile(suffix='.pdf', delete=False) as _tfp:
                _tfp.write(pdf_bytes); _tpp=_tfp.name
            _rawtl=_sp.run(['pdftotext','-layout',_tpp,'-'],capture_output=True,text=True,timeout=60).stdout
            try: os.unlink(_tpp)
            except: pass
            # ★v132: '글자수<30'만 보면 못 잡는 케이스가 있다 — 브라우저 '인쇄→PDF 저장'본은
            #   Type3 서브셋 폰트 + ToUnicode 없음이라 pdftotext가 글자는 1만자 넘게 뱉지만
            #   한글은 0자다(전부 깨진 제어문자). 실측: 권양영_보장.pdf = 18,333자 / 한글 0자.
            #   따라서 판정 기준을 '한글 글자 수'로 바꾼다.
            _hangul = sum(1 for _c in (_rawtl or '') if '\uac00' <= _c <= '\ud7a3')
            # ★★★v231 (2026.07.25 한정환 실측): Producer도 함께 읽는다.
            #   `Microsoft: Print To PDF` / `Print to PDF` 는 <b>글자층이 통째 이미지</b>가 된 확정 증거다.
            #   실측 대조 — 이정화(정상) Producer=`oz` 37,278자·한글 4,566자 /
            #              한정환(실패) Producer=`Microsoft: Print To PDF` <b>12자·한글 0자</b>(전부 페이지구분자).
            try:
                _pinf=_sp.run(['pdfinfo',_tpp],capture_output=True,text=True,timeout=30).stdout
            except Exception: _pinf=''
            _prod=''
            for _l in (_pinf or '').split('\n'):
                if _l.lower().startswith('producer:'): _prod=_l.split(':',1)[1].strip()
            _is_print = ('print to pdf' in _prod.lower()) or ('microsoft' in _prod.lower())
            print(f'[PDF_DIAG] chars={len(_rawtl or "")} hangul={_hangul} producer={_prod!r} printpdf={_is_print}')
            # ★★v231: <b>한글이 없으면 API 키가 있어도 이미지 PDF로 확정 안내</b>한다.
            #   구 조건은 `and not ANTHROPIC_API_KEY`라 <b>키가 있으면 안내가 안 나가고</b> 비전 OCR로 흘렀다.
            #   그러나 정본은 <b>OCR 금지</b>(3사 실측으로 데이터 파괴 확인) → 원본 재업로드를 요구하는 게 맞다.
            if not _rawtl or len(_rawtl.strip())<30 or _hangul<100:
                _img_pdf_nokey=True
                _img_prod=_prod
        except Exception: pass
        pdf_txt=pdf_to_txt(pdf_bytes)
        if pdf_txt.strip():
            try:
                _pdf_data=parse_txt(pdf_txt, _pdf_f.filename)
            except Exception:
                _pdf_data=None
    if _pdf_data and _pdf_data.get('contracts') and not _looks_broken(_pdf_data):
        data=_pdf_data; src_note='OCR PDF 직독(주)'
    elif _txt_data and _txt_data.get('contracts'):
        data=_txt_data; src_note='TXT 입력'
    elif _pdf_data and _pdf_data.get('contracts'):
        data=_pdf_data; src_note='OCR PDF(깨짐 감지)'
    else:
        data=_txt_data; src_note='추출 실패'

    try:
        if not data or not data.get('contracts'):
            if _img_pdf_nokey:
                _why = f'(생성기: {_img_prod})' if _img_prod else ''
                _vf = globals().get('_VISION_FAIL','')
                return JSONResponse({'ok':False,'source':'이미지PDF',
                    'error':f'글자를 읽을 수 없는 PDF입니다 {_why} — 한글 0자로 추출됩니다. '
                            '"인쇄 → PDF로 저장"으로 만든 파일은 글자층이 통째로 이미지가 되어 분석이 불가능합니다. '
                            'let: 리포트 화면에서 <b>인쇄가 아니라 PDF 다운로드(저장)</b> 버튼으로 받은 '
                            '원본 파일을 손대지 말고 그대로 올려주세요.'
                            + (f' [비전OCR 진단] {_vf}' if _vf else '')})
            # ★v101: 원인을 화면에서 바로 알 수 있게 추출 단계 수치를 함께 노출(진단용)
            _dbg = (f'계약을 찾지 못했습니다. [진단] 경로={src_note} / '
                    f'PDF추출글자={len(pdf_txt or "")} / TXT입력글자={len(txt or "")} / '
                    f'PDF계약={len((_pdf_data or {}).get("contracts") or [])} / '
                    f'TXT계약={len((_txt_data or {}).get("contracts") or [])}')
            return JSONResponse({'ok':False,'error':_dbg,'source':src_note})
        cust=data['client']; d=tempfile.mkdtemp(); now=datetime.datetime.now()
        xl=os.path.join(d,f'보장진단_{cust}.xlsx'); pt=os.path.join(d,f'보장분석지_{cust}.pptx')
        tx=os.path.join(d,f'치료비정리_{cust}.pptx')
        unmapped=build_excel(data,xl)
        if not recalc_xlsx(xl): inject_sum_cache(xl)   # ★v29u: Railway(LibreOffice 없음)에서도 합계 캐시 보장
        ppt_totals, sq, ss, ppt_splits = read_excel_totals(xl)   # 등식2: PPT는 완성 엑셀만 읽음
        ppt_ok=build_ppt(data,pt,ppt_totals,sq,ss,ppt_splits)
        # 치료비정리 PPT 폐기(v29) — 내용 부실, 보장설명지 PDF로 대체
        xlsx_b64=base64.b64encode(open(xl,'rb').read()).decode()
        response={'ok':True,'xlsx_b64':xlsx_b64,'xlsx_name':f'보장진단_{cust}.xlsx',
                  'summary':make_summary(data),'pptx_ready':ppt_ok,'source':src_note}
        if ppt_ok and os.path.exists(pt):
            response['pptx_b64']=base64.b64encode(open(pt,'rb').read()).decode()
            response['pptx_name']=f'보장분석지_{cust}.pptx'
        # ── 보장설명서: 충족률 PDF + ★보장진단서 PPT(편집가능) — 둘 다 실패해도 엑셀·PPT는 유지 ──
        rep=None
        try:
            from coverage_benchmark import map_excel_to_report
            # ★★★v138 흥국화재 10억통장 가입 판정(지점장 실측 정정 2026.07.21):
            #   '리셋월렛'은 <b>상품명이 아니라 담보명</b>으로 들어온다.
            #   실측(장기상) — 상품 '무배당 흥Good 든든한 3N5 간편종합보험' 안의
            #   담보 '플래티넘 건강 리셋월렛II(3대질병고액및중환자실치료,특정…) 100,000'(=10억).
            #   따라서 <b>상품명·담보명 두 곳을 모두</b> 본다(지점장 '2가지 다 적용' 지시).
            _r10=False
            try:
                for _c in (data.get('contracts') or []):
                    _co=str(_c.get('company','')).replace(' ','')
                    if '흥국' not in _co: continue
                    _pd=str(_c.get('product','')).replace(' ','')
                    _hit=('리셋월렛' in _pd) or ('리셋월랫' in _pd)
                    if not _hit:
                        for _k in (_c.get('dambo') or {}):
                            _kk=str(_k).replace(' ','')
                            if ('리셋월렛' in _kk) or ('리셋월랫' in _kk): _hit=True; break
                    if _hit: _r10=True; break
            except Exception: pass
            # ★v146 금액도 함께 전달(진단서 카드에 표기). 만원 단위 원본값.
            _r10amt=0
            try:
                for _c in (data.get('contracts') or []):
                    if '흥국' not in str(_c.get('company','')).replace(' ',''): continue
                    for _k,_v in (_c.get('dambo') or {}).items():
                        _kk=str(_k).replace(' ','')
                        if ('리셋월렛' in _kk) or ('리셋월랫' in _kk):
                            try: _r10amt=max(_r10amt,int(float(_v)))
                            except Exception: pass
            except Exception: pass
            print(f'[R10] 흥국화재 10억통장 가입판정={_r10} 금액={_r10amt}만원')
            rep=map_excel_to_report(xl, settings={'client':cust,'reset10':_r10,'reset10_amt':_r10amt,
                'branch':'온빛센터 바름지점','manager':'최은혜','title':'지점장','phone':''})
        except Exception as _re:
            response['report_error']='분석데이터 생성 실패: '+str(_re)
        if rep is not None:
            # ★ 보장설명지 PDF 별도 생성 중단(2026.07.11 지점장 지시): 보장진단서 PPT가 동일 내용(PDF 페이지 이미지)이라
            #    별도 PDF는 렌더 1회(약 60초)를 중복 유발 → 속도 위해 스킵. 필요 시 이 블록 복구.
            # ★ 보장진단서 PPT (편집가능) — 같은 rep로 생성
            try:
                from report_pptx import build_report_pptx
                rpx=os.path.join(d,f'보장진단서_{cust}.pptx')
                # ★v107: 같은 렌더에서 벡터 PDF(보장설명서)도 함께 받는다(추가 렌더 0회).
                #   PPT는 이미지라 확대·인쇄 시 글자가 뭉갠다 → 선명본은 이 PDF.
                rpdf=os.path.join(d,f'보장설명서_{cust}.pdf')
                build_report_pptx(rep, rpx, pdf_out=rpdf)
                if os.path.exists(rpdf):
                    response['report_b64']=base64.b64encode(open(rpdf,'rb').read()).decode()
                    response['report_name']=f'보장설명서_참고자료_{cust}.pdf'
                if os.path.exists(rpx):
                    response['report_pptx_b64']=base64.b64encode(open(rpx,'rb').read()).decode()
                    response['report_pptx_name']=f'보장진단서_{cust}.pptx'
            except Exception as _pe:
                response['report_pptx_error']=str(_pe)
        return JSONResponse(response)
    except Exception as e:
        return JSONResponse({'ok':False,'error':str(e),'trace':traceback.format_exc()[-1500:]})

# ── AI 질문답 ─────────────────────────────────────────────────────────
import httpx

def build_context(data):
    lines=[f"고객명: {data['client']}", f"계약 수: {len(data['contracts'])}건"]
    for ct in data['contracts']:
        lines.append(f"  - {ct['company']} [{ct['renewal']}] {ct['premium']:,}원")
    totals={}
    for ct in data['contracts']:
        for raw,amt in ct['dambo'].items():
            std=resolve(raw)
            if std: totals[std]=totals.get(std,0)+amt
    lines.append("\n매핑된 담보 합계 (만원):")
    for k,v in sorted(totals.items()): lines.append(f"  - {k}: {v:,}")
    unmapped=[]
    for ct in data['contracts']:
        for raw in ct['dambo']:
            if resolve(raw) is None and not any(x in raw for x in ['(1종)','(2종)','(3종)','(4종)','(5종)']):
                unmapped.append(raw)
    if unmapped:
        lines.append("\n자동매핑 실패 담보 (약관 확인 필요):")
        for u in sorted(set(unmapped)): lines.append(f"  - {u}")
    return '\n'.join(lines)

@app.post('/ask')
async def ask(body:dict):
    if body.get('pw')!=PW: return JSONResponse({'ok':False,'error':'비밀번호 오류'})
    question=body.get('question','').strip()
    data=body.get('data')
    if not question or not data: return JSONResponse({'ok':False,'error':'질문 또는 데이터 없음'})
    context=build_context(data)
    system=("보장분석 전문 AI입니다. 아래 분석 데이터에 관한 질문에만 답하세요.\n"
            "규칙:\n"
            "- 데이터에 없는 담보 -> 현재 계약에 없습니다. 별첨/약관 확인 필요.\n"
            "- 약관 해석 -> 약관을 직접 확인하세요.\n"
            "- 무관한 질문 -> 보장분석 관련 질문만 가능합니다.\n"
            "- 답변 2-3줄 이내, 간결하게, 한국어로.\n"
            f"[분석 데이터]\n{context}")
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp=await client.post('https://api.anthropic.com/v1/messages',
                headers={'x-api-key':os.environ.get('ANTHROPIC_API_KEY',''),
                         'anthropic-version':'2023-06-01','content-type':'application/json'},
                json={'model':'claude-haiku-4-5-20251001','max_tokens':300,
                      'system':system,'messages':[{'role':'user','content':question}]})
        r=resp.json()
        answer=r.get('content',[])[0].get('text','답변을 가져오지 못했습니다.')
        return JSONResponse({'ok':True,'answer':answer})
    except Exception as e:
        return JSONResponse({'ok':False,'error':str(e)})
