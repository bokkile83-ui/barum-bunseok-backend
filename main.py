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
    if any(k in t for k in ('퍼펙트플러스', '퍼텍트플러스', '퍼펙트통합', '퍼텍트통합')): return True
    return any(k in t for k in ('CI보험', '리빙케어', 'GI보험'))


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


def is_excluded(company, product='', contract_date='', expiry_date=''):
    t = re.sub(r'[\s（）()_·]|TM', '', str(company)+str(product))
    for kw in EXCLUDE:
        if kw in t: return True
    if _is_group_ins(product, contract_date, expiry_date): return True   # ★제외 5종: 단체보험
    if _is_oneyear(contract_date, expiry_date): return True              # ★제외 6종: 보험기간 1년(v102)
    if '운전자' in t: return False   # ★운전자·운전자상해보험은 포함(§4)
    # ★자동차보험(다이렉트/애니카/하이카 + 개인용/업무용/영업용/개인소유) = 제외
    if any(b in t for b in ('다이렉트','애니카','하이카','개인용자동차','업무용자동차')) and any(x in t for x in ('개인용','업무용','영업용','개인소유')):
        return True
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

def silson_gen(contract_date, ipv=None, product=''):
    """실손 세대 판별 — 5세대=2026.05부터(정본 확정). 4세대=2021.07~2026.04. 입원한도 3000=구형(1세대). 가입일 없으면 '' → [확인].
    ★v29v: 상품명 연도코드(YYMM 4자리, 예 '1804'=2018.04 출시)가 있으면 판정일로 우선 사용 —
    갱신 재가입일(계약일)로 세대를 오판(예 2018 실손을 4세대로)하는 것 차단."""
    if ipv==3000: return '1세대(구형)'
    try: ym=int(str(contract_date)[:4])*100+int(str(contract_date)[5:7])
    except: ym=0
    _pm=re.search(r'(?<!\d)(0[9]|1[0-9]|2[0-6])\.?(0[1-9]|1[0-2])(?!\d)', str(product or ''))   # ★v30 '25.01' 점 형식 포함
    if _pm:
        _pym=2000*100+int(_pm.group(1))*100+int(_pm.group(2))
        ym=_pym if not ym else min(ym,_pym)
    if not ym: return ''
    if ym<200910:  return '1세대'
    if ym<=201703: return '2세대'
    if ym<=202106: return '3세대'
    if ym<=202604: return '4세대'
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

def rule_extract(block_lines):
    block_lines=_split_cols(block_lines)   # ★다열 별첨 분해(OCR PDF 대응)
    block_lines=[l for l in block_lines if not (('표준금액' in str(l)) or ('권장금액' in str(l)) or ('적정금액' in str(l)))]  # ★표준금액 줄 제외
    """★v29t: 같은줄 우선 + 분리줄(코드/이름랩/금액뭉치) 순서 페어링(누락0). 김진구.txt 6계약 회귀검증 완료."""
    dambo={}; names=[]; amts=[]; pend=None
    def _add(_nm, _amt):
        # ★v61 심뇌혈관수술비 라인단위 분해(지침 §8.3.1 · 지점장 2026.07.15 재확정):
        #   '심뇌혈관…수술' = 심장수술비 + 뇌혈관수술비 각 100% 동일액.
        #   ★중복줄(상해·질병 등 같은 3,000이 2줄) = 합산 아니라 대표(max) — 6,000 오합산 방지.
        #   라인 단위로 쪼개므로 dambo 합산(6,000) 이전에 처리된다.
        _n=re.sub(r'\s','',str(_nm))
        if '심뇌혈관' in _n and '수술' in _n and '[확인]' not in _n:
            for _r in ('심장수술비[묶음]','뇌혈관수술비[묶음]'):   # ★태그 '뇌혈관' 금지→[묶음]
                dambo[_r]=max(dambo.get(_r,0), _amt)
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
            if len(nm)>2 and not re.search(r'납입면제|납입지원',nm): names.append(nm)
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
                if 0 < amt <= 200000 and len(name) > 2:
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
        print('[PDF_VISION] no api key -> skip'); return ''
    try:
        from pdf2image import convert_from_bytes
        import io
        pages = convert_from_bytes(pdf_bytes, dpi=170)
    except Exception as e:
        print(f'[PDF_RENDER_ERR] {e}'); return ''
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
            buf=io.BytesIO(); img.save(buf, format='PNG')
            b=base64.b64encode(buf.getvalue()).decode()
            r=httpx.post('https://api.anthropic.com/v1/messages',
                headers={'x-api-key':key,'anthropic-version':'2023-06-01','content-type':'application/json'},
                json={'model':'claude-haiku-4-5-20251001','max_tokens':4000,
                      'messages':[{'role':'user','content':[
                          {'type':'image','source':{'type':'base64','media_type':'image/png','data':b}},
                          {'type':'text','text':prompt}]}]}, timeout=120)
            if r.status_code==200:
                t=''.join(x.get('text','') for x in r.json().get('content',[]) if x.get('type')=='text')
                if t.strip(): out.append(t)
            else:
                print(f'[PDF_VISION_HTTP] p{idx} status={r.status_code} {r.text[:200]}')
        except Exception as e:
            print(f'[PDF_VISION_ERR] p{idx} {e}')
    txt='\n'.join(out)
    print(f'[PDF_VISION] pages={len(pages)} chars={len(txt)}')
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
            name = '[확인] 특정암 ' + name       # ★v98 F5: 특정암=일반암 산입 금지(등식1)
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
        if not expiry_date and re.search(r'종신', ht): expiry_date = '9999.12.31'
        if is_excluded(company, product, contract_date, expiry_date): continue   # 제외 5종(단체보험 포함)

        dambo = {}
        for nm, v in _sj_rows(block):
            if re.search(r'납입면제|납입지원', nm): continue
            dambo[nm] = dambo.get(nm, 0) + v
        if not dambo: continue

        # ★확정(3) 갱신판정: 만기 9999(종신)=비갱신 / 상품명 '갱신' 표기=갱신 / 그 외는 후처리 담보절반 규칙
        if _is_silson_prod(company, product): renewal = '갱신'   # ★★★실손=무조건 갱신(v103 영구지침)
        elif str(expiry_date).startswith('9999'): renewal = '비갱신(종신)'
        elif '갱신' in str(product) and '비갱신' not in str(product): renewal = '갱신'
        else: renewal = '비갱신'

        ci_jugye = []
        if _isci_prod(product):
            for bl in block:
                for m2 in re.finditer(r'(?<![_가-힣])주계약\s+([\d,]{3,})', bl):
                    try:
                        v2 = int(m2.group(1).replace(',', ''))
                        if 0 < v2 <= 200000: ci_jugye.append(v2)
                    except: pass
        contracts.append({'company': company, 'product': product, 'contract_date': contract_date,
                          'expiry_date': expiry_date, 'premium': premium, 'pay_period': pay_period,
                          'pay_count': '', 'renewal': renewal, 'dambo': dambo,
                          'ci_jugye': ci_jugye, 'ci_extra': [], 'ipwon': [], '_sj': True})
    print(f'[SINJEONG] 3열 포맷 감지 → 계약 {len(contracts)}건 / 담보 {sum(len(c["dambo"]) for c in contracts)}개')
    return contracts


def parse_txt(txt, filename=''):
    lines = [l.rstrip() for l in txt.replace('\r\n','\n').replace('\r','\n').split('\n')]
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
                if _lks and re.search(r'[가-힣]', _lks) and re.search(r'\s[\d][\d,]*\s*$', _lks) \
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
            _ct = re.sub(r'\d{1,3}\s*/\s*\d{2,3}\s*회',' ',_ct)
            _ct = re.sub(r'[（(]?\s*단위\s*:?\s*만원\s*[）)]?',' ',_ct)
            _ct = re.sub(r'\s+',' ',_ct).strip()
            _mc = re.search(r'^(.*?(?:화재|손해보험|손보|해상|생명|라이프|증권))\s*(.*)$', _ct)
            if _mc:
                company = re.sub(r'\s','',_mc.group(1)); product = _mc.group(2).strip()
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
        if is_excluded(company, product, contract_date, expiry_date):
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
        # ★v29t CI추가보장특약: 줄별 값 수집(병합 전 원값 보존)
        ci_extra=[]
        for _bl in block_lines:
            for _m in re.finditer(r'CI\s*추가보장특약\s+([\d,]+)', _bl):
                try:
                    _v=int(_m.group(1).replace(',',''))
                    if 0<_v<=200000: ci_extra.append(_v)
                except: pass
        # ★v29t 생보 입원특약 일당: 줄별 값 수집(병합 전 원값 보존)
        ipwon=[]
        for _bl in block_lines:
            _m=re.match(r'^\s*입원특약\s+([\d,]+)\s*$', _bl.strip())
            if _m:
                try:
                    _v=int(_m.group(1).replace(',',''))
                    if 0<_v<=1000: ipwon.append(_v)
                except: pass
        # ★v29t (지점장 확정 2026.07.02): 생보 '입원특약' 일당 = 상해·질병 둘 다 해당 →
        #   질병일당·상해일당 두 행에 한 줄 값(중복줄=동일특약 재출현 → max) 각각 기재. dambo 변환이라 엑셀·PPT 동일 반영.
        if ipwon and any(k in (company or '') for k in ('생명','라이프','AIA','메트라이프','우체국','공제')) and '입원특약' in dambo:
            _v1=max(ipwon)
            dambo.pop('입원특약', None)
            dambo['질병입원일당(입원특약)']=dambo.get('질병입원일당(입원특약)',0)+_v1
            dambo['상해입원일당(입원특약)']=dambo.get('상해입원일당(입원특약)',0)+_v1
        if company:
            contracts.append({'company':company,'ipwon':ipwon,'ci_extra':ci_extra,'product':product,'contract_date':contract_date,
                'expiry_date':expiry_date,'premium':premium,'pay_period':pay_period,
                'pay_count':pay_count,'renewal':renewal,'dambo':dambo,'ci_jugye':ci_jugye})
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
            if '심뇌혈관' not in _kk or '수술' not in _kk: continue
            _v = _c['dambo'].pop(_k)
            for _r in ('심장수술비', '뇌혈관수술비'):
                _nk = f'{_r}[묶음]'   # ★태그에 '뇌혈관' 금지(resolve 오인 방지)
                _c['dambo'][_nk] = _c['dambo'].get(_nk, 0) + _v
            print(f"[v51 심뇌혈관수술] {_c.get('company')} '{_k}' {_v} → 심장수술비 + 뇌혈관수술비 (각각 {_v})")

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
        # ★v44: 3열 신정원 포맷은 총회차가 없다 → judge_renewal ④(납입==보장) 적용 금지.
        #        parse_sinjeong이 이미 확정(3) 기준(9999=비갱신 / 상품명 '갱신'=갱신)으로 판정 완료.
        if not c.get('_sj'):
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
    _SEBU_FORCE = ('AIA', '라이나', 'AIG', '우체국')
    try:
        _sb = parse_sebu(lines)
        _nh  = float(_sb.get('뇌혈관진단비', 0) or 0)
        _jol = float(_sb.get('뇌졸증진단비', 0) or 0) or float(_sb.get('뇌졸중진단비', 0) or 0)
        _chu = float(_sb.get('뇌출혈진단비', 0) or 0)
        for _c in deduped:
            _co = re.sub(r'[\s（）()]', '', str(_c.get('company', '')))
            if not any(_f in _co for _f in _SEBU_FORCE):
                continue                                   # 대상 4개사만
            for _k in list(_c['dambo'].keys()):
                _kk = str(_k).replace(' ', '')
                if not (('뇌혈관' in _kk) and ('진단' in _kk)):            continue
                if any(x in _kk for x in ('수술','주요치료','산정특례','혈전','특정')): continue
                if any(x in _kk for x in ('Ⅰ','Ⅱ','Ⅲ','II','III')):        continue
                if not _sb:                                # ★세부가입현황 파싱 실패 → 추측 금지
                    _c['dambo']['[확인] 뇌혈관 축 판별불가(세부가입현황 미파싱) ' + str(_k)] = _c['dambo'].get(_k, 0)
                    print(f"[v43 확인큐] {_co} '{_k}' 세부가입현황 미파싱 → 수기확인")
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
    if str(raw).startswith('[확인]'): return None, 0   # ★v98 확인큐 항목은 표준행 매핑 금지(중복합산 차단)
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
    # ── 수술비 ──
    if has('수술'):
        if has('상해') and jong: return '상해 종수술비(1-5종)', jong
        if has('질병') and jong: return '질병 종수술비(1-5종)', jong
        if has('중대한','상해'): return '중대한상해수술비',0
        if has('5대기관') and has('비관혈'): return '5대기관 수술비 비관혈',0
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
    # ── 암 치료비 ── (지점장 2026.07.09 최종확정: '암주요치료비' 명시 > 하이클래스 > 유사암무시)
    if has('유사암') and has('주요치료'): return '__무시__',0   # ①유사암 주요치료비=무시(엑셀·PPT·설명지 전부)
    if has('암주요치료비') and no('유사암'): return '암주요치료비',0   # ★②담보명에 '암주요치료비' 있으면 하이클래스보다 우선→암주요치료비행 (하이클래스 암주요치료비형)
    if has('하이클래스'): return '하이클래스(암)',0   # ③암주요치료비 없는 하이클래스(항암약물형 등)→하이클래스(암)행. 2건이면 합산
    if has('주요치료') and no('순환계','2대','뇌','허혈','심장','심근','유사암','하이클래스'):   # ③하이클래스 없는 '병원+암주요치료비'→암주요치료비행. ★심장 추가(심장/순환계 주요치료비=2대주요치료비로, v38d)
        return '암주요치료비',0
    if has('고액항암') or (has('고액') and has('항암') and has('치료')): return '__무시__',0   # ★v30z 고액항암치료비=표적+양성자+세기조절+카티 합계값 → 무시(구성 치료비는 아래에서 각각 개별 매핑)
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
    if has('갑상선암') and has('진단') and no('주요치료','수술','일당'): return '유사암(갑.기.경.제)',0   # ★v30a 갑상선암(통합·초기·중증)=유사암(소액암)
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
    if has('고액암'): return '고액암',0
    # ★v29q-2 '암입원일당(…유사암들…)' = 암입원일당 키워드 우선 → 암일당 (유사암 판정보다 먼저)
    if ('암입원일당' in n) or (has('암') and has('입원일당')): return '암일당',0
    # ★v29q-1 '암진단비(…유사암들…)' = 암진단비 키워드 우선 → 일반암 (괄호 유사암 구성 무시)
    if has('소아암') and no('제외'): return None,0   # ★v30q 다발성소아암 등 = 일반암과 별개 담보 → [확인](합산 금지, 지점장 2026.07.03)
    if re.search(r'암\s*진단비\s*[(（]', r) and no('유사암제외'): return '일반암',0
    # 유사암 — 단 '유사암제외'(유사암을 뺀 일반 암진단)는 일반암
    if any(k in n for k in [_norm(x) for x in ['유사암','소액암','갑상선','경계성','제자리','기타피부','양성뇌종양']]) and no('유사암제외','유사암 제외'):
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
    if has('간병인') and has('요양병원') and no('제외'): return None,0  # ★요양병원 포함형 미기재(지점장)
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
    if has('질병') and has('종합') and has('일당'): return '질병종합병원일당',0
    # ★v30k 교통상해입원일당 ≠ 상해입원일당(합산 금지). 질환·부위·교통 접두 변형은 base 아님 → [확인]
    # ★병원규모(상급종합/종합) 명시 = 개별 전용행 / 일반 질병입원일당(밴드) = 합산 (지점장 2026.07.05)
    _dilqual = ('교통','암','뇌','심','허혈','간','신장','폐','위','골절','화상','특정','재해외','종합','요양','중환자','수술')
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
    """★정본 §8.8 실손 후처리 (2026.07.05 지점장): 손보/생보 구분 + 입원=명시값 최우선(1세대 1억·5천·3천 다 유지) + 통원·약값 규칙.
    우선순위 ①보장표 명시값 ②입원한도 3000짜리 구형=통원10·약5(입원 3000유지) ③회사유형(손보 통원25·약5 / 생보 통원20·약10) ④4세대(통원20·약0)."""
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
        _gen=silson_gen(cd, ipw, prod)
        if _gen in ('4세대','5세대') or (_ym and _ym>=202107):
            _set('통원', _tw if (_tw and _tw<=20) else 20)
            _set('약값', 0)
            continue
        # ★v41 정본(지점장 2026.07.12): 1세대 = 통원 한도에 약제비 포함 → 약값 행 미표기
        if _gen == '1세대':
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

    for i, ct in enumerate(contracts):
        col = 3 + i
        gen  = ct['renewal'] == '갱신'
        paid = '완납' in ct['renewal']
        h = ws.cell(1, col)
        h.value = f"{ct['company']}\n{ct['product']}\n[{ct['renewal']}]"
        h.font = W; h.alignment = AL
        h.fill = FILL_GREEN if paid else (FILL_BLUE if gen else FILL_RED)
        pm = ct['premium']
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
        _is_ci = _isci_prod(ct.get('product'))
        _cij = ct.get('ci_jugye') or []
        if _is_ci and _cij:
            _samang = max(_cij); _bonche=None; _pct=None
            _cand=[v for v in _cij if 0<v<_samang]
            for _ratio,_p in ((0.8,80),(0.5,50)):
                for v in sorted(set(_cand), key=lambda x:-_cand.count(x)):
                    if _samang and abs(v/_samang-_ratio)<0.02: _bonche=v; _pct=_p; break
                if _bonche: break
            if _bonche:
                for _nm in ('중대한 암','중대한 뇌졸증','중대한 급성심근'):
                    _r=nm2r.get(_nm)
                    if _r:
                        ws.cell(_r,col).value=_bonche; ws.cell(_r,col).font = BL if gen else BK
                # ★v29t 등식1: CI 사망은 dambo 합성 키로 주입 → 엑셀·PPT 분할합 동일 (셀 직접기재 폐기)
                dambo['일반사망(종신주계약)']=dambo.get('일반사망(종신주계약)',0)+_samang
                dambo['상해사망(종신주계약)']=dambo.get('상해사망(종신주계약)',0)+_samang   # §8.1 종신 1:1, 재해특약은 별도 합산
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
            if '진단' in _rn and '수술' not in _rn and '주요치료' not in _rn:
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
                        # ★★단독담보 원칙(지점장 확정 2026.07.14, 최상위): 회사담보명이 '허혈성심장질환진단비'
                        #   단독이면 어느 회사든 분해 금지 → '허혈성 진단비' 행 단독. 묶음 수식어가 붙은 것만 분해.
                        _solo = bool(re.match(r'^허혈(성)?심장질환진단(비)?$', _rmn0 if False else re.sub(r'\s','',_rn)))
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
            target_rows = nm2r_multi.get(std, [r]) if std == '2대 주요치료비' else [r]
            for tr in target_rows:
                existing = ws.cell(tr,col).value
                _rep1 = std in ('표적항암치료비','다빈치로봇수술비','n대수술비','입원','통원','약값','약','간병인','창상봉합술','항암방사선약물','중입자치료비','암주요치료비','통합전이암','간호통합병동')
                _rep1 = _rep1 or ('통합' in raw and std in ('일반암','유사암(갑.기.경.제)','통합전이암'))   # ★v30a §8.2 통합 계열=대표금액 1개
                if _rep1 and isinstance(existing,(int,float)):
                    ws.cell(tr,col).value = max(existing, amt)   # 표적·n대·창상봉합=대표 최댓값1건(★v29q-6) / 실손=중복합산 안함(한도)
                else:
                    ws.cell(tr,col).value = (existing+amt) if isinstance(existing,(int,float)) else amt
                # 실손(입원/통원/약값)은 갱신·비갱신 무관 항상 파랑
                ws.cell(tr,col).font = BL if (blue or std in ('입원','통원','약값','약','간병인','간병인지원일당','일상배상책임')) else BK   # ★v29w 실손·간병인·일배책 항상 파랑(§10)
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

        # ★ §8 생보 종신(만기 9999): 일반사망(종신) + 상해사망 1:1 복제
        if ct['expiry_date'].startswith('9999'):
            r_il = nm2r.get('일반사망'); r_sh = nm2r.get('상해사망')
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
            _g4=(silson_gen(ct.get('contract_date',''), None, ct.get('product','')) in ('4세대','5세대'))   # ★v29v 상품코드 반영
            _guhy=(_ipv==3000)                            # 입원한도 3,000=구형
            _twc=ws.cell(_rtw,col).value if _rtw else None
            _ykc=ws.cell(_ryk,col).value if _ryk else None
            if _rtw and not isinstance(_twc,(int,float)):  # ① 별첨 통원 없을 때만 디폴트
                _twd = 10 if _guhy else (20 if _g4 else (20 if _life else 25))
                ws.cell(_rtw,col).value=_twd; ws.cell(_rtw,col).font=BL
            _g1=(silson_gen(ct.get('contract_date',''), _ipv, ct.get('product','')) == '1세대')   # ★v41
            if _ryk and not isinstance(_ykc,(int,float)):  # ① 별첨 약값 없을 때만 디폴트
                _ykd = 0 if (_g4 or _g1) else (5 if _guhy else (10 if _life else 5))   # ★v41 1세대=약값 통원포함
                if _ykd: ws.cell(_ryk,col).value=_ykd; ws.cell(_ryk,col).font=BL   # 4세대 약0=미기재
            # ★ 실손 세대 자동판별 → 헤더에 라벨 기재
            _sg = silson_gen(ct.get('contract_date',''), _ipv, ct.get('product',''))
            _pm0=re.search(r'(?<!\d)(0[9]|1[0-9]|2[0-6])(0[1-9]|1[0-2])(?!\d)', str(ct.get('product','')))
            silson_trace.append((ct['company'], ct.get('contract_date',''), (_pm0.group(0) if _pm0 else '없음'), _sg or '판정불가'))   # ★v29z 세대 근거
            if _sg:
                _hc = ws.cell(1,col)
                if _hc.value and _sg not in str(_hc.value):
                    _hc.value = str(_hc.value) + f'\n({_sg} 실손)'

    # ★v29t (지점장 확정 2026.07.02): CI 존재 시 '중대한CI적용' 행 = CI 잔여액 + 비CI 계약의 일반사망 동일액 —
    #   CI 적용/미적용 각각의 총 사망액이 양쪽 행에서 가로합산되도록.
    _rci_all=None; _ril_all=None
    for _rr in range(6, ws.max_row+1):
        _b=str(ws.cell(_rr,2).value or '').strip()
        if _b=='중대한CI적용': _rci_all=_rr
        if _b=='일반사망': _ril_all=_rr
    _has_ci=any(_isci_prod(c.get('product')) for c in contracts)
    if _has_ci and _rci_all and _ril_all:
        for _ix,_c in enumerate(contracts):
            _cl=3+_ix
            _isci=_isci_prod(_c.get('product'))
            _ilv=ws.cell(_ril_all,_cl).value
            if (not _isci) and isinstance(_ilv,(int,float)) and not isinstance(ws.cell(_rci_all,_cl).value,(int,float)):
                ws.cell(_rci_all,_cl).value=_ilv
                ws.cell(_rci_all,_cl).font=ws.cell(_ril_all,_cl).font.copy()

    # ★ 합계 = 항상 표 맨 끝 열. 가로 SUM 수식(법칙22, 하드코딩 금지).
    last_col = 3 + n_ct
    # ★v30q 유사암 자동유도(지점장 2026.07.03): 계약에 일반암 있고 유사암 담보가 따로 없으면 유사암 = 그 일반암 × 10%
    _r일반암 = nm2r.get('일반암'); _r유사암 = nm2r.get('유사암(갑.기.경.제)')
    if _r일반암 and _r유사암:
        for _c in range(3, last_col):
            _v일 = ws.cell(_r일반암, _c).value
            _v유 = ws.cell(_r유사암, _c).value
            if isinstance(_v일,(int,float)) and _v일 > 0 and not isinstance(_v유,(int,float)):
                ws.cell(_r유사암, _c).value = round(_v일 * 0.1)
                try: ws.cell(_r유사암, _c).font = _copy.copy(ws.cell(_r일반암, _c).font)   # 일반암 색(갱신/비갱신) 따라감
                except: pass

    first_L = get_column_letter(3)
    last_ct_L = get_column_letter(last_col-1) if n_ct>0 else first_L
    hc = ws.cell(1, last_col)
    hc.value = '합계'; hc.font = W; hc.fill = FILL_SUM; hc.alignment = AL
    # 보험료 합계 = 숫자만 표기(§3): 수식 아닌 계산된 숫자값. 글자 검정(흰바탕)
    if n_ct>0:
        ws.cell(2, last_col).value = f'=SUM(C2:{last_ct_L}2)'   # ★v29t §5: 보험료 합계도 동적 SUM
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
    ws2.cell(6,1,'[확인] 자동매핑 실패 담보 (마스터 미수록 또는 약관 확인 후 수기 기재)')
    ws2.cell(7,1,'회사'); ws2.cell(7,2,'담보명'); ws2.cell(7,3,'금액(만원)'); ws2.cell(7,4,'보장범위(참고)'); ws2.cell(7,5,'약관검색')
    LINKF = Font(color='0000FF', underline='single')
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
    return unmapped

def read_excel_totals(path):
    """완성 엑셀에서 담보명->합계 읽음. 등식2: PPT는 이것만 본다.
       끝열 =SUM() 캐시(LibreOffice 의존) 대신 데이터셀(C~끝열-1) 직접 합산 → 재계산 없어도 PPT=엑셀 보장."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['보장분석']; last = ws.max_column
    out = {}; sq=[0]*5; ss=[0]*5
    for r in range(6, ws.max_row+1):
        nm = ws.cell(r,2).value
        if not nm: continue
        nm = str(nm).strip()
        endv = ws.cell(r,last).value
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
        # 숫자 합계: 끝열 캐시값 있으면 사용, 없으면(=SUM 미계산) 데이터셀 C~끝열-1 직접 합산
        if isinstance(endv,(int,float)) and endv:
            out[nm] = endv
        else:
            s = 0
            for c in range(3, last):
                v = ws.cell(r,c).value
                if isinstance(v,(int,float)): s += v
            if s: out[nm] = s
    return out, sq, ss

def build_ppt(data, out, totals=None, surg_q=None, surg_s=None):
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
    _silson={'입원','통원','약값','약','MRI','도수치료','비급여주사','간병인','일상배상책임'}  # 간병인·일상배상책임=무조건 파랑
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
    _gensum={}; _nonsum={}
    for ct in contracts:
        _gen=(ct.get('renewal','')=='갱신')
        for raw,amt in ct.get('dambo',{}).items():
            if not amt: continue
            st=resolve(raw)
            if not st: continue
            tgt=_gensum if (_gen or '갱신' in raw) else _nonsum   # ★v29t: 담보명 (갱신형) = 파랑(엑셀 792행과 동일 기준)
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
        if box not in by: return
        tf=by[box].text_frame
        if pi>=len(tf.paragraphs): return
        p=tf.paragraphs[pi]
        if ri>=len(p.runs): return
        gs=_gensum.get(std,0); ns=_nonsum.get(std,0)
        if std in _silson:
            _v = totals.get(std,0)            # ★실손=완성 엑셀값(입원5천캡·통원디폴트 반영). _gensum 원본합산(상해+질병=1만) 사용 안 함
            if not _v: return
            segs=[(f'{prefix}{_v:,}{suffix}', _BLUE)]
            _seg(p.runs[ri], segs); return
        # ★v30f 등식1 (지점장 승인 2026.07.03): PPT 표기 총액의 정본 = 완성 엑셀 끝열.
        #   대표 1건·MAX·캡·[확인] 분리로 끝열 ≠ raw 분할합이면 끝열 값 단색 표기(색=우세 성분측).
        _T = totals.get(std, None)
        if isinstance(_T,(int,float)) and _T>0 and (gs+ns)!=_T:
            if gs>=ns: gs,ns=int(_T),0
            else: gs,ns=0,int(_T)
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
        if box in by:
            tf=by[box].text_frame
            if pi<len(tf.paragraphs):
                p=tf.paragraphs[pi]
                if ri<len(p.runs):
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
    _sg=silson_gen(실손가입일, totals.get('입원'), _실손상품)   # ★실손 세대 자동판별(상품명 연도코드 반영)
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
def health(): return {'ok':True,'version':'v112-8k-20260720'}

# ★★v101 진단 엔드포인트(2026.07.20): 폰에서 링크 한 번만 눌러
#   Railway 컨테이너에 pdftotext(poppler)가 실제로 살아있는지 확인한다.
#   'KB(PDF)는 죽고 롯데(txt)는 산다'의 원인을 서버 로그 없이 확정하기 위함.
@app.get('/diag')
def diag():
    import subprocess, shutil
    out = {'version': 'v112-8k-20260720'}
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
    _pdf_data=None; pdf_txt=''; _img_pdf_nokey=False
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
            if (not _rawtl or len(_rawtl.strip())<30) and not os.environ.get('ANTHROPIC_API_KEY',''):
                _img_pdf_nokey=True
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
                return JSONResponse({'ok':False,'source':'이미지PDF',
                    'error':'이미지 PDF입니다 — OCR.PDF로 변환해서 올리세요'})
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
        ppt_totals, sq, ss = read_excel_totals(xl)   # 등식2: PPT는 완성 엑셀만 읽음
        ppt_ok=build_ppt(data,pt,ppt_totals,sq,ss)
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
            rep=map_excel_to_report(xl, settings={'client':cust,
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
