# -*- coding: utf-8 -*-
import os, re, tempfile, datetime, base64, traceback, json, httpx, urllib.parse
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Pt
 
app = FastAPI(title="BARUM 보장분석 v7")
PW   = os.environ.get("ACCESS_PW", os.environ.get("BARUM_PW", "1009"))
HERE = os.path.dirname(os.path.abspath(__file__))
TPL_XL  = os.path.join(HERE, "MASTER_보장분석_엑셀_영구표본.xlsx")
TPL_PPT = os.path.join(HERE, "MASTER_보장분석지_PPT_빈폼.pptx")
 
W   = Font(color='FFFFFF', name='맑은 고딕', size=9, bold=True)
BL  = Font(color='0070C0', name='맑은 고딕', size=9)
BK  = Font(color='000000', name='맑은 고딕', size=9)
FILL_RED   = PatternFill('solid', fgColor='C00000')
FILL_BLUE  = PatternFill('solid', fgColor='0070C0')
FILL_GREEN = PatternFill('solid', fgColor='375623')
FILL_SUM   = PatternFill('solid', fgColor='2E75B6')
AL = Alignment(horizontal='center', vertical='center', wrap_text=True)
 
EXCLUDE = ['실효','미납해지','농업인','자동차보험']  # NH농협=포함, '농업인' 표기만 제외
def is_excluded(company, product=''):
    return any(kw in company+product for kw in EXCLUDE)
 
def judge_renewal(product, expiry, pay_count, contract='', pay_period=''):
    # 지침 §7 판정 순서
    # 1) '갱신형' 명시 -> 갱신
    if '갱신형' in product and '비갱신' not in product: return '갱신'
    if '갱신' in product and '비갱신' not in product: return '갱신'
    # 2) 만기 9999(종신) -> 비갱신
    if expiry.startswith('9999'): return '비갱신(종신)'
    # 3) 총회차 240 초과 -> 비갱신
    try:
        _, b = pay_count.split('/')
        if int(b.strip()) > 240: return '비갱신'
    except: pass
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
 
def get_종번호(name):
    for i,k in enumerate(['(1종)','(2종)','(3종)','(4종)','(5종)'],1):
        if k in name: return i
    return 0
 
def rule_extract(block_lines):
    """기존 규칙 추출(담보명+금액 같은 줄). 폴백용."""
    dambo = {}
    UNIT = r'(?:\s*(원|만원|만))?'
    for l in block_lines:
        l = l.strip()
        m = re.search(r'^(.+?)\s{2,}([\d,]+)' + UNIT + r'\s*$', l) or re.search(r'^(.+?)\s+([\d,]+)' + UNIT + r'\s*$', l)
        if m:
            name = re.sub(r'\s+', ' ', m.group(1).strip())
            try:
                amt = int(m.group(2).replace(',',''))
                if (m.group(3) or '') == '원': amt = amt // 10000
                if 0 < amt <= 200000 and len(name) > 2:
                    dambo[name] = dambo.get(name,0) + amt
            except: pass
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
            json={'model':'claude-sonnet-4-6','max_tokens':4000,'messages':[{'role':'user','content':prompt}]}, timeout=90)
        txt = ''.join(b.get('text','') for b in r.json().get('content',[]) if b.get('type')=='text')
        txt = txt.strip().replace('```json','').replace('```','').strip()
        out = json.loads(txt)
        return {str(k).strip(): int(v) for k,v in out.items() if isinstance(v,(int,float)) and 0 < v <= 200000 and len(str(k).strip())>2}
    except Exception:
        return {}
 
def parse_txt(txt, filename=''):
    lines = [l.rstrip() for l in txt.replace('\r\n','\n').replace('\r','\n').split('\n')]
    client = ''
    # ★ 정본 §2: 고객명 = 파일명 우선
    if filename:
        base = re.sub(r'\.[Tt][Xx][Tt]$', '', filename).strip()
        fm = re.match(r'^([가-힣]{2,4})', base)
        if fm: client = fm.group(1)
    # 폴백: 내용에서 (마스킹 '박*은' 형태도 허용)
    if not client:
        for l in lines[:30]:
            l = l.strip()
            m2 = re.search(r'([가-힣]{2,4})\s+고객님', l)
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
    while i < n:
        l = lines[i].strip()
        if '실효계약 리스트' in l or '미납해지' in l: break
        if '정상계약 리스트' not in l: i += 1; continue
        i += 1
        while i < n and not lines[i].strip(): i += 1
        if i >= n: break
        company = lines[i].strip(); i += 1
        if is_excluded(company):
            while i < n and '정상계약 리스트' not in lines[i] and '실효계약 리스트' not in lines[i]: i += 1
            continue
        contract_date = expiry_date = pay_period = pay_count = ''; premium = 0
        for j in range(i, min(i+5, n)):
            l = lines[j]
            m = re.search(r'(\d{4}\.\d{2}\.\d{2})\s*~\s*(\d{4}\.\d{2}\.\d{2})', l)
            if m: contract_date = m.group(1); expiry_date = m.group(2)
            m2 = re.search(r'(\d{1,3})\s*/\s*(\d{2,3})\s*회', l) or re.search(r'월납\s+(\d{1,3})\s*/\s*(\d{2,3})', l) or re.search(r'(?<![\d.])(\d{1,3})\s*/\s*(\d{2,3})(?![\d.])', l)
            if m2 and not pay_count: pay_count = f"{m2.group(1)}/{m2.group(2)}"
            m3 = re.search(r'([\d,]+)원', l)
            if m3:
                v = int(m3.group(1).replace(',',''))
                if 1000 < v < 5000000: premium = v
            m4 = re.search(r'(\d+)년납', l)
            if m4: pay_period = f"{m4.group(1)}년납"
            m5 = re.search(r'월납\s*/\s*(\d+)년', l)
            if m5 and not pay_period: pay_period = f"{m5.group(1)}년납"
        while i < n and not lines[i].strip(): i += 1
        product = ''
        for j in range(i, min(i+6, n)):
            l = lines[j].strip()
            if l and not re.search(r'계약자|납입주기|보험료|보장기간', l):
                if len(l) > 5 and not re.search(r'^\d+$', l) and not re.search(r'^\d{4}\.\d{2}', l):
                    product = l; i = j + 1; break
        renewal = judge_renewal(product, expiry_date, pay_count, contract_date, pay_period)
        # 담보 블록 텍스트 수집 (다음 '정상계약/실효계약 리스트'까지)
        block_lines = []; j = i
        while j < n:
            if '정상계약 리스트' in lines[j] or '실효계약 리스트' in lines[j]: break
            block_lines.append(lines[j]); j += 1
        i = j
        # 추출: LLM 우선(깨진 별첨 복원), 키 없거나 실패 시 규칙 폴백
        dambo = llm_extract('\n'.join(block_lines)) or rule_extract(block_lines)
        if company:
            contracts.append({'company':company,'product':product,'contract_date':contract_date,
                'expiry_date':expiry_date,'premium':premium,'pay_period':pay_period,
                'pay_count':pay_count,'renewal':renewal,'dambo':dambo})
    # ★ 페이지 분할 중복 제거 (정본 체크리스트 ①②): 동일 계약키 병합
    merged = {}
    order = []
    for c in contracts:
        key = (c['company'], c['contract_date'], c['expiry_date'], c['premium'])
        if key not in merged:
            merged[key] = c; order.append(key)
        else:
            m = merged[key]
            # 담보 병합: 같은 담보명은 큰 값 유지(중복가산 방지), 새 담보는 추가
            for k, v in c['dambo'].items():
                m['dambo'][k] = max(m['dambo'].get(k, 0), v)
            # 더 긴(덜 잘린) 상품명 채택
            if len(c['product']) > len(m['product']): m['product'] = c['product']
            # 회차/기간 비어있으면 채움
            if not m['pay_count'] and c['pay_count']: m['pay_count'] = c['pay_count']
            if not m['pay_period'] and c['pay_period']: m['pay_period'] = c['pay_period']
    deduped = [merged[k] for k in order]
    # 한장보장표 회차 주입 (별첨에 없던 pay_count 보정)
    for c in deduped:
        if not c['pay_count']:
            pc = paycount_map.get((c['company'], c['contract_date'], c['expiry_date'])) \
                 or paycount_map.get((c['contract_date'], c['expiry_date']))
            if pc: c['pay_count'] = pc
    # 병합·회차 보정 반영하여 갱신 재판정 (정본 §7 규칙대로만)
    for c in deduped:
        c['renewal'] = judge_renewal(c['product'], c['expiry_date'], c['pay_count'], c['contract_date'], c['pay_period'])
    return {'client':client,'contracts':deduped}
 
# ★ DMAP — 마스터 엑셀 B열 기준 100% 일치
DMAP = {
    # 사망
    '상해사망':'상해사망','상해사망(갱신형) [보통약관]':'상해사망','일반상해사망':'상해사망',
    '기본계약(상해사망(간편가입Ⅲ))담보':'상해사망',
    '질병사망':'질병사망(80세)',
    # 후유장애
    '상해후유장해3%':'상해후유3%','상해후유80%':'상해후유80%',
    '질병후유장해3%':'질병후유3%','질병후유80%':'질병후유80%',
    # 암 — B열: 고액암/일반암/중대한 암/유사암(갑.기.경.제)/표적항암치료비/하이클래스(암)/중입자치료비/양성자치료/세기조절치료/다빈치로봇수술비/암수술/암일당/항암방사선약물
    '일반암진단비':'일반암','암진단비':'일반암','암진단Ⅱ(유사암제외)(간편가입Ⅲ)담보':'일반암',
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
    '허혈심장질환진단비Ⅲ(건강맞춤형Ⅱ)(갱신형)':'협심증','허혈심장질환진단비':'협심증',
    '허혈심장질환진단(간편가입Ⅲ)담보':'협심증',
    '급성심근경색증진단':'급성심근경색','급성심근경색증진단(간편가입Ⅲ)담보':'급성심근경색',
    '중증질환자(심장질환)산정특례대상진단비(연간1회한)(건강맞춤형Ⅱ)(갱신형)':'산정특례심장',
    '허혈심장질환수술비Ⅲ(건강맞춤형Ⅱ)(갱신형)':'심장수술비','허혈심장질환수술비':'심장수술비',
    '급성심근경색증(I21)혈전용해치료비':'혈전용해치료비',
    # 일당 — B열: 질병일당/질병수술일당/질병종합병원일당/상해일당/간병인/간호통합병동/1인실 종합병원/1인실 상급병원/질병중환자실/상해중환자실
    '간병인사용질병입원일당(1일이상)(요양병원)(간편가입)(갱신형)':'간병인',
    '간호간병통합서비스질병입원일당(1-180일)(간편가입)(갱신형)':'간호통합병동',
    '상급종합병원질병입원일당(상급병실(1인실),1일이상60일한도)(간편가입)(갱신형)':'1인실 상급병원',
    '종합병원질병입원일당(상급병실(1인실),1일이상30일한도)(간편가입)(갱신형)':'1인실 종합병원',
    # 수술비
    '질병수술비':'질병수술비','질병수술(간편가입Ⅲ)담보':'질병수술비',
    '상해수술비(건강맞춤형Ⅱ)(갱신형)':'상해수술비','상해수술(간편가입Ⅲ)담보':'상해수술비',
    '골절수술(간편가입Ⅲ)담보':'골절수술비','화상수술(간편가입Ⅲ)담보':'화상수술비',
    '120대질병수술Ⅱ(간편가입Ⅲ)(질병수술3(24대질병))담보':'120대수술비',
    '5대기관질병수술(관혈/비관혈)(연간1회한)(간편가입Ⅲ)담보':'5대기관 수술비 관혈',
    '중대한특정상해수술(간편가입Ⅲ)담보':'중대한상해수술비',
    # 운전자 — B열: 대인/대물/합의금/6주미만/변호사/자부상
    '교통사고처리지원금':'대인','교통사고 처리지원금':'대인',
    '교통사고벌금(대물)':'대물','교통사고벌금(대인)':'대인',
    '6주미만처리지원금':'6주미만',
    '변호사선임비용':'변호사','자동차사고 변호사선임비용':'변호사',
    '자동차부상위로금':'자부상',
    '무보험차에 의한 상해':'일상배상책임',
    # 골절 — B열: 골절(치아파절포함)/골절(치아파절제외)/5대골절진단비
    '골절진단(간편가입Ⅲ)담보':'골절(치아파절제외)','골절진단비':'골절(치아파절제외)',
    # 응급실
    '응급실내원비(응급)':'응급실(응급)',
    # 화상 — B열: 진 단 비/중증화상진단비
    '화상진단비':' 진 단 비','화상진단비(건강맞춤형Ⅱ)(갱신형)':' 진 단 비',
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
 
def resolve(raw):
    if raw in DMAP: return DMAP[raw]
    for k,v in DMAP.items():
        if k in raw: return v
    if any(k in raw for k in ['(1종)','(2종)','(3종)','(4종)','(5종)']): return None
    # 암 일반 폴백: '암'+'진단' 포함 & 특수암/치료 아님 -> 일반암
    if '암' in raw and '진단' in raw:
        if not any(x in raw for x in ['유사','고액','소액','표적','방사선','약물','수술','일당',
                                      '양성자','세기','중입자','전이','보험료']):
            return '일반암'
    return None
 
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
        "- 라이나(라이나생명) '뇌혈관진단비'는 실제 뇌출혈 보장 → 뇌출혈진단비\n"
        "- 하이클래스=하이클래스(암), 표적항암약물허가=표적항암치료비(여러 건이면 가장 큰 1건만)\n"
        "- 유사암 진단금액은 가입연도 2020년 이하면 일반암의 1/10, 2021년 이상이면 1/5로 환산 기재\n"
        "- 뇌혈관수술=뇌혈관수술비, 항암방사선/약물치료비=항암방사선약물, 암수술=암수술\n"
        "- 화상 진단비='진 단 비'(화상 구분 행), 중증화상=중증화상진단비\n"
        "- 생명보험 종신 주계약/기본계약(사망보장)=일반사망\n"
        "- 운전자: 교통사고처리지원금=대인, 벌금(대물)=대물, 변호사선임=변호사, 자동차부상위로금=자부상\n"
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
            json={'model':'claude-sonnet-4-6','max_tokens':8000,
                  'messages':[{'role':'user','content':prompt}]}, timeout=90)
        txt=''.join(b.get('text','') for b in r.json().get('content',[]) if b.get('type')=='text')
        txt=txt.strip().replace('```json','').replace('```','').strip()
        out=json.loads(txt)
        return {k:{'std':(v.get('std') if isinstance(v,dict) else None),
                   'jong':(v.get('jong',0) if isinstance(v,dict) else 0),
                   'note':(v.get('note','') if isinstance(v,dict) else '')} for k,v in out.items()}
    except Exception:
        return {r:{'std':None,'jong':0,'note':''} for r in raw_names}
 
def build_excel(data, out):
    wb = openpyxl.load_workbook(TPL_XL)
    ws = wb['보장분석']
    client = data['client']; contracts = data['contracts']
 
    # 담보명 -> 행번호 맵 (A/B열 유지)
    nm2r = {}
    nm2r_norm = {}   # 공백무시 보조키 ('진 단 비' 등 라벨 변형 흡수)
    for r in range(6, ws.max_row+1):
        v = ws.cell(r,2).value
        if v:
            k = str(v).strip()
            nm2r[k] = r
            nm2r_norm[re.sub(r'\s','',k)] = r
 
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
 
    for i, ct in enumerate(contracts):
        col = 3 + i
        gen  = ct['renewal'] == '갱신'
        paid = '완납' in ct['renewal']
        h = ws.cell(1, col)
        h.value = f"{ct['company']}\n[{ct['renewal']}]"
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
        jong_acc = {'상해 종수술비(1-5종)':[0]*5, '질병 종수술비(1-5종)':[0]*5}
        jong_blue = {'상해 종수술비(1-5종)':False, '질병 종수술비(1-5종)':False}
 
        for raw, amt in dambo.items():
            m = LLMMAP.get(raw) or {}
            std = m.get('std'); jong = m.get('jong', 0) or 0
            if not std:                       # LLM 미반환 시 사전 폴백
                std = resolve(raw)
                if not jong: jong = get_종번호(raw)
            blue = gen or ('갱신' in raw)      # ★ 담보명에 (갱신) 표시 -> 파랑
            # 수술비 1~5종 -> 종별 슬래시 누적
            if std in jong_acc and 1 <= jong <= 5:
                jong_acc[std][jong-1] += amt
                if blue: jong_blue[std] = True
                continue
            r = nm2r.get(std)
            if r is None and std:             # 공백무시 재매칭 (화상 '진 단 비' 등)
                r = nm2r_norm.get(re.sub(r'\s','', std))
            if not std or r is None:          # 마스터 미수록/매핑실패 -> [확인]
                unmapped.append((col, ct['company'], raw, amt, m.get('note','') or ''))
                continue
            existing = ws.cell(r,col).value
            if std == '표적항암치료비' and isinstance(existing,(int,float)):
                ws.cell(r,col).value = max(existing, amt)   # §8 표적=최댓값 1건
            else:
                ws.cell(r,col).value = (existing+amt) if isinstance(existing,(int,float)) else amt
            ws.cell(r,col).font = BL if blue else BK
 
        for nm, vals in jong_acc.items():     # 종수술비 슬래시 기재(§6)
            if any(vals):
                r = nm2r.get(nm)
                if r:
                    ws.cell(r,col).value = '/'.join(str(x) for x in vals)
                    ws.cell(r,col).font = BL if (gen or jong_blue[nm]) else BK
 
        # ★ §8 생보 종신(만기 9999): 일반사망(종신) + 상해사망 1:1 복제
        if ct['expiry_date'].startswith('9999'):
            r_il = nm2r.get('일반사망'); r_sh = nm2r.get('상해사망')
            v = ws.cell(r_il,col).value if r_il else None
            if isinstance(v,(int,float)) and r_sh and not isinstance(ws.cell(r_sh,col).value,(int,float)):
                ws.cell(r_sh,col).value = v
                ws.cell(r_sh,col).font = BL if gen else BK
 
    # ★ 합계 = 항상 표 맨 끝 열. 가로 SUM 수식(법칙22, 하드코딩 금지).
    last_col = 3 + n_ct
    first_L = get_column_letter(3)
    last_ct_L = get_column_letter(last_col-1) if n_ct>0 else first_L
    hc = ws.cell(1, last_col)
    hc.value = '합계'; hc.font = W; hc.fill = FILL_SUM; hc.alignment = AL
    # 보험료 합계 = 숫자만 표기(§3): 수식 아닌 계산된 숫자값. 글자 검정(흰바탕)
    if n_ct>0:
        ws.cell(2, last_col).value = sum(c['premium'] for c in contracts)
        ws.cell(2, last_col).font = BK
 
    for r in range(6, ws.max_row+1):
        slash_t=[0]*5; is_slash=False; has_num=False
        for col in range(3, last_col):
            v = ws.cell(r,col).value
            if isinstance(v,(int,float)): has_num=True
            elif isinstance(v,str) and '/' in v:
                is_slash = True
                for k,p in enumerate(v.split('/')[:5]):
                    try: slash_t[k] += int(p)
                    except: pass
        sc = ws.cell(r, last_col)
        if is_slash and any(slash_t):
            sc.value = '/'.join(str(x) for x in slash_t); sc.font = BK   # 슬래시 행은 §3 SUM 예외
        elif has_num:
            sc.value = f'=SUM({first_L}{r}:{last_ct_L}{r})'; sc.font = BK
 
    ws.column_dimensions['B'].width = 22
    for c in range(3, last_col+1):
        ws.column_dimensions[get_column_letter(c)].width = 12
 
    # ★ 합계 이후 잔재 열 삭제 (§3: 합계 = 맨 끝 열)
    if ws.max_column > last_col:
        ws.delete_cols(last_col+1, ws.max_column - last_col)
 
    # ── 확인사항 시트: LLM 매핑 실패 담보 노출(자가진단, §10) ──
    if '📋확인사항' in wb.sheetnames: del wb['📋확인사항']
    ws2 = wb.create_sheet('📋확인사항')
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
        cell = ws2.cell(rr,5,'🔗약관검색')
        cell.hyperlink = "https://search.naver.com/search.naver?query=" + urllib.parse.quote(q)
        cell.font = LINKF
    ws2.column_dimensions['B'].width = 34; ws2.column_dimensions['D'].width = 40; ws2.column_dimensions['E'].width = 12
    wb.save(out)
 
def build_ppt(data, out):
    if not os.path.exists(TPL_PPT): return False
    prs = Presentation(TPL_PPT)
    sl = prs.slides[0]
    by = {sh.name:sh for sh in sl.shapes if sh.has_text_frame}
    client = data['client']; contracts = data['contracts']
    now = datetime.datetime.now()
 
    totals = {}
    surg_q=[0]*5; surg_s=[0]*5
    for ct in contracts:
        for raw,amt in ct['dambo'].items():
            i = get_종번호(raw)-1
            if i>=0:
                if '질병수술비' in raw: surg_q[i]+=amt
                if '상해수술비' in raw: surg_s[i]+=amt
            std = resolve(raw)
            if std: totals[std]=totals.get(std,0)+amt
 
    def g(nm): return totals.get(nm,0)
    def r_set(box,pi,ri,val):
        if box not in by: return
        tf=by[box].text_frame
        if pi<len(tf.paragraphs):
            p=tf.paragraphs[pi]
            if ri<len(p.runs): p.runs[ri].text=val
 
    for b in ['TextBox 49','TextBox 56']:
        if b in by: by[b].text_frame.word_wrap=False
 
    by['TextBox 21'].text_frame.word_wrap=False
    by['TextBox 21'].text_frame.paragraphs[0].runs[0].text=f'{client} 님의 보장'
    by['TextBox 21'].text_frame.paragraphs[0].runs[1].text='(전)'
    by['TextBox 36'].text_frame.paragraphs[0].runs[0].text=f'{now.year}년'
    by['TextBox 35'].text_frame.paragraphs[0].runs[0].text=f'{now.month:02d}월'
    by['TextBox 29'].text_frame.paragraphs[0].runs[0].text=f'{now.day:02d}일 기준'
 
    if g('질병사망(80세)'): r_set('TextBox 10',2,2,f': {g("질병사망(80세)"):,}')
    if g('상해사망'): r_set('TextBox 11',0,1,f': {g("상해사망"):,}')
    종신_d=0
    for ct in contracts:
        if '종신' in ct['renewal']:
            for raw,v in ct['dambo'].items():
                if resolve(raw)=='상해사망': 종신_d+=v
    if 종신_d: r_set('TextBox 10',3,1,f': {종신_d:,}')
 
    if g('상해후유3%'): r_set('TextBox 8',2,1,f'3% : {g("상해후유3%"):,}')
    if g('질병후유3%'): r_set('TextBox 8',0,1,f'3% : {g("질병후유3%"):,}')
    if g('상해후유80%'): r_set('TextBox 8',3,1,f'80% : {g("상해후유80%"):,}')
    if g('질병후유80%'): r_set('TextBox 8',1,1,f'80% : {g("질병후유80%"):,}')
 
    if g('뇌혈관진단비'): by['TextBox 46'].text_frame.paragraphs[0].runs[0].text=f'뇌혈관\n{g("뇌혈관진단비"):,}'
    if g('뇌졸증진단비'): by['TextBox 47'].text_frame.paragraphs[0].runs[0].text=f'뇌졸증\n{g("뇌졸증진단비"):,}'
    if g('뇌출혈진단비'): by['TextBox 48'].text_frame.paragraphs[0].runs[0].text=f'뇌출혈\n{g("뇌출혈진단비"):,}'
    if g('산정특례뇌혈관'): r_set('TextBox 49',0,3,f': {g("산정특례뇌혈관"):,}')
    if g('혈전용해치료비'): r_set('TextBox 49',1,1,f': {g("혈전용해치료비"):,}')
 
    if g('협심증'): by['TextBox 54'].text_frame.paragraphs[0].runs[0].text=f'허혈성\n{g("협심증"):,}'
    if g('급성심근경색'): by['TextBox 55'].text_frame.paragraphs[0].runs[0].text=f'급성심근\n{g("급성심근경색"):,}'
    if g('산정특례심장'): r_set('TextBox 56',0,3,f': {g("산정특례심장"):,}')
 
    if g('일반암'): r_set('TextBox 14',0,1,f': {g("일반암"):,}')
    if g('유사암(갑.기.경.제)'): r_set('TextBox 14',1,2,f': {g("유사암(갑.기.경.제)"):,}')
    if g('항암방사선약물'): r_set('TextBox 14',4,1,f': {g("항암방사선약물"):,} / ')
    if g('표적항암치료비'): r_set('TextBox 14',5,1,f': {g("표적항암치료비"):,} / ')
    if g('세기조절치료'): r_set('TextBox 14',5,4,f': {g("세기조절치료"):,}')
    if g('양성자치료'): r_set('TextBox 14',5,5,f': {g("양성자치료"):,}')
    if g('다빈치로봇수술비'): r_set('TextBox 14',7,1,f': {g("다빈치로봇수술비"):,}')
 
    if g('질병수술비'): r_set('TextBox 17',0,1,f': {g("질병수술비"):,}')
    if any(surg_q): r_set('TextBox 17',3,0,f'({"/".join(str(x) for x in surg_q)})'); r_set('TextBox 17',3,2,'')
    if g('뇌혈관수술비'): r_set('TextBox 17',5,1,f': {g("뇌혈관수술비"):,}')
    if g('심장수술비'): r_set('TextBox 17',7,1,f': {g("심장수술비"):,}')
    if g('상해수술비'): r_set('TextBox 19',0,1,f': {g("상해수술비"):,}')
    if any(surg_s): r_set('TextBox 19',3,0,f'({"/".join(str(x) for x in surg_s)})'); r_set('TextBox 19',3,2,'')
    if g('골절수술비'): r_set('TextBox 19',4,1,f': {g("골절수술비"):,}')
 
    실손_dates=[ct['contract_date'] for ct in contracts
        if any('실손' in k or '입원의료비' in k for k in ct['dambo']) and ct['contract_date']]
    실손가입일=min(실손_dates) if 실손_dates else '___________'
    by['TextBox 59'].text_frame.word_wrap=False
    by['TextBox 59'].text_frame.paragraphs[0].runs[0].text='실손'
    by['TextBox 59'].text_frame.paragraphs[1].runs[0].text='('
    by['TextBox 59'].text_frame.paragraphs[1].runs[1].text='가입일:'
    by['TextBox 59'].text_frame.paragraphs[1].runs[2].text=f'{실손가입일})'
    for r in by['TextBox 59'].text_frame.paragraphs[1].runs: r.font.size=Pt(8)
    if g('입원'): r_set('TextBox 6',0,1,f': {g("입원"):,}')
    if g('통원'): r_set('TextBox 6',1,1,f': {g("통원"):,} / ')
    if g('MRI'): r_set('TextBox 6',2,0,f'MRI : {g("MRI"):,}')
    if g('도수치료'): r_set('TextBox 6',3,1,f': {g("도수치료"):,}')
    if g('비급여주사'): r_set('TextBox 6',4,1,f': {g("비급여주사"):,}')
 
    if g('골절(치아파절제외)'): r_set('TextBox 7',0,1,f': {g("골절(치아파절제외)"):,}')
    if g(' 진 단 비'): r_set('TextBox 7',2,1,f': {g(" 진 단 비"):,}')
    if g('깁스진단비'): r_set('TextBox 7',5,1,f': {g("깁스진단비"):,}')
    if g('응급실(응급)'): r_set('TextBox 7',6,1,f': {g("응급실(응급)"):,}')
    if g('일상배상책임'): r_set('TextBox 5',0,1,f': {g("일상배상책임"):,}')
    if g('대인'): r_set('TextBox 9',0,1,f': {g("대인"):,}')
    if g('대물'): r_set('TextBox 9',1,1,f': {g("대물"):,}')
    if g('변호사'): r_set('TextBox 9',4,1,f': {g("변호사"):,}')
    if g('질병일당'): r_set('TextBox 22',0,1,f': {g("질병일당"):,} / ')
    if g('상해일당'): r_set('TextBox 22',1,1,f': {g("상해일당"):,} / ')
    if g('1인실 상급병원'): r_set('TextBox 22',3,2,f': {g("1인실 상급병원"):,}')
    if g('1인실 종합병원'): r_set('TextBox 22',4,2,f': {g("1인실 종합병원"):,}')
    if g('간병인'): r_set('TextBox 22',7,1,f': {g("간병인"):,} / ')
    if g('간호통합병동'): r_set('TextBox 22',8,2,f': {g("간호통합병동"):,}')
    if g('크라운'): r_set('TextBox 13',0,1,f': {g("크라운"):,}')
    if g('임플란트'): r_set('TextBox 13',1,1,f': {g("임플란트"):,}')
 
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
<title>MAKEONE 보장분석실</title>
<style>
:root{--bg:#0c0d10;--panel:#15171c;--line:#2a2d34;--acc:#E0463B;--acc2:#F4897F;--ink:#EAECEF;--mute:#929aa6;--green:#4ADE80;--blue:#5B9BFF}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:'Pretendard','Noto Sans KR',sans-serif;line-height:1.55}
#gate{position:fixed;inset:0;z-index:100;background:var(--bg);display:flex;flex-direction:column;align-items:center;justify-content:center;padding:30px 26px;text-align:center}
#gate .kick{font-size:14px;font-weight:800;letter-spacing:.45em;color:var(--acc);margin-bottom:14px}
#gate h1{font-size:30px;font-weight:800;margin-bottom:14px}
#gate .s{font-size:14px;color:var(--mute);margin-bottom:38px}
#gate .pw{width:100%;max-width:420px;background:#1a1c22;border:1px solid var(--line);border-radius:14px;padding:18px 20px;font-size:17px;color:var(--ink);text-align:center;letter-spacing:.3em;outline:none}
#gate .pw:focus{border-color:var(--acc)}
#gate .go{width:100%;max-width:420px;margin-top:14px;border:none;border-radius:14px;padding:18px;font-size:17px;font-weight:800;color:#fff;background:var(--acc);cursor:pointer}
#gate .err{color:var(--acc2);font-size:13px;font-weight:700;margin-top:14px;min-height:18px}
.shake{animation:sh .35s}@keyframes sh{0%,100%{transform:translateX(0)}25%{transform:translateX(-8px)}75%{transform:translateX(8px)}}
.app{max-width:520px;margin:0 auto;height:100vh;display:none;flex-direction:column}
header{padding:14px 18px;border-bottom:1px solid var(--line);background:linear-gradient(135deg,#1a1115,#0d0e11 60%,#1c1216);display:flex;align-items:center;gap:10px}
.logo{width:32px;height:32px;border-radius:9px;border:1px solid var(--acc);display:flex;align-items:center;justify-content:center;font-size:16px}
h1{font-size:14px;font-weight:800}h1 b{color:var(--acc2)}.sub{font-size:10px;color:var(--mute)}
.chat{flex:1;overflow-y:auto;padding:16px 12px;display:flex;flex-direction:column;gap:12px}
.msg{max-width:90%;font-size:13px}
.me{align-self:flex-end;background:rgba(224,70,59,.14);border:1px solid rgba(224,70,59,.32);border-radius:14px 14px 4px 14px;padding:9px 13px}
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
.up{flex:1;border:1.5px dashed rgba(224,70,59,.5);border-radius:12px;padding:13px;text-align:center;font-size:13px;font-weight:700;cursor:pointer;color:var(--acc2)}
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
  <div class="kick">MAKEONE</div><h1>MAKEONE 보장분석실</h1>
  <div class="s">접속 비밀번호를 입력하세요</div>
  <input id="pw" class="pw" type="password" inputmode="numeric" placeholder="비밀번호" autocomplete="off">
  <button id="go" class="go">접속</button><div id="gerr" class="err"></div>
</div>
<div class="app" id="app">
  <header><div class="logo">📋</div><div><h1>MAKEONE <b>보장분석실</b></h1>
    <div class="sub">TXT 넣으면 → 엑셀+PPT 개별 다운로드 · 최은혜 지점장</div></div></header>
  <div class="chat" id="chat">
    <div class="msg bot">보장분석지 <b>TXT 파일</b>을 올려주세요. 엑셀·PPT 파일을 각각 드려요.<br><br>
      <span style="font-size:11px;color:var(--mute)">💡 Adobe Acrobat → 편집 → 모두선택(Ctrl+A) → 복사(Ctrl+C)<br>→ 메모장 붙여넣기 → .txt 저장 → 여기 업로드</span></div>
  </div>
  <div class="bar">
    <label class="up" id="up">📄 <span id="uplabel">보장분석지 TXT 선택</span></label>
    <button class="send" id="send" disabled>분석</button>
  </div>
  <div class="qlbl" id="qlbl">📋 분석된 보장분석지에 대해 질문하세요</div>
  <div class="qbar" id="qbar">
    <input class="qinput" id="qinput" placeholder="예: 심장 담보 왜 빠졌어요?" autocomplete="off">
    <button class="qbtn" id="qbtn">질문</button>
  </div>
  <footer>미래를 <b>바르게</b> 설계합니다 · BARUM</footer>
</div>
<input type="file" id="fi" accept=".txt,text/plain" style="display:none">
<script>
const $=s=>document.querySelector(s);let ACCESS='';
async function unlock(){const v=$("#pw").value;$("#gerr").textContent="확인 중…";
  try{const r=await fetch("/check",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({pw:v})});
    const j=await r.json();if(j.ok){ACCESS=v;$("#gerr").textContent="";$("#gate").style.display="none";$("#app").style.display="flex";}else{fail();}}
  catch(e){$("#gerr").textContent="서버 연결 실패";}}
function fail(){$("#gerr").textContent="비밀번호가 올바르지 않습니다.";$("#gate").classList.add("shake");setTimeout(()=>$("#gate").classList.remove("shake"),350);$("#pw").value="";$("#pw").focus();}
$("#go").onclick=unlock;$("#pw").addEventListener("keydown",e=>{if(e.key==="Enter")unlock();});window.addEventListener("load",()=>$("#pw").focus());
const chat=$("#chat");let file=null;
$("#up").onclick=()=>$("#fi").click();
$("#fi").onchange=e=>{file=e.target.files[0];if(file){$("#uplabel").textContent=file.name;$("#send").disabled=false;}};
function esc(s){return String(s==null?"":s).replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));}
function add(html,cls){const d=document.createElement("div");d.className="msg "+cls;d.innerHTML=html;chat.appendChild(d);chat.scrollTop=chat.scrollHeight;return d;}
function b64toBlob(b64,mime){const bin=atob(b64);const arr=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);return new Blob([arr],{type:mime});}
function dl(blob,fname){const u=URL.createObjectURL(blob);const a=document.createElement("a");a.href=u;a.download=fname;document.body.appendChild(a);a.click();document.body.removeChild(a);setTimeout(()=>URL.revokeObjectURL(u),3000);}
$("#send").onclick=async()=>{
  if(!file)return;add("📄 "+esc(file.name),"me");
  $("#send").disabled=true;$("#up").style.opacity=.5;
  const loading=add('<div style="display:flex;align-items:center;gap:11px"><span class="spin"></span><div style="flex:1"><div id="ldmsg" style="font-weight:800">📄 TXT 파싱 중…</div><div id="ldtime" style="font-size:11px;color:var(--mute);margin-top:2px">0초 · 기다려 주세요</div></div></div>',"bot");
  const t0=Date.now();const steps=["📄 TXT 파싱 중…","🔎 담보 추출 중…","📊 엑셀 생성 중…","🖼 PPT 채우는 중…","✅ 완성 중…"];let si=0;
  const timer=setInterval(()=>{si=Math.min(si+1,steps.length-1);const s=Math.floor((Date.now()-t0)/1000);const tm=document.getElementById("ldtime");const mm=document.getElementById("ldmsg");if(tm)tm.textContent=s+"초 경과";if(mm)mm.textContent=steps[si];},8000);
  const fd=new FormData();fd.append("file",file);fd.append("pw",ACCESS);
  try{
    const r=await fetch("/analyze",{method:"POST",body:fd});clearInterval(timer);loading.remove();
    const j=await r.json();
    if(!j.ok){add('<span class="err">⚠ '+esc(j.error||"실패")+'</span>',"bot");}
    else{
      const xlBlob=b64toBlob(j.xlsx_b64,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      dl(xlBlob,j.xlsx_name);
      let ptCard='';
      if(j.pptx_b64){
        const ptBlob=b64toBlob(j.pptx_b64,"application/vnd.openxmlformats-officedocument.presentationml.presentation");
        setTimeout(()=>dl(ptBlob,j.pptx_name),800);
        ptCard=`<div class="file-card pt"><span class="ic">📊</span><span class="nm">${esc(j.pptx_name)}<br><span style="font-size:10px;color:var(--mute)">보장분석 PPT</span></span><span class="dl">저장완료 ✓</span></div>`;}
      add('<b>✅ 분석 완료!</b><div class="summary-box">'+j.summary+'</div><div class="file-cards">'+
        `<div class="file-card xl"><span class="ic">📗</span><span class="nm">${esc(j.xlsx_name)}<br><span style="font-size:10px;color:var(--mute)">보장진단 엑셀</span></span><span class="dl">저장완료 ✓</span></div>`+ptCard+'</div>',"bot");}
  }catch(e){clearInterval(timer);loading.remove();add('<span class="err">오류: '+esc(e.message)+'</span>',"bot");}
  if(j&&j.data){analysisData=j.data;document.getElementById("qbar").style.display="flex";document.getElementById("qlbl").style.display="block";}
  file=null;$("#uplabel").textContent="보장분석지 TXT 선택";$("#send").disabled=true;$("#fi").value="";$("#up").style.opacity=1;
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
</script></body></html>'''
 
@app.get('/health')
def health(): return {'ok':True,'version':'v8-ai-chat'}
 
@app.get('/',response_class=HTMLResponse)
def home(): return INDEX_HTML
 
@app.post('/check')
async def check_pw(body:dict): return {'ok':body.get('pw')==PW}
 
@app.post('/analyze')
async def analyze(file:UploadFile=File(...),pw:str=Form('')):
    if pw!=PW: return JSONResponse({'ok':False,'error':'비밀번호 오류'})
    if not (file.filename.lower().endswith('.txt') or 'text' in (file.content_type or '')):
        return JSONResponse({'ok':False,'error':'TXT 파일만 가능'})
    raw=await file.read()
    for enc in ['utf-8','cp949','euc-kr']:
        try: txt=raw.decode(enc); break
        except: pass
    else: txt=raw.decode('utf-8',errors='ignore')
    try:
        data=parse_txt(txt, file.filename)
        if not data['contracts']:
            return JSONResponse({'ok':False,'error':'계약을 찾지 못했습니다.'})
        cust=data['client']; d=tempfile.mkdtemp(); now=datetime.datetime.now()
        xl=os.path.join(d,f'보장진단_{cust}.xlsx'); pt=os.path.join(d,f'보장분석지_{cust}.pptx')
        build_excel(data,xl); recalc_xlsx(xl); ppt_ok=build_ppt(data,pt)
        xlsx_b64=base64.b64encode(open(xl,'rb').read()).decode()
        response={'ok':True,'xlsx_b64':xlsx_b64,'xlsx_name':f'보장진단_{cust}.xlsx',
                  'summary':make_summary(data),'pptx_ready':ppt_ok}
        if ppt_ok and os.path.exists(pt):
            response['pptx_b64']=base64.b64encode(open(pt,'rb').read()).decode()
            response['pptx_name']=f'보장분석지_{cust}.pptx'
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
 
