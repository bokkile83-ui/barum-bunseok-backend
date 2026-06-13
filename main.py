# -*- coding: utf-8 -*-
"""
BARUM 보장분석 백엔드 v5 — 2026.06.13 최종
TXT(Adobe추출) → 엑셀 + PPT → ZIP 반환
지침: 6월12일 최종지침 전부 반영
"""
import os, re, tempfile, zipfile, datetime, shutil, copy
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Pt

app = FastAPI(title="BARUM 보장분석 v5")
PW   = os.environ.get("ACCESS_PW", os.environ.get("BARUM_PW", "1009"))
HERE = os.path.dirname(os.path.abspath(__file__))
TPL_XL  = os.path.join(HERE, "MASTER_보장분석_엑셀_영구표본.xlsx")
TPL_PPT = os.path.join(HERE, "MASTER_보장분석지_PPT_빈폼.pptx")

# ── 프론트엔드 ────────────────────────────────────────────────────────
PAGE = """<!doctype html>
<html lang=ko><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1,maximum-scale=1">
<title>MAKEONE 보장분석실</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0c0d10;color:#eaecef;font-family:'Apple SD Gothic Neo',system-ui;
  display:flex;min-height:100vh;align-items:center;justify-content:center}
.c{background:linear-gradient(135deg,#0b1f3a,#12294b);border:1px solid #1e3358;
  border-top:3px solid #c9a14a;border-radius:16px;padding:30px 24px;
  max-width:440px;width:92%;text-align:center}
h1{font-size:20px;color:#e6c879;margin-bottom:4px}
.sub{font-size:11px;color:#7a90a8;margin-bottom:20px;line-height:1.6}
input[type=password]{width:100%;padding:12px;border-radius:10px;
  border:1px solid #2a4570;background:#0b1f3a;color:#eaf0f8;font-size:14px;margin-top:10px}
.file-area{width:100%;margin-top:12px;padding:16px;border-radius:10px;
  border:2px dashed #2a4570;background:#061526;color:#7a90a8;
  font-size:13px;cursor:pointer;transition:.2s}
.file-area:hover{border-color:#c9a14a;color:#e6c879}
.file-area.active{border-color:#4ade80;color:#4ade80;background:#061d10}
.btn{width:100%;margin-top:14px;padding:14px;border-radius:10px;border:none;
  background:#c9a14a;color:#0b1f3a;font-size:16px;font-weight:800;cursor:pointer}
.btn:disabled{opacity:.4;cursor:not-allowed}
.msg{margin-top:12px;font-size:13px;min-height:18px}
.ok{color:#4ade80}.err{color:#ff9aa6}.wait{color:#f5b547}
.prog{height:4px;background:#1e3358;border-radius:2px;margin-top:10px;overflow:hidden}
.bar{height:100%;background:linear-gradient(90deg,#c9a14a,#e6c879);border-radius:2px;
  transition:width .5s;width:0%}
.tip{margin-top:16px;font-size:10px;color:#3a5070;line-height:1.8}
</style>
<div class=c>
  <h1>MAKEONE 보장분석실</h1>
  <div class=sub>
    Adobe Acrobat으로 보장분석지 PDF 텍스트 추출 후<br>
    .txt 파일을 업로드하면 엑셀+PPT 세트를 받습니다
  </div>
  <input type=password id=pw placeholder="비밀번호" autocomplete=off>
  <div class=file-area id=fa onclick=document.getElementById('f').click()>
    📄 TXT 파일 선택 (클릭 또는 드래그)
    <input type=file id=f accept=".txt,text/plain" style=display:none onchange=onFile()>
  </div>
  <button class=btn id=btn onclick=go() disabled>분석 → 엑셀+PPT ZIP 다운로드</button>
  <div class=prog><div class=bar id=bar></div></div>
  <div class=msg id=msg></div>
  <div class=tip>
    Adobe Acrobat → 편집 → 모두선택(Ctrl+A) → 복사(Ctrl+C)<br>
    → 메모장 붙여넣기 → .txt 저장 → 여기 업로드
  </div>
</div>
<script>
let fname='';
function onFile(){
  const f=document.getElementById('f').files[0];
  if(!f)return;
  fname=f.name;
  document.getElementById('fa').textContent='✅ '+fname;
  document.getElementById('fa').className='file-area active';
  document.getElementById('btn').disabled=false;
}
document.getElementById('fa').addEventListener('dragover',e=>{e.preventDefault();});
document.getElementById('fa').addEventListener('drop',e=>{
  e.preventDefault();
  const f=e.dataTransfer.files[0];
  if(f){document.getElementById('f').files;
    const dt=new DataTransfer();dt.items.add(f);
    document.getElementById('f').files=dt.files;onFile();}
});
function setMsg(t,c){const m=document.getElementById('msg');m.textContent=t;m.className='msg '+c;}
function setProg(p){document.getElementById('bar').style.width=p+'%';}
async function go(){
  const f=document.getElementById('f').files[0];
  const pw=document.getElementById('pw').value;
  if(!f){setMsg('TXT 파일을 선택하세요','err');return;}
  if(!pw){setMsg('비밀번호를 입력하세요','err');return;}
  document.getElementById('btn').disabled=true;
  setProg(5);setMsg('📄 TXT 파싱 중…','wait');
  const steps=['🔎 별첨 담보 추출 중…','📊 엑셀 생성 중…','🖼 PPT 채우는 중…','✅ ZIP 완성 중…'];
  let si=0;
  const timer=setInterval(()=>{si=Math.min(si+1,steps.length-1);
    setMsg(steps[si],'wait');setProg(15+si*20);},7000);
  try{
    const fd=new FormData();fd.append('file',f);fd.append('pw',pw);
    const r=await fetch('/analyze',{method:'POST',body:fd});
    clearInterval(timer);setProg(100);
    if(!r.ok){const t=await r.text();setMsg('❌ '+t,'err');
      document.getElementById('btn').disabled=false;return;}
    const b=await r.blob(),u=URL.createObjectURL(b),a=document.createElement('a');
    const cn=r.headers.get('content-disposition')||'';
    const fn=cn.match(/filename="?([^"]+)"?/)?.[1]||'보장분석.zip';
    a.href=u;a.download=fn;a.click();
    setMsg('✅ 완료! ZIP 다운로드됨','ok');
  }catch(e){clearInterval(timer);setMsg('❌ 오류: '+e.message,'err');}
  document.getElementById('btn').disabled=false;
}
</script>"""

# ── 스타일 ────────────────────────────────────────────────────────────
W   = Font(color='FFFFFF', name='맑은 고딕', size=9, bold=True)
BL  = Font(color='0070C0', name='맑은 고딕', size=9)   # 갱신
BK  = Font(color='000000', name='맑은 고딕', size=9)   # 비갱신
RD  = Font(color='FF0000', name='맑은 고딕', size=9)   # 미가입/확인
FILL_RED   = PatternFill('solid', fgColor='C00000')
FILL_BLUE  = PatternFill('solid', fgColor='0070C0')
FILL_GREEN = PatternFill('solid', fgColor='375623')
FILL_SUM   = PatternFill('solid', fgColor='2E75B6')
FILL_YEL   = PatternFill('solid', fgColor='FFFF00')
FILL_연두  = PatternFill('solid', fgColor='E2EFDA')
AL = Alignment(horizontal='center', vertical='center', wrap_text=True)

# ── 제외 판정 ──────────────────────────────────────────────────────────
EXCLUDE = ['실효','미납해지','농업인NH안전보험','자동차보험']

def is_excluded(company, product, status=''):
    text = company + product + status
    return any(kw in text for kw in EXCLUDE)

# ── 갱신 판정 (지침 6월12일 최종) ────────────────────────────────────
def judge_renewal(product, expiry, pay_count, company=''):
    # ① 담보명에 갱신형 명시
    if '갱신형' in product and '비갱신' not in product:
        return '갱신'
    if '갱신' in product and '비갱신' not in product:
        return '갱신'
    # ② 만기 9999 → 종신
    if expiry.startswith('9999'):
        return '비갱신(종신)'
    # ③ 납입=보장년수
    try:
        parts = pay_count.split('/')
        if len(parts) == 2:
            total = int(parts[1].strip())
            if total > 240:
                return '비갱신'
            # 납입/보장 비교는 날짜 기반으로
    except: pass
    return '비갱신'

# ── TXT 파싱 ──────────────────────────────────────────────────────────
def parse_txt(txt):
    lines = txt.replace('\r\n','\n').replace('\r','\n').split('\n')
    lines = [l.rstrip() for l in lines]

    # 고객명 추출
    client = '고객'
    for line in lines[:30]:
        l = line.strip()
        m = re.match(r'^([가-힣]{2,5})\s*$', l)
        if m and len(m.group(1)) <= 4:
            client = m.group(1)
            break
        m2 = re.search(r'([가-힣]{2,4})\s+고객님', l)
        if m2:
            client = m2.group(1)
            break

    contracts = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i].strip()

        # 실효계약 리스트 도달 → 전부 스킵
        if '실효계약 리스트' in line or '미납해지' in line:
            break

        if '정상계약 리스트' not in line:
            i += 1
            continue

        i += 1
        # 빈줄 스킵
        while i < n and not lines[i].strip(): i += 1
        if i >= n: break

        company = lines[i].strip()
        i += 1

        # 제외 판정
        if is_excluded(company, ''):
            # 다음 정상계약/실효계약까지 스킵
            while i < n and '정상계약 리스트' not in lines[i] and '실효계약 리스트' not in lines[i]:
                i += 1
            continue

        # 계약자 정보 행 (보험료, 보장기간, 납입)
        contract_date, expiry_date, premium, pay_period, pay_count = '', '', 0, '', ''
        while i < n and not lines[i].strip(): i += 1

        # 계약정보 파싱 (최대 3줄)
        for j in range(i, min(i+5, n)):
            l = lines[j]
            m = re.search(r'(\d{4}\.\d{2}\.\d{2})\s*~\s*(\d{4}\.\d{2}\.\d{2})', l)
            if m:
                contract_date = m.group(1)
                expiry_date   = m.group(2)
            m2 = re.search(r'(\d+)\s*/\s*(\d+)\s*회', l)
            if m2:
                pay_count = f"{m2.group(1)}/{m2.group(2)}"
            m3 = re.search(r'([\d,]+)원', l)
            if m3:
                v = int(m3.group(1).replace(',',''))
                if 1000 < v < 5000000:
                    premium = v
            m4 = re.search(r'(\d+)년납', l)
            if m4:
                pay_period = f"{m4.group(1)}년납"
            m5 = re.search(r'월납\s*/\s*(\d+)년', l)
            if m5 and not pay_period:
                pay_period = f"{m5.group(1)}년납"

        # 상품명 찾기 (계약자 행 다음, 빈줄 건너뜀)
        while i < n and not lines[i].strip(): i += 1
        product = ''
        for j in range(i, min(i+6, n)):
            l = lines[j].strip()
            # 상품명: 한글+숫자 혼합, 담보명이 아닌 것
            if l and not re.search(r'계약자|납입주기|보험료|보장기간', l):
                if len(l) > 5 and not re.search(r'^\d+$', l):
                    if not re.search(r'^\d{4}\.\d{2}', l):
                        product = l
                        i = j + 1
                        break

        renewal = judge_renewal(product, expiry_date, pay_count, company)

        # 담보 파싱
        dambo = {}
        j = i
        while j < n:
            l = lines[j].strip()
            if '정상계약 리스트' in l or '실효계약 리스트' in l:
                i = j
                break
            # 담보명 + 금액 패턴
            m = re.search(r'^(.+?)\s{2,}([\d,]+)\s*$', l)
            if not m:
                m = re.search(r'^(.+?)\s+([\d,]+)\s*$', l)
            if m:
                name = m.group(1).strip()
                try:
                    amount = int(m.group(2).replace(',',''))
                    if 0 < amount <= 100000 and len(name) > 2:
                        # 담보명 정제
                        name = re.sub(r'\s+', ' ', name)
                        dambo[name] = dambo.get(name, 0) + amount
                except: pass
            j += 1
        else:
            i = j

        if company:
            contracts.append({
                'company': company, 'product': product,
                'contract_date': contract_date, 'expiry_date': expiry_date,
                'premium': premium, 'pay_period': pay_period,
                'pay_count': pay_count, 'renewal': renewal,
                'dambo': dambo
            })

    return {'client': client, 'contracts': contracts}

# ── 담보명 표준화 사전 ─────────────────────────────────────────────────
DMAP = {
    # 사망
    '상해사망(갱신형) [보통약관]': '상해사망',
    '상해사망': '상해사망',
    '일반상해사망': '상해사망',
    '기본계약(상해사망(간편가입Ⅲ))담보': '상해사망',
    '일반상해사망후유장해': '상해사망',
    '질병사망': '질병사망(80세)',
    # 후유장해
    '상해후유장해3%': '상해후유3%',
    '상해후유80%': '상해후유80%',
    '질병후유장해3%': '질병후유3%',
    '질병후유80%': '질병후유80%',
    # 암
    '일반암진단비': '일반암',
    '암진단비': '일반암',
    '암진단Ⅱ(유사암제외)(간편가입Ⅲ)담보': '일반암',
    '고액암진단비': '고액암',
    '갑상선암.기타피부암.유사암진단비Ⅲ': '유사암(갑.기.경.제)',
    '유사암진단비': '유사암(갑.기.경.제)',
    '유사암진단Ⅱ(양성뇌종양포함)(간편가입Ⅲ)담보': '유사암(갑.기.경.제)',
    '표적항암약물허가치료비': '표적항암치료비',
    '표적항암약물허가치료(간편가입Ⅲ)(갱신형)담보': '표적항암치료비',
    '항암방사선.약물치료비': '항암방사선약물',
    '항암방사선치료(간편가입Ⅲ)담보': '항암방사선약물',
    '항암약물치료(간편가입Ⅲ)담보': '항암방사선약물',
    '항암방사선(세기조절)치료(간편가입Ⅲ)(갱신형)담보': '세기조절치료',
    '항암세기조절방사선치료비': '세기조절치료',
    '항암방사선(양성자)치료(간편가입Ⅲ)(갱신형)담보': '양성자치료',
    '항암양성자방사선치료비': '양성자치료',
    '암수술(간편가입Ⅲ)담보': '암수술',
    '카티(CAR-T)항암약물허가치료비': '카티치료비',
    '카티(CAR-T)항암약물허가치료(연간1회한)(간편가입Ⅲ)(갱신형)담보': '카티치료비',
    # 뇌혈관
    '뇌혈관질환진단비Ⅲ(건강맞춤형Ⅱ)(갱신형)': '뇌혈관진단비',
    '뇌혈관질환진단비': '뇌혈관진단비',
    '뇌혈관질환진단(간편가입Ⅲ)담보': '뇌혈관진단비',
    '뇌졸중진단비': '뇌졸증진단비',
    '뇌졸중진단(간편가입Ⅲ)담보': '뇌졸증진단비',
    '뇌졸중진단비(건강맞춤형Ⅱ)(갱신형)': '뇌졸증진단비',
    '뇌출혈진단': '뇌출혈진단비',
    '중증질환자(뇌혈관질환)산정특례대상진단비(연간1회한)(건강맞춤형Ⅱ)(갱신형)': '산정특례뇌혈관',
    '뇌혈관질환수술비Ⅲ(건강맞춤형Ⅱ)(갱신형)': '뇌혈관수술비',
    '뇌혈관질환수술비': '뇌혈관수술비',
    '심뇌혈관질환수술(간편가입Ⅲ)담보': '뇌혈관수술비',
    '뇌경색증(I63)혈전용해치료비': '혈전용해치료비',
    '혈전용해치료비Ⅱ(뇌졸중)(간편가입Ⅲ)담보': '혈전용해치료비',
    # 심장
    '허혈심장질환진단비Ⅲ(건강맞춤형Ⅱ)(갱신형)': '협심증',
    '허혈심장질환진단비': '협심증',
    '허혈심장질환진단(간편가입Ⅲ)담보': '협심증',
    '급성심근경색증진단': '급성심근경색',
    '급성심근경색증진단(간편가입Ⅲ)담보': '급성심근경색',
    '중증질환자(심장질환)산정특례대상진단비(연간1회한)(건강맞춤형Ⅱ)(갱신형)': '산정특례심장',
    '허혈심장질환수술비Ⅲ(건강맞춤형Ⅱ)(갱신형)': '허혈성수술비',
    '허혈심장질환수술비': '허혈성수술비',
    '급성심근경색증(I21)혈전용해치료비': '혈전용해치료비',
    '혈전용해치료비Ⅱ(특정심장질환)(간편가입Ⅲ)담보': '혈전용해치료비',
    # 수술비 (기본)
    '질병수술비': '질병수술비',
    '질병수술비(건강맞춤형Ⅱ)(갱신형)': '질병수술비',
    '질병수술비(백내장및대장용종제외)(건강맞춤형Ⅱ)(갱신형)': '질병수술비',
    '질병수술(간편가입Ⅲ)담보': '질병수술비',
    '상해수술비(건강맞춤형Ⅱ)(갱신형)': '상해수술비',
    '상해수술(간편가입Ⅲ)담보': '상해수술비',
    '골절수술(간편가입Ⅲ)담보': '골절수술비',
    '골절수술비': '골절수술비',
    '화상수술(간편가입Ⅲ)담보': '화상수술비',
    '120대질병수술Ⅱ(간편가입Ⅲ)(질병수술3(24대질병))담보': '120대수술비',
    '5대기관질병수술(관혈/비관혈)(연간1회한)(간편가입Ⅲ)담보': '5대기관 수술비 관혈',
    '중대한특정상해수술(간편가입Ⅲ)담보': '중대한상해수술비',
    # 입원/일당
    '질병입원의료비': '실손입원', '상해입원의료비': '실손입원',
    '질병외래의료비': '실손통원',
    '도수/체외충격파/증식치료': '도수치료',
    '비급여주사제': '비급여주사',
    'MRI검사의료비': 'MRI',
    '간병인사용질병입원일당(1일이상)(요양병원)(간편가입)(갱신형)': '간병인',
    '간호간병통합서비스질병입원일당(1-180일)(간편가입)(갱신형)': '간호통합병동',
    '간호간병통합서비스 질병입원일당(1-60일)(간편가입)(갱신형)': '간호통합병동',
    '상급종합병원질병입원일당(상급병실(1인실),1일이상60일한도)(간편가입)(갱신형)': '1인실 상급병원',
    '종합병원질병입원일당(상급병실(1인실),1일이상30일한도)(간편가입)(갱신형)': '1인실 종합병원',
    # 운전자
    '교통사고처리지원금': '교통사고처리지원금',
    '교통사고 처리지원금': '교통사고처리지원금',
    '교통사고벌금(대인)': '교통사고벌금(대인)',
    '교통사고벌금(대물)': '교통사고벌금(대물)',
    '자동차사고 변호사선임비용': '변호사선임비용',
    '변호사선임비용': '변호사선임비용',
    '무보험차에 의한 상해': '무보험차',
    # 기타
    '골절진단(간편가입Ⅲ)담보': '골절(치아파절제외)',
    '골절진단비': '골절(치아파절제외)',
    '깁스치료담보': '깁스진단비',
    '깁스치료': '깁스진단비',
    '응급실내원비(응급)': '응급실내원비(응급)',
    '가족생활배상책임': '일상배상책임',
    '일상생활배상책임': '일상배상책임',
    '보험료납입지원': None,  # 제외
    # 치아
    '치과치료(보존치료)': '보존치료',
    '치과치료(보철치료)': '보철치료',
}

def resolve(raw):
    """담보명 표준화"""
    if raw in DMAP:
        return DMAP[raw]
    for k,v in DMAP.items():
        if k in raw:
            return v
    # 1~5종 수술비 직접 판단
    if '1~5종' in raw or '(1종)' in raw or '(2종)' in raw or \
       '(3종)' in raw or '(4종)' in raw or '(5종)' in raw:
        return None  # 슬래시로 처리
    return raw

def get_종번호(name):
    for i,k in enumerate(['(1종)','(2종)','(3종)','(4종)','(5종)'],1):
        if k in name: return i
    return 0

# ── 엑셀 생성 ─────────────────────────────────────────────────────────
def build_excel(data, out):
    wb = openpyxl.load_workbook(TPL_XL)
    ws = wb['보장분석']

    client = data['client']
    contracts = data['contracts']

    ws.cell(1,1).value = client + ' 보장진단'
    ws.cell(1,1).font = Font(name='맑은 고딕', size=11, bold=True)
    ws.row_dimensions[1].height = 30

    # B열 담보명→행 매핑
    nm2r = {}
    for r in range(6, ws.max_row+1):
        v = ws.cell(r,2).value
        if v: nm2r[str(v).strip()] = r

    def sv(nm, col, val, gen=False, check=False):
        r = nm2r.get(nm)
        if not r: return
        try:
            c = ws.cell(r,col)
            if check:
                c.value = '[확인]'; c.font = RD; c.fill = FILL_YEL
            else:
                c.value = val
                c.font = BL if gen else BK
        except: pass

    total_premium = 0

    for i, ct in enumerate(contracts):
        col = 3 + i
        gen   = ct['renewal'] == '갱신'
        paid  = ct['renewal'] == '완납'
        종신  = ct['renewal'] == '비갱신(종신)'

        # 헤더 1행
        h = ws.cell(1, col)
        h.value = f"{ct['company']}\n[{ct['renewal']}]"
        h.font = W; h.alignment = AL
        h.fill = FILL_GREEN if paid else (FILL_BLUE if gen else FILL_RED)

        # 행2: 보험료 (숫자만, 원/한글 금지)
        pm = ct['premium']
        ws.cell(2,col).value = pm if pm else ''
        ws.cell(2,col).font = BL if gen else BK
        total_premium += pm if pm else 0

        # 행3~5
        ws.cell(3,col).value = ct['contract_date']; ws.cell(3,col).font = BL if gen else BK
        ws.cell(4,col).value = ct['expiry_date'];   ws.cell(4,col).font = BL if gen else BK
        ws.cell(5,col).value = f"{ct['pay_period']} ({ct['pay_count']})" if ct['pay_period'] else ''
        ws.cell(5,col).font = BL if gen else BK

        dambo = ct['dambo']

        # 1~5종 슬래시 처리
        q_종 = {k:v for k,v in dambo.items() if '질병수술비' in k and get_종번호(k)>0}
        s_종 = {k:v for k,v in dambo.items() if '상해수술비' in k and get_종번호(k)>0}

        if q_종:
            vals = [0]*5
            for k,v in q_종.items():
                idx = get_종번호(k)-1
                if 0<=idx<5: vals[idx] = v
            r = nm2r.get('질병 종수술비(1-5종)')
            if r:
                ws.cell(r,col).value = '/'.join(str(x) for x in vals)
                ws.cell(r,col).font = BL if gen else BK

        if s_종:
            vals = [0]*5
            for k,v in s_종.items():
                idx = get_종번호(k)-1
                if 0<=idx<5: vals[idx] = v
            r = nm2r.get('상해 종수술비(1-5종)')
            if r:
                ws.cell(r,col).value = '/'.join(str(x) for x in vals)
                ws.cell(r,col).font = BL if gen else BK

        # 개별 담보 입력
        skip_set = set(list(q_종.keys()) + list(s_종.keys()))

        for raw, amount in dambo.items():
            if raw in skip_set: continue
            std = resolve(raw)
            if std is None: continue
            if std in nm2r:
                r = nm2r[std]
                existing = ws.cell(r,col).value
                if isinstance(existing,(int,float)):
                    ws.cell(r,col).value = existing + amount
                elif existing and '/' in str(existing):
                    pass
                else:
                    ws.cell(r,col).value = amount
                ws.cell(r,col).font = BL if gen else BK

    # 합계열 — 마지막 계약 다음 열
    last_col = 3 + len(contracts)

    # 기존 col10 합계열 내용 제거
    for r in range(1, ws.max_row+1):
        try:
            if ws.cell(r,10).value and ws.cell(r,10).value != '' and last_col != 10:
                ws.cell(r,10).value = None
        except: pass

    ws.cell(1,last_col).value = '합계'
    ws.cell(1,last_col).font = W
    ws.cell(1,last_col).fill = FILL_SUM
    ws.cell(1,last_col).alignment = AL

    # 행2 보험료 합계 (숫자만)
    ws.cell(2,last_col).value = total_premium if total_premium else ''
    ws.cell(2,last_col).font = W

    # SUM 수식으로 합계 계산
    for r in range(6, ws.max_row+1):
        slash_t=[0]*5; is_slash=False; total=0
        for col in range(3,last_col):
            v=ws.cell(r,col).value
            if isinstance(v,(int,float)): total+=int(v)
            elif isinstance(v,str) and '/' in v:
                is_slash=True
                for k,p in enumerate(v.split('/')[:5]):
                    try: slash_t[k]+=int(p)
                    except: pass
        sc=ws.cell(r,last_col)
        try:
            if is_slash and any(slash_t):
                sc.value='/'.join(str(x) for x in slash_t); sc.font=BK
            elif total>0:
                sc.value=total; sc.font=BK
            else: sc.value=None
        except: pass

    # 열 너비
    ws.column_dimensions['B'].width=22
    for c in range(3,last_col+1):
        ws.column_dimensions[get_column_letter(c)].width=15

    # 확인사항 시트
    if '📋확인사항' in wb.sheetnames:
        del wb['📋확인사항']
    ws2 = wb.create_sheet('📋확인사항')
    ws2.cell(1,1,f'{client} · 자동분석 {datetime.datetime.now():%Y.%m.%d}').font=Font(bold=True,color='C9A14A')
    ws2.cell(2,1,'분류'); ws2.cell(2,2,'대상'); ws2.cell(2,3,'내용')
    ws2.cell(3,1,'완료'); ws2.cell(3,2,f'계약 {len(contracts)}건')
    ws2.cell(3,3,f'월보험료합계: {total_premium:,}원')

    wb.save(out)

# ── PPT 채우기 ────────────────────────────────────────────────────────
def build_ppt(data, out):
    prs = Presentation(TPL_PPT)
    sl = prs.slides[0]
    by = {sh.name: sh for sh in sl.shapes if sh.has_text_frame}

    client    = data['client']
    contracts = data['contracts']
    now       = datetime.datetime.now()

    # 합계 계산
    totals = {}
    for ct in contracts:
        for raw, amount in ct['dambo'].items():
            std = resolve(raw)
            if std:
                totals[std] = totals.get(std,0) + amount

    def get(nm): return totals.get(nm, 0)

    def r_set(box, pi, ri, val):
        if box not in by: return
        tf = by[box].text_frame
        if pi < len(tf.paragraphs):
            p = tf.paragraphs[pi]
            if ri < len(p.runs):
                p.runs[ri].text = val

    def force(box, old, new):
        if box not in by: return
        tf = by[box].text_frame
        for p in tf.paragraphs:
            full=''.join(r.text for r in p.runs)
            if old not in full: continue
            nf=full.replace(old,new)
            if p.runs:
                p.runs[0].text=nf
                for r in p.runs[1:]: r.text=''
            return

    # 제목/날짜
    by['TextBox 21'].text_frame.word_wrap = False
    by['TextBox 21'].text_frame.paragraphs[0].runs[0].text = f'{client} 님의 보장'
    by['TextBox 21'].text_frame.paragraphs[0].runs[1].text = '(전)'
    by['TextBox 36'].text_frame.paragraphs[0].runs[0].text = f'{now.year}년'
    by['TextBox 35'].text_frame.paragraphs[0].runs[0].text = f'{now.month:02d}월'
    by['TextBox 29'].text_frame.paragraphs[0].runs[0].text = f'{now.day:02d}일 기준'
    for b in ['TextBox 36','TextBox 35','TextBox 29']:
        by[b].text_frame.word_wrap = False

    # 사망
    q_death = get('질병사망(80세)')
    s_death = get('상해사망')
    종신_death = 0
    for ct in contracts:
        if '종신' in ct['renewal']:
            for raw,v in ct['dambo'].items():
                std = resolve(raw)
                if std == '상해사망': 종신_death += v

    if q_death: r_set('TextBox 10',2,2,f'{q_death:,}')
    if 종신_death: r_set('TextBox 10',3,1,f': {종신_death:,}')
    if s_death: r_set('TextBox 11',0,1,f': {s_death:,}')

    # 후유장해
    if get('상해후유3%'): r_set('TextBox 8',2,1,f'3% : {get("상해후유3%"):,}')
    if get('질병후유3%'): r_set('TextBox 8',0,1,f'3% : {get("질병후유3%"):,}')
    if get('상해후유80%'): r_set('TextBox 8',3,1,f'80% : {get("상해후유80%"):,}')
    if get('질병후유80%'): r_set('TextBox 8',1,1,f'80% : {get("질병후유80%"):,}')

    # 뇌혈관
    if get('뇌혈관진단비'):
        by['TextBox 46'].text_frame.paragraphs[0].runs[0].text = f'뇌혈관\n{get("뇌혈관진단비"):,}'
    if get('뇌졸증진단비'):
        by['TextBox 47'].text_frame.paragraphs[0].runs[0].text = f'뇌졸증\n{get("뇌졸증진단비"):,}'
    if get('뇌출혈진단비'):
        by['TextBox 48'].text_frame.paragraphs[0].runs[0].text = f'뇌출혈\n{get("뇌출혈진단비"):,}'
    if get('산정특례뇌혈관'): r_set('TextBox 49',0,3,f': {get("산정특례뇌혈관"):,}')
    if get('혈전용해치료비'):  r_set('TextBox 49',1,1,f': {get("혈전용해치료비"):,}')

    # 심장
    if get('협심증'):
        by['TextBox 54'].text_frame.paragraphs[0].runs[0].text = f'허혈성\n{get("협심증"):,}'
    if get('급성심근경색'):
        by['TextBox 55'].text_frame.paragraphs[0].runs[0].text = f'급성심근\n{get("급성심근경색"):,}'
    if get('산정특례심장'): r_set('TextBox 56',0,3,f': {get("산정특례심장"):,}')

    # 암
    암진 = get('일반암')
    유사암 = get('유사암(갑.기.경.제)')
    항암 = get('항암방사선약물')
    표적 = get('표적항암치료비')
    세기 = get('세기조절치료')
    양성자 = get('양성자치료')
    암수술 = get('암수술')
    다빈치 = get('다빈치로봇수술비')

    if 암진:   r_set('TextBox 14',0,1,f': {암진:,}')
    if 유사암: r_set('TextBox 14',1,2,f': {유사암:,}')
    if 항암:   r_set('TextBox 14',4,1,f': {항암:,} / ')
    if 표적:   r_set('TextBox 14',5,1,f': {표적:,} / ')
    if 세기:   r_set('TextBox 14',5,4,f': {세기:,}')
    if 양성자: r_set('TextBox 14',5,5,f': {양성자:,}')
    if 다빈치: r_set('TextBox 14',7,1,f': {다빈치:,}')

    # 수술비 (질병)
    q수술 = get('질병수술비')
    if q수술: r_set('TextBox 17',0,1,f': {q수술:,}')
    # 1~5종 슬래시 (PPT엔 합산값)
    q슬 = [0]*5; s슬 = [0]*5
    for ct in contracts:
        for raw,v in ct['dambo'].items():
            i = get_종번호(raw)
            if i>0:
                if '질병수술비' in raw: q슬[i-1]+=v
                if '상해수술비' in raw: s슬[i-1]+=v
    if any(q슬):
        r_set('TextBox 17',3,0,f'({"/".join(str(x) for x in q슬)})')
        r_set('TextBox 17',3,2,'')
    뇌혈수술 = get('뇌혈관수술비')
    허혈수술 = get('허혈성수술비')
    심수술   = get('심장수술비')
    if 뇌혈수술: r_set('TextBox 17',5,1,f': {뇌혈수술:,}')
    if 허혈수술: r_set('TextBox 17',6,2,f': {허혈수술:,}')
    if 심수술:   r_set('TextBox 17',7,1,f': {심수술:,}')

    # 수술비 (상해)
    s수술 = get('상해수술비')
    if s수술: r_set('TextBox 19',0,1,f': {s수술:,}')
    if any(s슬):
        r_set('TextBox 19',3,0,f'({"/".join(str(x) for x in s슬)})')
        r_set('TextBox 19',3,2,'')
    if get('골절수술비'): r_set('TextBox 19',4,1,f': {get("골절수술비"):,}')

    # 실손
    실손입원 = get('실손입원')
    실손통원 = get('실손통원')
    도수 = get('도수치료')
    비급여 = get('비급여주사')
    mri  = get('MRI')
    # 가입일: 가장 오래된 계약
    실손_dates = []
    for ct in contracts:
        if any('실손' in k or '입원의료비' in k for k in ct['dambo']):
            if ct['contract_date']: 실손_dates.append(ct['contract_date'])
    실손가입일 = min(실손_dates) if 실손_dates else '___________'

    by['TextBox 59'].text_frame.word_wrap=False
    by['TextBox 59'].text_frame.paragraphs[0].runs[0].text='실손'
    by['TextBox 59'].text_frame.paragraphs[1].runs[0].text='('
    by['TextBox 59'].text_frame.paragraphs[1].runs[1].text='가입일:'
    by['TextBox 59'].text_frame.paragraphs[1].runs[2].text=f'{실손가입일})'
    for r in by['TextBox 59'].text_frame.paragraphs[1].runs:
        r.font.size = Pt(8)

    if 실손입원: r_set('TextBox 6',0,1,f': {실손입원:,}')
    if 실손통원: r_set('TextBox 6',1,1,f': {실손통원:,} / ')
    if mri:      r_set('TextBox 6',2,0,f'MRI : {mri:,}')
    if 도수:     r_set('TextBox 6',3,1,f': {도수:,}')
    if 비급여:   r_set('TextBox 6',4,1,f': {비급여:,}')

    # 상해/기타
    골절 = get('골절(치아파절제외)')
    깁스  = get('깁스진단비')
    응급 = get('응급실내원비(응급)')
    화상 = get(' 진 단 비') or get('화상진단비')
    if 골절: r_set('TextBox 7',0,1,f': {골절:,}')
    if 화상: r_set('TextBox 7',2,1,f': {화상:,}')
    if 깁스:  r_set('TextBox 7',5,1,f': {깁스:,}')
    if 응급: r_set('TextBox 7',6,1,f': {응급:,}')

    # 일배책
    일배책 = get('일상배상책임')
    if 일배책: r_set('TextBox 5',0,1,f': {일배책:,}')

    # 운전자
    대인 = get('교통사고처리지원금')
    대물 = get('교통사고벌금(대물)')
    변호 = get('변호사선임비용')
    if 대인: r_set('TextBox 9',0,1,f': {대인:,}')
    if 대물: r_set('TextBox 9',1,1,f': {대물:,}')
    if 변호: r_set('TextBox 9',4,1,f': {변호:,}')

    # 입원/간병
    질병일당 = get('질병일당')
    상해일당 = get('상해일당')
    간병인   = get('간병인')
    요양병원 = get('요양병원일당') or get('요양병원')
    간호통합 = get('간호통합병동')
    인실상급 = get('1인실 상급병원')
    인실종합 = get('1인실 종합병원')
    if 질병일당: r_set('TextBox 22',0,1,f': {질병일당:,} / ')
    if 상해일당: r_set('TextBox 22',1,1,f': {상해일당:,} / ')
    if 인실상급: r_set('TextBox 22',3,2,f': {인실상급:,}')
    if 인실종합: r_set('TextBox 22',4,2,f': {인실종합:,}')
    if 간병인:   r_set('TextBox 22',7,1,f': {간병인:,} / ')
    if 요양병원: r_set('TextBox 22',7,3,f': {요양병원:,}')
    if 간호통합: r_set('TextBox 22',8,2,f': {간호통합:,}')

    # 치아
    크라운 = get('크라운') or get('보존치료')
    임플란트 = get('임플란트') or get('보철치료')
    if 크라운:   r_set('TextBox 13',0,1,f': {크라운:,}')
    if 임플란트: r_set('TextBox 13',1,1,f': {임플란트:,}')

    prs.save(out)

# ── FastAPI 엔드포인트 ─────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def home(): return PAGE

@app.get("/health")
def health(): return {"ok": True, "version": "v5-txt-set"}

@app.post("/analyze")
async def analyze(file: UploadFile = File(...), pw: str = Form("")):
    if pw != PW:
        raise HTTPException(401, "비밀번호 오류")
    if not (file.filename.lower().endswith(".txt") or 'text' in (file.content_type or '')):
        raise HTTPException(400, "TXT 파일만 가능. Adobe Acrobat 텍스트 추출본을 업로드하세요.")

    raw = await file.read()
    for enc in ['utf-8','cp949','euc-kr']:
        try: txt = raw.decode(enc); break
        except: pass
    else: txt = raw.decode('utf-8', errors='ignore')

    d   = tempfile.mkdtemp()
    now = datetime.datetime.now()
    xl  = os.path.join(d, f"보장진단_{now:%H%M%S}.xlsx")
    pt  = os.path.join(d, f"보장분석지_{now:%H%M%S}.pptx")
    zp  = os.path.join(d, f"보장분석_{now:%H%M%S}.zip")

    try:
        data = parse_txt(txt)
        if not data['contracts']:
            raise HTTPException(400, "계약을 찾지 못했습니다. (정상계약 리스트) 섹션이 포함된 TXT인지 확인하세요.")
        build_excel(data, xl)
        if os.path.exists(TPL_PPT):
            build_ppt(data, pt)
    except HTTPException: raise
    except Exception as e:
        raise HTTPException(500, f"분석오류: {e}")

    client = data['client']
    xl_name = f"보장진단_{client}.xlsx"
    pt_name = f"보장분석지_{client}.pptx"
    zip_name = f"보장분석_{client}_{now:%m%d}.zip"

    with zipfile.ZipFile(zp, "w") as z:
        z.write(xl, xl_name)
        if os.path.exists(pt):
            z.write(pt, pt_name)

    return FileResponse(zp, filename=zip_name,
                        media_type="application/zip",
                        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'})
