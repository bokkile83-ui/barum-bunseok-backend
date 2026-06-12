# -*- coding: utf-8 -*-
"""
BARUM 보장분석 백엔드 v3 — 2026.06.12
PDF → Claude Vision API → 담보 추출 → 엑셀(zip) 반환
이미지형/텍스트형 PDF 모두 처리.
"""
import os, tempfile, datetime, zipfile, base64, json, re
import anthropic
import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = FastAPI(title="BARUM 보장분석 v3")

PW   = os.environ.get("ACCESS_PW", os.environ.get("BARUM_PW", "1009"))
HERE = os.path.dirname(os.path.abspath(__file__))
TPL  = os.path.join(HERE, "template_canon.xlsx")
AKEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ── 프론트엔드 HTML ────────────────────────────────────────────────────
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
  max-width:400px;width:92%;text-align:center}
h1{font-size:20px;color:#e6c879;margin-bottom:6px}
.sub{font-size:12px;color:#7a90a8;margin-bottom:24px}
input[type=password],input[type=file]{width:100%;padding:12px;border-radius:10px;
  border:1px solid #2a4570;background:#0b1f3a;color:#eaf0f8;font-size:14px;margin-top:10px}
.btn{width:100%;margin-top:16px;padding:14px;border-radius:10px;border:none;
  background:#c9a14a;color:#0b1f3a;font-size:16px;font-weight:800;cursor:pointer}
.btn:disabled{opacity:.5;cursor:not-allowed}
.msg{margin-top:14px;font-size:13px;min-height:20px}
.ok{color:#4ade80}.err{color:#ff9aa6}.wait{color:#f5b547}
.prog{height:4px;background:#1e3358;border-radius:2px;margin-top:10px;overflow:hidden}
.prog-fill{height:100%;background:#c9a14a;border-radius:2px;transition:width .4s;width:0%}
</style>
<div class=c>
  <h1>MAKEONE 보장분석실</h1>
  <div class=sub>보장분석지 PDF → 보장진단 엑셀 자동 생성</div>
  <input type=password id=pw placeholder="비밀번호" autocomplete=off>
  <input type=file id=f accept=.pdf>
  <button class=btn id=btn onclick=go()>분석 → 엑셀 다운로드</button>
  <div class=prog><div class=prog-fill id=pf></div></div>
  <div class=msg id=msg></div>
</div>
<script>
let timer;
function setMsg(txt,cls){document.getElementById('msg').textContent=txt;document.getElementById('msg').className='msg '+cls;}
function setProg(pct){document.getElementById('pf').style.width=pct+'%';}
async function go(){
  const f=document.getElementById('f').files[0];
  const pw=document.getElementById('pw').value;
  const btn=document.getElementById('btn');
  if(!f){setMsg('PDF를 선택하세요','err');return;}
  if(!pw){setMsg('비밀번호를 입력하세요','err');return;}
  btn.disabled=true; setProg(5);
  const steps=['📄 PDF 읽는 중…','🔎 Claude AI 분석 중…','📊 엑셀 생성 중…','✅ 마무리 중…'];
  let si=0; setMsg(steps[0],'wait');
  timer=setInterval(()=>{si=Math.min(si+1,steps.length-1);setMsg(steps[si],'wait');setProg(10+si*22);},15000);
  try{
    const fd=new FormData();fd.append('file',f);fd.append('pw',pw);
    const r=await fetch('/analyze',{method:'POST',body:fd});
    clearInterval(timer); setProg(100);
    if(!r.ok){const t=await r.text();setMsg('실패: '+t,'err');btn.disabled=false;return;}
    const b=await r.blob(),u=URL.createObjectURL(b),a=document.createElement('a');
    a.href=u;a.download='보장분석.zip';a.click();
    setMsg('완료 ✅ zip 다운로드됨','ok');
  }catch(e){clearInterval(timer);setMsg('오류: '+e.message,'err');}
  btn.disabled=false;
}
</script>"""

# ── 담보 프롬프트 ─────────────────────────────────────────────────────
SYSTEM = """당신은 보험 보장분석지 전문 분석 AI입니다.
업로드된 보험 보장분석지 PDF 이미지에서 모든 계약의 담보명과 가입금액을 빠짐없이 추출합니다.

규칙:
- NH농협생명·NH농협손보 = 완전 제외
- 금액 단위 = 만원
- 추측 금지, 모르면 null
- 1~5종 수술비 = "10/50/100/500/1000" 슬래시 형식
- 갱신판정: ①갱신형 명시→갱신 ②9999만기→비갱신(종신) ③납입=보장→갱신 ④나머지→비갱신

반드시 아래 JSON 형식만 출력 (다른 텍스트 없이):
{
  "client": "고객명",
  "contracts": [
    {
      "no": 1,
      "company": "회사명",
      "product": "상품명",
      "contract_date": "2020.03.15",
      "expiry_date": "2040.03.15",
      "premium": 50000,
      "payment_period": "20년납",
      "payment_count": "67/240",
      "renewal": "갱신 또는 비갱신 또는 비갱신(종신)",
      "담보": {
        "일반사망": 10000,
        "상해사망": 10000,
        "일반암": 3000,
        "뇌혈관진단비": 3000,
        "급성심근경색": 3000,
        "질병수술비": 50,
        "상해수술비": 100,
        "실손입원": 5000,
        "질병일당": 2,
        "상해일당": 2
      }
    }
  ],
  "memo": [
    {"tag": "확인", "item": "담보명", "note": "판독불가 증권확인필요"}
  ]
}"""

# ── Claude Vision으로 PDF 분석 ────────────────────────────────────────
def analyze_with_claude(pdf_path: str) -> dict:
    client = anthropic.Anthropic(api_key=AKEY)

    # PDF → 페이지별 이미지 변환
    doc = fitz.open(pdf_path)
    images = []
    for pi in range(len(doc)):
        page = doc[pi]
        mat = fitz.Matrix(1.5, 1.5)  # 해상도 150%
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("jpeg")
        b64 = base64.standard_b64encode(img_bytes).decode()
        images.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
        })

    # 텍스트 추출도 병행 (텍스트형 PDF 대응)
    txt_content = ""
    for pi in range(len(doc)):
        txt = doc[pi].get_text()
        if txt.strip():
            txt_content += f"\n[페이지{pi+1}]\n{txt}"

    # 메시지 구성
    content = images[:20]  # 최대 20페이지
    if txt_content:
        content.append({
            "type": "text",
            "text": f"아래는 PDF에서 추출한 텍스트입니다 (이미지와 함께 참고):\n{txt_content[:8000]}"
        })
    content.append({
        "type": "text",
        "text": "이 보장분석지의 모든 계약과 담보값을 추출해서 JSON으로만 응답해주세요."
    })

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        system=SYSTEM,
        messages=[{"role": "user", "content": content}]
    )

    raw = resp.content[0].text.strip()
    # JSON 파싱
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)
    return json.loads(raw)


# ── 엑셀 생성 ─────────────────────────────────────────────────────────
# 표준 담보 행 매핑 (template_canon.xlsx 기준)
DAMBO_ALIAS = {
    '일반암(갱신)':'일반암','일반암(비갱신)':'일반암','비갱신암':'일반암',
    '뇌혈관진단':'뇌출혈진단비','허혈성진단비':'허혈성','허혈성심장질환':'허혈성',
    '질병사망(80세)':'질병사망','재해사망':'상해사망',
    '상해후유장해3%':'상해후유3%','질병후유장해3%':'질병후유3%',
    '실손입원의료비':'실손입원','실손통원의료비':'실손통원',
}

def build_excel(data: dict, template: str, out: str):
    try:
        wb = openpyxl.load_workbook(template)
        ws = wb['보장진단']
    except Exception:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '보장진단'

    client_name = data.get('client', '고객')
    contracts = [c for c in data.get('contracts', [])
                 if '농협' not in c.get('company', '')]

    # 고객명
    ws.cell(1, 1).value = f"{client_name} 보장진단"

    # 헤더 색상
    FILL_RED   = PatternFill('solid', fgColor='C00000')
    FILL_BLUE  = PatternFill('solid', fgColor='0070C0')
    FILL_GREEN = PatternFill('solid', fgColor='375623')
    WHITE = Font(color='FFFFFF', name='맑은 고딕', size=9, bold=True)
    BLUE  = Font(color='0070C0', name='맑은 고딕', size=9)
    BLACK = Font(color='000000', name='맑은 고딕', size=9)
    RED   = Font(color='FF0000', name='맑은 고딕', size=9)

    for i, c in enumerate(contracts):
        col = 3 + i
        renewal = c.get('renewal', '비갱신')
        is_gen  = '갱신' in renewal and '비갱신' not in renewal

        # 행1: 회사명
        h = ws.cell(1, col)
        h.value = f"{c.get('company','?')} [{renewal}]"
        h.font  = WHITE
        h.fill  = FILL_BLUE if is_gen else FILL_RED
        h.alignment = Alignment(horizontal='center', wrap_text=True)

        # 행2~7: 보험료/날짜/납입
        ws.cell(2, col).value = c.get('premium', '')
        ws.cell(3, col).value = c.get('contract_date', '')
        ws.cell(4, col).value = c.get('expiry_date', '')
        ws.cell(5, col).value = c.get('payment_period', '')
        ws.cell(6, col).value = c.get('payment_count', '')

        # 담보값 기입 — B열 담보명과 매칭
        dambo = c.get('담보', {})
        for row in range(8, ws.max_row + 1):
            b_val = ws.cell(row, 2).value
            if not b_val:
                continue
            # 직접 매칭
            val = dambo.get(b_val)
            # 별칭 매칭
            if val is None:
                for alias, std in DAMBO_ALIAS.items():
                    if std == b_val and alias in dambo:
                        val = dambo[alias]
                        break
            # 부분 매칭
            if val is None:
                for k, v in dambo.items():
                    if k in b_val or b_val in k:
                        val = v
                        break
            if val is not None and val != 0:
                cell = ws.cell(row, col)
                cell.value = val
                cell.font = BLUE if is_gen else BLACK

    # 합계열
    last_col = 3 + len(contracts)
    ws.cell(1, last_col).value = '합계'
    ws.cell(1, last_col).font = WHITE
    ws.cell(1, last_col).fill = PatternFill('solid', fgColor='2E75B6')

    for row in range(8, ws.max_row + 1):
        total = 0
        slash_vals = [0,0,0,0,0]
        is_slash = False
        for col in range(3, last_col):
            v = ws.cell(row, col).value
            if isinstance(v, (int, float)):
                total += v
            elif isinstance(v, str) and '/' in v:
                is_slash = True
                for k, p in enumerate(v.split('/')[:5]):
                    try: slash_vals[k] += int(p)
                    except: pass
        sc = ws.cell(row, last_col)
        if is_slash:
            sc.value = '/'.join(str(x) for x in slash_vals)
        elif total > 0:
            sc.value = total
            sc.font = BLACK

    # 확인사항 시트
    memo = data.get('memo', [])
    if memo:
        sh2 = wb.create_sheet('📋확인사항')
        sh2.cell(1,1,f'{client_name} · 확인사항').font = Font(bold=True, color='C9A14A')
        sh2.cell(2,1,'분류'); sh2.cell(2,2,'대상'); sh2.cell(2,3,'내용')
        for j, m in enumerate(memo[:80]):
            sh2.cell(3+j,1,m.get('tag',''))
            sh2.cell(3+j,2,m.get('item',''))
            sh2.cell(3+j,3,m.get('note',''))

    wb.save(out)
    return {'고객': client_name, '계약': len(contracts)}


# ── FastAPI 엔드포인트 ─────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def home():
    return PAGE

@app.get("/health")
def health():
    return {"ok": True, "version": "v3-claude-vision"}

@app.post("/analyze")
async def analyze(file: UploadFile = File(...), pw: str = Form("")):
    if pw != PW:
        raise HTTPException(401, "비밀번호 오류")
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "PDF만 가능")
    if not AKEY:
        raise HTTPException(500, "ANTHROPIC_API_KEY 환경변수 없음")

    # PDF 저장
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t:
        t.write(await file.read())
        pdf_path = t.name

    d  = tempfile.mkdtemp()
    xl = os.path.join(d, "보장진단.xlsx")
    zp = os.path.join(d, f"보장분석_{datetime.datetime.now():%H%M%S}.zip")

    try:
        # Claude Vision 분석
        data = analyze_with_claude(pdf_path)
        # 엑셀 생성
        build_excel(data, TPL, xl)
    except Exception as e:
        raise HTTPException(500, f"분석오류: {e}")

    # ZIP 반환
    with zipfile.ZipFile(zp, "w") as z:
        z.write(xl, "보장진단.xlsx")

    return FileResponse(zp, filename="보장분석.zip", media_type="application/zip")
