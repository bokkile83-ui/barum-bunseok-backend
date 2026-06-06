# -*- coding: utf-8 -*-
import os, io, json, base64, zipfile, re, tempfile, traceback
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, HTMLResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import anthropic

SYSTEM_PROMPT = r"""너는 BARUM(최은혜 지점장) 보장분석 엔진이다. 첨부된 '보장분석 리포트' PDF를 계약별로 정확히 읽어 JSON만 출력한다. 설명 금지, 마크다운 금지, 순수 JSON만.

[★최우선 정상화 5규칙 — 아래 A~L보다 우선]
0(제목). customer는 반드시 PDF 본문 '계약자' 이름. 파일명(noname/복사 등) 사용 금지.
1(종신). 만기 9999 또는 상품명에 '종신' → 사망액을 '일반사망'과 '상해사망' 두 칸에 같은 금액 1:1 기재. 별첨 사망이 [1]/[2]·두 줄이면 첫째=일반사망, 둘째=상해사망. (예: NEW알뜰종신 6000/6000 → 일반6000+상해6000)
2(3대진단비). 접두접미 '[갱신형]·갱신계약·(갱신)·중대한'은 떼고 본담보로 매핑(renew/CI 표시만 유지). 암: 암진단비/일반암=일반암 · 갑상선/제자리/경계성/기타피부/소액/비침습=유사암 · 전립선/방광/특정=특정암 · 통합전이/전이=전이암진단비(대표1·합산금지). ★유사암칸에 수천만원이 들어가면 오분류이니 일반암으로 재확인. 뇌: 뇌혈관진단/뇌졸증/뇌출혈 구분. 심장: '심장질환(특정Ⅰ)'=허혈성진단비, '(특정Ⅱ)'=급성심근경색. 허혈성 가입시 급성심근·협심증 동액 중복채움.
3(수술). 1~5종은 질병/상해 구분해 surg15에 [1종,2종,3종,4종,5종] 배열. 16/116/119대 등 n대수술=대표 1개 ndae. 상해수술비·질병수술비는 별도 row.
4(갱신vs비갱신). 담보명에 '(갱신형)·[갱신형]·갱신계약' → 그 담보 renew=true(파랑). 계약 renewal: 납입기간==보장기간=갱신 / 9999·운전자·납입≠보장·납입회차>240=비갱신.

[법칙]
A. 분석순서: ①표지 다음 보유계약목록 인식 ②각 계약 별첨(상품별 보장현황)을 1장씩 분석 ③계약별 세로 컬럼에 정확히 기재. ★A담보가 B회사로 가면 안 됨 — 계약별 대조 필수.
B. 기재범위: 표준 폼 담보. 폼에 없는 생소담보는 rows에 넣지 말고 memo에 (type=제외).
C. 갱신/비갱신: 갱신=납입기간==보험기간. 비갱신=불일치 또는 납입>240회 또는 만기9999(종신). 9999=종신→일반사망+상해사망 1:1. 손보 질병사망=80세 고정.
D. 특약명에 '(갱신)' '[갱신형]' 포함 → 그 값만 renew=true (그 칸만 파랑).
E. 1~5종 수술비: 질병/상해 구분해 surg15에 회사별 [1종,2종,3종,4종,5종] 배열. (상급종합병원 1~5종은 제외)
F. n대수술비(16대/116대/119대 등) = 대표금액(최대값) 1개만 ndae에 회사별.
G. 통합암·통합전이암 = 대표금액 1개만 (전이암진단비 row, 합산 금지).
H. 하이클래스 = 비급여주요치료비 = 같은 담보로 인식.
I. 제외 담보: 5/7대 장기이식·조혈모세포이식·각막이식 → rows 제외, memo(제외).
J. 심장담보: 약관 포함관계 반영. 허혈성심장질환(I20~25)은 급성심근경색·협심증을 포함 → 허혈성 가입시 급성심근경색·협심증 값도 같은 금액으로 채움(중복보장). 심장 카테고리 row name은 다음 중 사용: 부정맥질환(I49), 심부전, 염증, 협심증, 빈맥, 심장병증, 심장판막, 급성심근경색, 허혈성, 심장수술비. '심장질환(특정Ⅰ)'=허혈성, '(특정Ⅱ)'=급성심근경색(증권확인 필요시 memo).
K. 카테고리(폼 고정 순서): 사망/후유장애/암/뇌혈관/심장/일당/입원비/수술비/운전자/골절/화상/화재치아/기타/실손. 운전자 끝=일상생활배상책임, 화상=응급실내원비·깁스치료비 포함, 기타=독감·간병인지원일당·재가급여·시설급여. 폼 행은 삭제 금지, 폼에 없는 담보만 memo로.
L. 보험사 표기: 계약별 회사명은 보유계약 '회사'칸 원문 그대로. 발행사(let:/롯데 등) 표기를 다른 계약에 전파 금지. 회사명 축약·통일 금지. 마스킹된 건 '미확인'.

[출력 JSON 스키마]
{
  "customer": "고객명(파일/내용 기반, 없으면 '고객')",
  "contracts": [
    {"col":"C","company":"라이나생명","product":"THE간편한건강","renewal":"갱신","premium":29240,"start":"2023.11.13","end":"2033.11.13","pay":"10년납","count":"31/120"}
  ],
  "rows": [
    {"category":"사망","name":"상해사망","values":{"K":20000,"J":24000},"renew":[]}
  ],
  "surg15": {"질병":{"E":[10,30,50,100,300]},"상해":{"L":[20,50,200,500,1000]}},
  "ndae": {"L":1000,"H":100},
  "memo": [{"type":"약관","target":"KB 심장 특정Ⅰ/Ⅱ","note":"특정Ⅰ=허혈성 가정, 증권확인"}]
}

규칙: col은 계약 순서대로 C,D,E... 부여. 금액 단위=만원(정수). values 키=계약 col. 모든 계약을 contracts에 포함. 폼 표준담보만 rows에. JSON만 출력."""

CHAT_PROMPT = r"""너는 BARUM 보장분석 검수 AI다. 아래는 방금 분석한 고객의 보장진단 결과(JSON)다.

[분석결과]
{{DATA}}

행동:
- ISTJ 톤(결론먼저·단정·간결, "~이다/~해야한다"). 위로·감성 금지. 숫자는 콤마.
- '팩폭' → 계약별 세로로 누락·중복·A담보가B회사 오귀속을 표로 솔직히 지적. 자화자찬 금지.
- '수정' → 어느 담보를 어느 회사 칸으로 옮겨야 맞는지 구체적으로.
- 일반 질문 → 위 결과 근거로 답. 모르면 모른다고.
- 심장 특정Ⅰ/Ⅱ는 증권확인 권고. 담보 삭제 제안 금지(추가만)."""


CATS=[
('사망',['일반사망(종신)','질병사망(80세)','질병사망(__세)','상해사망','교통상해사망']),
('후유장애',['상해후유3%','상해후유80%','질병후유3%','질병후유80%','교통후유장해']),
('암',['암진단비','암진단비(갱신)','유사암(갑.기.경.제)','3대암진단비','통합전이암진단비','중입자치료비','항암치료비','표적항암치료비','다빈치로봇수술비','양성자','세기조절','암수술','암주요치료비(상급병원)','하이클래스(비급여주요치료비)','암통원비(상급병원)','암일당','CAR-T세포치료비','13대항암치료비','면역항암치료비','호르몬항암치료비']),
('뇌혈관',['뇌혈관진단비','뇌졸증진단비','뇌졸증진단비(갱신)','뇌출혈진단비','산정특례(뇌혈관)','혈전용해치료비(뇌졸중)','순환계주요치료비','뇌혈관수술비']),
('심장',['부정맥질환(I49)','심부전','염증','협심증','빈맥','심근병증','심장판막','급성심근경색','허혈성','산정특례(심장)','순환계주요치료비','심장수술비']),
('일당',['질병일당','상해일당','질병중환자실','상해중환자실','1인실 상급병원일당','1인실 종합병원일당','간병인사용일당','간병인지원일당(질병)','간병인지원일당(상해)','요양병원일당','간호통합병동일당','암일당(입원)']),
('입원비',['상해입원','질병입원']),
('수술비',['상해수술비','질병수술비','7대질병 수술비','상해종수술비(1-5종)','질병종수술비(1-5종)','119대 수술비','골절수술비','5대골절수술비','중대상해수술비','화상수술비','창상봉합수술비']),
('운전자',['교통사고처리지원금','6주미만처리지원금','교통사고벌금(대인)','교통사고벌금(대물)','변호사선임비용','자동차부상위로금','일상생활배상책임']),
('골절',['골절진단비','골절진단비(치아제외)','5대골절진단비','5대골절수술비']),
('화상',['화상진단비','중대화상진단비','응급실내원비(응급)','깁스치료비']),
('화재/치아',['화재벌금/화재보험','크라운','임플란트','치아 기타']),
('기타',['독감','간병인지원일당','재가급여','시설급여']),
('실손',['실손입원','실손통원','실손 약(처방조제)','MRI','도수치료','비급여주사'])]

BLUE='FF0000FF';RED='FFFF0000';WHITE='FFFFFFFF';BLACK='FF000000';GREEN='FF006400';GRAY='FFEAEAEA'
thin=Side('thin',color='FF999999');med=Side('medium',color='FF333333')
GRID=Border(top=med,bottom=med,left=thin,right=thin)

def build_excel(data, path):
    contracts={c['col']:c for c in data['contracts']}
    cols=[c['col'] for c in data['contracts']]
    colnum={c:i+3 for i,c in enumerate(cols)}
    matrix={}
    for r in data.get('rows',[]):
        for col,amt in r['values'].items():
            matrix.setdefault((r['category'],r['name']),{})[col]=(amt, col in (r.get('renew') or []))
    # 심장 허혈성→급성심근경색·협심증 중복채움
    for col in cols:
        h=matrix.get(('심장','허혈성'),{}).get(col)
        if h:
            for tgt in ['급성심근경색','협심증']:
                k=('심장',tgt); cur=matrix.setdefault(k,{}).get(col)
                if (not cur) or cur[0]<h[0]: matrix[k][col]=(h[0],h[1])
    surg=data.get('surg15',{}); ndae=data.get('ndae',{})

    wb=Workbook(); ws=wb.active; ws.title='고객 보장진단'
    ws['A1']=f"{data.get('customer','고객')} 고객님\n보장진단"; ws.merge_cells('A1:B7')
    ws['A1'].alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); ws['A1'].font=Font(bold=True,size=12)
    for c in cols:
        cn=colnum[c]; info=contracts[c]; gb=info.get('renewal','비갱신')
        completed = False
        try:
            cnt=info.get('count','/'); a,b=cnt.split('/'); completed = a.strip()==b.strip() and a.strip()!=''
        except: pass
        fill = GREEN if completed else (BLUE if gb=='갱신' else RED)
        h=ws.cell(row=1,column=cn,value=info.get('company','')); h.fill=PatternFill('solid',fgColor=fill)
        h.font=Font(color=WHITE,bold=True,size=9); h.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); h.border=GRID
        meta=[gb,info.get('premium',''),info.get('start',''),info.get('end',''),info.get('pay',''),info.get('count','')]
        for i,v in enumerate(meta):
            cc=ws.cell(row=2+i,column=cn,value=v); cc.alignment=Alignment(horizontal='center'); cc.border=GRID; cc.font=Font(size=8)
            if i==1 and isinstance(v,(int,float)): cc.number_format='#,##0'
    for i,lb in enumerate(['갱신구분','보험료','가입일','만기일','납입','회차']): ws.cell(row=2+i,column=17,value=lb).font=Font(size=8,italic=True)
    ws.cell(row=1,column=17,value='합계').font=Font(bold=True); ws.cell(row=1,column=18,value='메모').font=Font(bold=True)
    for rr in range(1,8): ws.cell(row=rr,column=17).border=GRID; ws.cell(row=rr,column=18).border=GRID

    def renew_of(c): return contracts[c].get('renewal')=='갱신'
    r=8
    for cat,rows in CATS:
        extra=sorted({k[1] for k in matrix if k[0]==cat and k[1] not in rows})
        allrows=rows+extra; start=r
        for rw in allrows:
            ws.cell(row=r,column=2,value=rw).font=Font(size=9); ws.cell(row=r,column=2).border=GRID; ws.cell(row=r,column=1).border=GRID
            for c in cols: ws.cell(row=r,column=colnum[c]).border=GRID
            ws.cell(row=r,column=18).border=GRID
            q=ws.cell(row=r,column=17); q.border=GRID; q.font=Font(bold=True,size=8); q.number_format='#,##0'
            if rw in ('상해종수술비(1-5종)','질병종수술비(1-5종)'):
                kind='상해' if '상해' in rw else '질병'; tot=[0]*5
                kk=surg.get(kind,{})
                for c in cols:
                    seq=kk.get(c)
                    if seq:
                        for j in range(5): tot[j]+=seq[j]
                        cell=ws.cell(row=r,column=colnum[c],value='/'.join(f'{x:,}' for x in seq))
                        cell.font=Font(color=(BLUE if renew_of(c) else BLACK),size=8); cell.alignment=Alignment(horizontal='center',shrink_to_fit=True)
                if any(tot): q.value='/'.join(f'{x:,}' for x in tot); q.number_format='General'
            elif rw=='119대 수술비':
                for c,a in ndae.items():
                    if c in colnum:
                        cell=ws.cell(row=r,column=colnum[c],value=a); cell.number_format='#,##0'
                        cell.font=Font(color=(BLUE if renew_of(c) else BLACK),size=8); cell.alignment=Alignment(horizontal='center')
                if ndae: q.value=f'=SUM(C{r}:P{r})'
            else:
                has=False
                for c in cols:
                    v=matrix.get((cat,rw),{}).get(c)
                    if v:
                        has=True; amt,rn=v
                        cell=ws.cell(row=r,column=colnum[c],value=amt); cell.number_format='#,##0'
                        cell.font=Font(color=(BLUE if (renew_of(c) or rn) else BLACK),size=8); cell.alignment=Alignment(horizontal='center')
                if has: q.value=f'=SUM(C{r}:P{r})'
            r+=1
        ws.merge_cells(start_row=start,start_column=1,end_row=r-1,end_column=1)
        a=ws.cell(row=start,column=1,value=cat); a.font=Font(bold=True,size=9)
        a.fill=PatternFill('solid',fgColor=GRAY); a.alignment=Alignment(horizontal='center',vertical='center'); a.border=GRID
    ws.column_dimensions['A'].width=6; ws.column_dimensions['B'].width=20
    for c in cols: ws.column_dimensions[ws.cell(row=1,column=colnum[c]).column_letter].width=7.2
    ws.column_dimensions['Q'].width=12; ws.column_dimensions['R'].width=8
    # 메모시트
    ms=wb.create_sheet('📋확인사항'); ms['A1']=f"{data.get('customer','고객')} 고객님 — 확인사항 (①기재 ②약관 ③PPT외제외 ④팩폭)"; ms['A1'].font=Font(bold=True,size=11)
    ms.append([]); ms.append(['구분','대상','내용'])
    for cc in ms[3]: cc.font=Font(bold=True); cc.fill=PatternFill('solid',fgColor=GRAY)
    for m in data.get('memo',[]): ms.append([m.get('type',''),m.get('target',''),m.get('note','')])
    ms.column_dimensions['A'].width=12; ms.column_dimensions['B'].width=40; ms.column_dimensions['C'].width=50
    wb.save(path)
    return path

INDEX_HTML = """<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>BARUM 보장분석</title><style>
*{box-sizing:border-box;font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif}
body{background:#0f1420;color:#e9edf5;margin:0;padding:18px;display:flex;justify-content:center}
.c{max-width:720px;width:100%}
h1{color:#c9a24b;font-size:20px;border-bottom:2px solid #c9a24b;padding-bottom:8px;margin:0 0 14px}
.box{background:#141b2b;border:1px solid #2a3550;border-radius:12px;padding:16px;margin-bottom:14px}
label{font-size:13px;color:#8a93a6;display:block;margin:10px 0 4px}
input{width:100%;background:#1b2436;border:1px solid #2a3550;color:#e9edf5;border-radius:8px;padding:11px;font-size:14px}
button{background:#c9a24b;color:#0f1420;border:0;border-radius:8px;padding:12px 16px;font-weight:700;font-size:14px;cursor:pointer}
button:disabled{opacity:.5}
#st{margin-top:12px;font-size:13px;white-space:pre-wrap;color:#e3b34d;word-break:break-word}
.err{color:#ff6b6b !important}
.bar{display:flex;gap:6px;flex-wrap:wrap;margin:6px 0 10px}
.bar button{background:#1b2436;border:1px solid #c9a24b;color:#c9a24b;border-radius:18px;padding:6px 12px;font-size:12px;font-weight:400}
#result{display:none}
#chat{background:#141b2b;border:1px solid #2a3550;border-radius:10px;height:300px;overflow-y:auto;padding:10px;margin-bottom:8px}
.msg{margin-bottom:9px;display:flex}.msg.u{justify-content:flex-end}
.bub{max-width:86%;padding:8px 11px;border-radius:11px;font-size:13px;line-height:1.5;white-space:pre-wrap;word-break:break-word}
.u .bub{background:#6d2e46;color:#fff}.a .bub{background:#1f293d}
.inp{display:flex;gap:7px}
#q{flex:1;background:#1b2436;border:1px solid #2a3550;color:#e9edf5;border-radius:9px;padding:10px;font-size:14px}
.hint{font-size:11px;color:#6b7488;margin-top:8px}
</style></head><body><div class="c">
<h1>BARUM 보장분석</h1>
<div class="box">
<label>접근 비밀번호</label><input id="pw" type="password" placeholder="1009">
<label>보장분석 리포트 PDF</label><input id="f" type="file" accept="application/pdf">
<button id="b" onclick="go()" style="width:100%;margin-top:14px">분석 시작</button>
<div id="st"></div>
<div class="hint">PDF 1개 → 계약별 분석 → 엑셀 ZIP + 아래에서 팩폭·수정 채팅. 결과는 증권 최종확인.</div>
</div>
<div class="box" id="result">
<b style="color:#c9a24b">검수 채팅 (팩폭 · 수정)</b>
<div class="bar"><button onclick="quick('팩폭')">🔍 팩폭</button><button onclick="quick('수정')">📝 수정</button><button onclick="quick('계약별로 무슨 담보가 어느 회사에 들어갔는지 요약')">📋 계약별 요약</button></div>
<div id="chat"></div>
<div class="inp"><input id="q" placeholder="질문… (팩폭/수정)" onkeydown="if(event.key==='Enter')send2()"><button onclick="send2()">전송</button></div>
</div></div>
<script>
let DATA=null, hist=[];
const st=document.getElementById('st');
function setSt(t,err){st.textContent=t;st.className=err?'err':''}
async function go(){
 const pw=document.getElementById('pw').value, f=document.getElementById('f').files[0], b=document.getElementById('b');
 if(!pw||!f){setSt('비밀번호와 PDF를 넣어주세요.',true);return}
 b.disabled=true; setSt('분석 중… 1~2분. 끊기면 아래에 이유가 표시됩니다.');
 const fd=new FormData(); fd.append('password',pw); fd.append('file',f);
 const ctrl=new AbortController(); const to=setTimeout(()=>ctrl.abort(),290000);
 try{
  const r=await fetch('/analyze',{method:'POST',body:fd,signal:ctrl.signal});
  clearTimeout(to);
  const txt=await r.text();
  let j=null; try{j=JSON.parse(txt)}catch(_){}
  if(!j){ setSt('실패 [HTTP '+r.status+' '+r.statusText+']\\n\\n응답이 JSON 아님:\\n'+txt.slice(0,700),true); b.disabled=false; return }
  if(!j.ok){ setSt('실패: '+(j.error||'(원인 메시지 없음)')+(j.raw?'\\n\\n[원문]\\n'+j.raw:'')+(j.trace?'\\n\\n[위치]\\n'+j.trace:''),true); b.disabled=false; return }
  DATA=j.data;
  const bin=atob(j.zip_b64); const arr=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
  const u=URL.createObjectURL(new Blob([arr],{type:'application/zip'}));
  const a=document.createElement('a'); a.href=u; a.download=j.filename||'barum.zip'; a.click();
  setSt('완료. ZIP 다운로드됨 (계약 '+((j.data.contracts||[]).length)+'건). 아래에서 팩폭하세요.');
  document.getElementById('result').style.display='block'; chatAdd('a','분석 끝. "팩폭" 누르면 누락·오귀속 점검.');
 }catch(e){
  clearTimeout(to);
  if(e.name==='AbortError') setSt('실패: 시간초과(약 5분). 분석이 오래 걸려 끊김 → MODEL을 claude-haiku-4-5-20251001로 바꾸거나 PDF 페이지 줄이기.',true);
  else setSt('실패: '+(e.name||'오류')+' / '+(e.message||'(메시지 없음)')+'\\n→ 네트워크 끊김/타임아웃 추정. Wi-Fi에서 재시도.',true);
 }
 b.disabled=false;
}
const chat=document.getElementById('chat');
function chatAdd(role,t){const d=document.createElement('div');d.className='msg '+(role==='u'?'u':'a');d.innerHTML='<div class="bub"></div>';d.querySelector('.bub').textContent=t;chat.appendChild(d);chat.scrollTop=chat.scrollHeight;return d}
function quick(t){document.getElementById('q').value=t;send2()}
async function send2(){
 const q=document.getElementById('q'); const text=q.value.trim(); if(!text||!DATA)return;
 chatAdd('u',text); q.value=''; const ld=chatAdd('a','…'); hist.push({role:'user',content:text});
 try{
  const r=await fetch('/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({data:DATA,messages:hist})});
  const j=await r.json(); const out=j.ok?j.reply:('오류: '+j.error);
  ld.querySelector('.bub').textContent=out; if(j.ok)hist.push({role:'assistant',content:out});
 }catch(e){ld.querySelector('.bub').textContent='오류: '+e.message}
 chat.scrollTop=chat.scrollHeight;
}
</script></body></html>
"""

app=FastAPI()
ACCESS_PW=os.environ.get('ACCESS_PW','1009')
MODEL=os.environ.get('MODEL','claude-sonnet-4-6')
def client():
    key=os.environ.get('ANTHROPIC_API_KEY')
    if not key: raise RuntimeError('ANTHROPIC_API_KEY 환경변수가 없습니다. Railway Variables에 추가하세요.')
    return anthropic.Anthropic(api_key=key)
@app.get('/health')
def health():
    return {'ok':True,'model':MODEL,'has_key':bool(os.environ.get('ANTHROPIC_API_KEY'))}
@app.get('/', response_class=HTMLResponse)
def home():
    return INDEX_HTML
@app.post('/analyze')
async def analyze(password: str = Form(...), file: UploadFile = File(...)):
    try:
        if password.strip()!=ACCESS_PW: return JSONResponse({'ok':False,'error':'비밀번호 오류 (기본 1009)'})
        pdf=await file.read()
        if not pdf: return JSONResponse({'ok':False,'error':'PDF가 비어있습니다.'})
        fname=re.sub(r'\\.pdf$','',(file.filename or ''),flags=re.I).strip()
        b64=base64.standard_b64encode(pdf).decode()
        # ★PyMuPDF(fitz) 텍스트 선추출 — OZ리포트뷰어 CID폰트 PDF는 이미지/직접읽기로 한글이 누락(→'미확인').
        #   fitz는 ToUnicode CMap으로 회사명·상품명·담보명 한글을 복구한다. 이 텍스트를 1순위 근거로 준다.
        pdf_text=''
        try:
            import fitz
            _doc=fitz.open(stream=pdf, filetype='pdf')
            _parts=[]
            for _i,_pg in enumerate(_doc):
                _parts.append(f'===== p{_i+1} =====\n'+_pg.get_text())
            pdf_text='\n'.join(_parts)[:60000]
            _doc.close()
        except Exception:
            pdf_text=''
        _user=[{'type':'document','source':{'type':'base64','media_type':'application/pdf','data':b64}}]
        if pdf_text.strip():
            _user.append({'type':'text','text':'아래는 같은 PDF에서 추출한 원문 텍스트다. 회사명·상품명·담보명·금액·날짜는 이 텍스트를 1순위 근거로 삼아라(PDF 이미지가 흐릿하거나 마스킹돼 보여도 이 텍스트의 한글을 신뢰). 텍스트에 회사명이 있으면 절대 \'미확인\'으로 쓰지 마라.\n\n===PDF 추출 원문===\n'+pdf_text})
        _user.append({'type':'text','text':'이 PDF를 법칙대로 계약별로 분석하라. 설명·사고과정·서론 절대 금지. JSON 객체 하나만 출력. 첫 글자는 { 이다.'})
        msg=client().messages.create(model=MODEL, max_tokens=20000, system=SYSTEM_PROMPT,
            messages=[{'role':'user','content':_user},{'role':'assistant','content':'{'}])
        raw='{'+''.join(b.text for b in msg.content if getattr(b,'type',None)=='text')
        cleaned=raw.replace('```json','').replace('```','')
        start=cleaned.find('{')
        if start<0: return JSONResponse({'ok':False,'error':'분석 JSON을 못 찾음','raw':raw[:1500]})
        try:
            data,_=json.JSONDecoder().raw_decode(cleaned[start:])
        except json.JSONDecodeError as je:
            return JSONResponse({'ok':False,'error':f'JSON 파싱 실패: {je}','raw':cleaned[start:start+1500]})
        m_cust=re.search(r'계약자\s*([가-힣][가-힣*]{1,5})', pdf_text or '')
        pdf_cust=m_cust.group(1).strip() if m_cust else ''
        bad=(not fname) or bool(re.search(r'noname|복사|file|보장|pdf|^\d', fname, re.I))
        if pdf_cust: data['customer']=pdf_cust          # ★PDF 계약자 최우선
        elif not bad: data['customer']=fname            # 깨끗한 파일명만 허용
        cust=re.sub(r'[^\w가-힣]','',str(data.get('customer','고객'))) or '고객'
        xlsx=os.path.join(tempfile.gettempdir(),f'보장진단_{cust}.xlsx')
        build_excel(data, xlsx)
        buf=io.BytesIO()
        with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
            z.write(xlsx, os.path.basename(xlsx))
            z.writestr('분석결과.json', json.dumps(data,ensure_ascii=False,indent=2))
        return JSONResponse({'ok':True,'data':data,'zip_b64':base64.b64encode(buf.getvalue()).decode(),'filename':f'barum_{cust}.zip','pptx_filled':False})
    except anthropic.APIStatusError as e:
        return JSONResponse({'ok':False,'error':f'Anthropic API 오류 {e.status_code}: {getattr(e,"message",str(e))}. 모델명({MODEL})/키 확인.'})
    except Exception as e:
        return JSONResponse({'ok':False,'error':str(e),'trace':traceback.format_exc()[-1500:]})
@app.post('/chat')
async def chat(req: dict):
    try:
        data=req.get('data',{}); messages=req.get('messages',[])
        sysp=CHAT_PROMPT.replace('{{DATA}}',json.dumps(data,ensure_ascii=False)[:12000])
        msg=client().messages.create(model=MODEL, max_tokens=1500, system=sysp, messages=messages)
        out=''.join(b.text for b in msg.content if getattr(b,'type',None)=='text')
        return {'ok':True,'reply':out}
    except Exception as e:
        return {'ok':False,'error':str(e)}
