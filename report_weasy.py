# -*- coding: utf-8 -*-
"""weasyprint 기반 보장진단 리포트 PDF 생성 (chromium 불필요)"""
from weasyprint import HTML
import math, html as _html

NAVY="#0B2340"; NAVY2="#16365C"; GOLD="#C5A052"; GOLDL="#E6C878"; GOLDD="#9C7C32"
INK="#1C2430"; MUT="#6B7686"; LINE="#D9DEE6"; GOOD="#1F7A4D"; PART="#9C7C32"; GAP="#C0444C"
BLUE="#1456B0"

def _donut(pct, color, size=92, sw=11):
    r=(size-sw)/2; c=size/2; C=2*math.pi*r; off=C*(1-pct/100)
    return f'''<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}">
<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="#EAEEF4" stroke-width="{sw}"/>
<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="{color}" stroke-width="{sw}" stroke-linecap="round"
 stroke-dasharray="{C:.1f}" stroke-dashoffset="{off:.1f}" transform="rotate(-90 {c} {c})"/>
<text x="{c}" y="{c}" text-anchor="middle" dy="0.36em" font-size="21" font-weight="700" fill="{NAVY}" font-family="NanumSquareRound">{pct}%</text>
</svg>'''

def build_report_pdf(rep, out):
    """rep: 리포트 데이터 dict (아래 sample_rep 구조). out: 저장 경로(.pdf)"""
    cust=_html.escape(rep['client'])
    branch=_html.escape(rep.get('branch','')); mgr=_html.escape(rep.get('manager',''))
    title=_html.escape(rep.get('title','')); phone=_html.escape(rep.get('phone',''))
    n_contract=rep['n_contract']; premium=rep['premium']; renew=rep['renew']; nonrenew=rep['nonrenew']; gap_cnt=rep['gap_count']

    # ── 보장현황 카드 (영역별) ──
    badge={'full':('충실','#E4F0EA',GOOD),'part':('일부','#FBF1D8',GOLDD),'gap':('취약','#F7E4E6',GAP)}
    def cov_card(cat):
        bt,bbg,bc=badge[cat['status']]
        items=''.join(
            f'<span class="it {("bl" if it.get("blue") else "")} {("r" if it.get("none") else "")}">'
            f'{("" if it.get("none") else "<b>")}{_html.escape(it["t"])}{("" if it.get("none") else "</b>")}'
            f'{(" "+_html.escape(it["v"])) if it.get("v") else ""}</span>'
            for it in cat['items'])
        return f'''<td class="cov-cell"><div class="cov-h"><span class="cn">{_html.escape(cat["name"])}</span>
<span class="bd" style="background:{bbg};color:{bc}">{bt}</span></div><div class="items">{items}</div></td>'''
    cov=rep['coverage']; rows=''
    for i in range(0,len(cov),2):
        left=cov_card(cov[i]); right=cov_card(cov[i+1]) if i+1<len(cov) else '<td></td>'
        rows+=f'<tr>{left}{right}</tr>'

    # ── 강점/공백 ──
    def li(arr): return ''.join(f'<li><b>{_html.escape(a["h"])}</b> — {_html.escape(a["d"])}</li>' for a in arr)
    # ── 갱신/비갱신 ──
    def prem_rows(arr,blue):
        col=BLUE if blue else INK
        return ''.join(f'<div class="pr"><span>{_html.escape(c["nm"])}</span><b style="color:{col}">{c["v"]}</b></div>' for c in arr)
    # ── 보험료 막대 ──
    mx=max((c['amt'] for c in rep['premium_bars']),default=1)
    bars=''.join(
        f'''<tr><td class="bl">{_html.escape(c["nm"])}</td>
<td class="track-td"><div class="track"><div class="fill" style="width:{c["amt"]/mx*100:.1f}%;background:{(BLUE if c["renew"] else NAVY2)}"></div></div></td>
<td class="bv" style="color:{(BLUE if c["renew"] else NAVY)}">{c["amt"]:,}</td></tr>'''
        for c in rep['premium_bars'])
    # ── 도넛 ──
    def dcolor(p): return GOOD if p>=70 else (GOLDD if p>=40 else GAP)
    donuts=''.join(
        f'<td class="dcell"><div>{_donut(d["pct"],dcolor(d["pct"]))}</div><div class="dn">{_html.escape(d["name"])}</div></td>'
        for d in rep['donuts'])
    drows=''
    per=5
    for i in range(0,len(rep['donuts']),per):
        cells=''.join(
            f'<td class="dcell"><div>{_donut(d["pct"],dcolor(d["pct"]))}</div><div class="dn">{_html.escape(d["name"])}</div></td>'
            for d in rep['donuts'][i:i+per])
        # pad
        cells+='<td></td>'*(per-len(rep['donuts'][i:i+per]))
        drows+=f'<tr>{cells}</tr>'

    def bcolor(p): return GOOD if p>=70 else (GOLDD if p>=40 else GAP)
    band=rep.get('band_label','40대')
    brows=''.join(
        f'<tr><td class="bn">{_html.escape(d["name"])}</td><td>{_html.escape(d["have"])}</td>'
        f'<td>{_html.escape(d["rec"])}</td>'
        f'<td style="color:{bcolor(d["pct"])};font-weight:800">{d["pct"]}%</td></tr>'
        for d in rep.get('donut_detail',[]))
    age_warn='' if rep.get('age_known') else ' <b>[확인]</b> 고객 나이·성별 미추출 → 40대 표준밴드로 산정(추후 정밀화).'

    def _cv(v): return f'<span style="color:{GAP}">미가입</span>' if v=='미가입' else f'<b style="color:{NAVY}">{_html.escape(v)}</b>'
    crows=''.join(f'<td class="cc"><div class="cl">{_html.escape(c["name"])}</div><div class="cval">{_cv(c["value"])}</div></td>'
                  for c in rep.get('chiryo',[]))
    crows=f'<tr>{crows}</tr>' if crows else '<tr><td>—</td></tr>'

    css=f'''
@page {{ size:A4; margin:0; }}
* {{ margin:0; padding:0; box-sizing:border-box; font-family:'NanumSquareRound','Noto Sans CJK KR',sans-serif; }}
body {{ color:{INK}; }}
.pg {{ width:210mm; height:297mm; position:relative; page-break-after:always; background:#fff; }}
.pg:last-child {{ page-break-after:auto; }}
.top {{ background:linear-gradient(135deg,{NAVY},{NAVY2}); color:#fff; padding:9mm 11mm 8mm; position:relative; }}
.top .bar {{ position:absolute; left:0; bottom:0; width:100%; height:1.1mm; background:linear-gradient(90deg,{GOLD},{GOLDD} 55%,transparent); }}
.top .eb {{ font-size:9pt; letter-spacing:2px; color:{GOLDL}; font-weight:700; }}
.top .nm {{ font-size:20pt; font-weight:800; margin-top:2mm; }}
.top .nm b {{ color:{GOLDL}; }}
.top .pgn {{ position:absolute; right:11mm; top:9mm; text-align:right; font-size:9pt; color:#9FB0C6; }}
.top .pgn b {{ display:block; font-size:22pt; color:#fff; font-weight:400; }}
.body {{ padding:7mm 11mm; }}
.sect {{ font-size:12pt; font-weight:800; color:{NAVY}; border-bottom:1.5pt solid {NAVY}; padding-bottom:2mm; margin-bottom:4mm; }}
.sect span {{ font-size:8pt; color:{GOLDD}; font-weight:700; margin-left:3mm; }}
.meta {{ width:100%; border-collapse:collapse; border-top:0.5pt solid {LINE}; border-bottom:0.5pt solid {LINE}; margin-bottom:6mm; }}
.meta td {{ padding:4mm 5mm; border-right:0.5pt solid {LINE}; width:25%; }}
.meta td:last-child {{ border-right:none; }}
.meta .k {{ font-size:8.5pt; color:{MUT}; }}
.meta .v {{ font-size:15pt; font-weight:800; color:{NAVY}; margin-top:1mm; }}
.meta .v small {{ font-size:9pt; color:{MUT}; font-weight:400; }}
.cov {{ width:100%; border-collapse:collapse; }}
.cov-cell {{ width:50%; vertical-align:top; padding:3mm 4mm 3mm 0; border-bottom:0.5pt solid {LINE}; }}
.cov-h {{ margin-bottom:2mm; }}
.cov-h .cn {{ font-size:11pt; font-weight:700; color:{NAVY}; }}
.cov-h .bd {{ font-size:8pt; font-weight:700; padding:0.5mm 2.5mm; border-radius:8pt; float:right; }}
.items .it {{ display:inline-block; font-size:8.5pt; padding:0.8mm 2mm; margin:0.6mm 1mm 0.6mm 0; border-radius:2pt; background:#EEF1F5; color:{INK}; }}
.items .it b {{ color:{NAVY}; }}
.items .it.bl b {{ color:{BLUE}; }}
.items .it.r {{ color:{MUT}; background:transparent; border:0.5pt dashed {LINE}; }}
.diag {{ width:100%; border-collapse:collapse; }}
.diag td {{ width:50%; vertical-align:top; padding-right:5mm; }}
.diag td:last-child {{ padding-right:0; padding-left:5mm; }}
.dc {{ border:0.5pt solid {LINE}; border-radius:4pt; overflow:hidden; }}
.dc .h {{ padding:3mm 4mm; color:#fff; font-weight:700; font-size:11pt; }}
.dc.g .h {{ background:{GOOD}; }} .dc.w .h {{ background:{GAP}; }}
.dc ul {{ list-style:none; padding:3.5mm 4mm; }}
.dc li {{ font-size:9.5pt; line-height:1.5; margin-bottom:2mm; padding-left:3mm; position:relative; }}
.dc li b {{ color:{NAVY}; }}
.ren {{ width:100%; border-collapse:collapse; margin-top:2mm; }}
.ren td {{ width:50%; vertical-align:top; padding-right:5mm; }}
.ren td:last-child {{ padding-right:0; padding-left:5mm; }}
.rbox {{ border:0.5pt solid {LINE}; border-radius:4pt; overflow:hidden; }}
.rbox .rh {{ padding:2mm 4mm; color:#fff; font-weight:700; font-size:10.5pt; }}
.rbox.b .rh {{ background:{BLUE}; }} .rbox.k .rh {{ background:#2A3340; }}
.rbox .rh small {{ float:right; font-weight:400; font-size:8.5pt; }}
.rbox .pr {{ padding:1mm 4mm; font-size:9.5pt; overflow:hidden; }}
.rbox .pr span {{ float:left; }} .rbox .pr b {{ float:right; }}
.note {{ margin-top:4mm; padding:3.5mm 4mm; background:#F2EEE2; border-left:2.5pt solid {GOLD}; font-size:8.5pt; color:#5C5340; line-height:1.5; }}
.note b {{ color:{GOLDD}; }}
.pbar {{ margin-top:3mm; width:100%; border-collapse:collapse; }}
.pbar td {{ padding:1.2mm 0; vertical-align:middle; }}
.pbar td {{ padding:0.6mm 0; }}
.pbar .bl {{ width:26mm; text-align:right; font-size:9pt; font-weight:600; padding-right:3mm; }}
.pbar .track-td {{ width:130mm; }}
.pbar .track {{ height:5mm; background:#EEF1F5; border-radius:3mm; }}
.pbar .fill {{ height:5mm; border-radius:3mm; }}
.pbar .bv {{ width:24mm; text-align:right; font-size:9pt; font-weight:700; padding-left:3mm; }}
.dtab {{ width:100%; border-collapse:collapse; }}
.dcell {{ text-align:center; padding:1.5mm 0; }}

.dcell .dn {{ font-size:9pt; font-weight:700; color:{NAVY}; margin-top:1mm; }}
.legend {{ display:flex; justify-content:center; gap:7mm; margin:1.5mm 0 1mm; font-size:9pt; color:{INK}; }}
.legend span {{ display:flex; align-items:center; gap:2mm; }}
.legend i {{ width:4mm; height:4mm; border-radius:1mm; display:inline-block; }}
.sect2 {{ font-size:11.5pt; font-weight:800; color:{NAVY}; margin:3.5mm 0 1.5mm; border-bottom:1.5pt solid {GOLD}; padding-bottom:1.5mm; }}
.sect2 span {{ font-size:8.5pt; font-weight:600; color:{MUT}; letter-spacing:.5px; margin-left:2mm; }}
.btab {{ width:100%; border-collapse:collapse; font-size:9.5pt; }}
.btab th {{ background:{NAVY}; color:#fff; padding:2mm 3mm; text-align:center; font-weight:700; }}
.btab td {{ padding:1.2mm 3mm; text-align:center; border-bottom:.5pt solid {LINE}; color:{INK}; }}
.btab td.bn {{ text-align:left; font-weight:700; color:{NAVY}; }}
.ctab {{ width:100%; border-collapse:separate; border-spacing:2.5mm 0; table-layout:fixed; }}
.ctab .cc {{ background:#F6F8FB; border:0.8pt solid {LINE}; border-radius:2mm; padding:2mm 1.5mm; text-align:center; vertical-align:top; }}
.ctab .cl {{ font-size:8.5pt; color:{MUT}; font-weight:700; margin-bottom:1.5mm; }}
.ctab .cval {{ font-size:11pt; }}
.ft {{ position:absolute; bottom:0; left:0; width:100%; padding:4mm 11mm; background:{NAVY}; color:#9FB0C6; font-size:8pt; overflow:hidden; }}
.ft b {{ color:{GOLD}; font-weight:700; }}
.ft .r {{ float:right; }}
'''

    doc=f'''<!DOCTYPE html><html><head><meta charset="utf-8"><style>{css}</style></head><body>
<!-- P1 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>1</b>보장 현황</div><div class="bar"></div></div>
 <div class="body">
  <table class="meta"><tr>
   <td><div class="k">보유 계약</div><div class="v">{n_contract}<small> 건</small></div></td>
   <td><div class="k">월 납입보험료</div><div class="v">{premium:,}<small> 원</small></div></td>
   <td><div class="k">갱신 / 비갱신</div><div class="v">{renew}<small> / {nonrenew}</small></div></td>
   <td><div class="k">보장 공백</div><div class="v" style="color:{GAP}">{gap_cnt}<small> 영역</small></div></td>
  </tr></table>
  <div class="sect">보장 현황 <span>CATEGORY COVERAGE</span></div>
  <table class="cov">{rows}</table>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 1 / 3</span></div>
</div>
<!-- P2 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>2</b>AI 진단</div><div class="bar"></div></div>
 <div class="body">
  <div class="sect">AI 진단 요약 <span>DIAGNOSIS</span></div>
  <table class="diag"><tr>
   <td><div class="dc g"><div class="h">✓ 보유 강점</div><ul>{li(rep["strength"])}</ul></div></td>
   <td><div class="dc w"><div class="h">! 보장 공백</div><ul>{li(rep["weak"])}</ul></div></td>
  </tr></table>
  <div class="sect" style="margin-top:4mm">갱신 / 비갱신 구조 <span>RENEWAL</span></div>
  <table class="ren"><tr>
   <td><div class="rbox b"><div class="rh">갱신형 {renew}건<small>보험료 인상 가능</small></div>{prem_rows(rep["renew_list"],True)}</div></td>
   <td><div class="rbox k"><div class="rh">비갱신형 {nonrenew}건<small>만기까지 고정</small></div>{prem_rows(rep["nonrenew_list"],False)}</div></td>
  </tr></table>
  <div class="sect" style="margin-top:4mm">월 보험료 구성 <span>PREMIUM</span></div>
  <table class="pbar">{bars}</table>
  <div class="sect" style="margin-top:4mm">주요 치료비 정리 <span>TREATMENT BENEFITS</span></div>
  <table class="ctab">{crows}</table>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 2 / 3</span></div>
</div>
<!-- P3 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>3</b>부위별 충족률</div><div class="bar"></div></div>
 <div class="body">
  <div class="sect">부위별 충족률 <span>COVERAGE LEVEL</span></div>
  <table class="dtab">{drows}</table>
  <div class="legend">
    <span><i style="background:{GOOD}"></i>충실 70%↑</span>
    <span><i style="background:{GOLDD}"></i>보강권장 40–69%</span>
    <span><i style="background:{GAP}"></i>취약 40%↓</span>
  </div>
  <div class="sect2">충족률 산정 근거 <span>보유 ÷ {band} 권장</span></div>
  <table class="btab">
    <tr><th>영역</th><th>보유</th><th>권장({band})</th><th>충족률</th></tr>
    {brows}
  </table>
  <div class="note">※ <b>충족률 = 보유 ÷ 연령밴드 권장액 × 100</b> (상한 100%). 권장액은 업계 적정 가입금액 가이드(암 진단비 5천만~1억·뇌혈관 3천만~5천만·허혈성 심장 3천만 등) 기준이며 {band} 표준밴드를 적용했습니다. 운전자·실손·일당·응급실은 핵심담보 보유개수 기준입니다. 개인 소득·가족력에 따라 권장액은 상담을 통해 조정됩니다.{age_warn}</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 3 / 3</span></div>
</div>
</body></html>'''
    HTML(string=doc).write_pdf(out)
    return True

# ── 정기철 샘플 데이터 (실제는 build_excel 결과에서 매핑) ──
if __name__=='__main__':
    sample={
     'client':'정기철','branch':'온빛센터 바름지점','manager':'최은혜','title':'지점장','phone':'010-XXXX-XXXX',
     'n_contract':8,'premium':530554,'renew':5,'nonrenew':3,'gap_count':3,
     'coverage':[
      {'name':'사망·후유장애','status':'part','items':[{'t':'교통사망','v':'2.8억'},{'t':'상해사망','v':'5,100만'},{'t':'질병사망 없음','none':True}]},
      {'name':'암','status':'full','items':[{'t':'표적항암','v':'8,000만'},{'t':'일반암','v':'5,000만'},{'t':'고액암','v':'5,000만'}]},
      {'name':'뇌혈관','status':'part','items':[{'t':'뇌혈관진단','v':'1,500만'},{'t':'산정특례','v':'500만'},{'t':'뇌졸중 없음','none':True}]},
      {'name':'심장 (＋빈맥)','status':'gap','items':[{'t':'허혈성','v':'1,000만'},{'t':'급성심근','v':'500만'},{'t':'협심·빈맥 없음','none':True}]},
      {'name':'수술비','status':'full','items':[{'t':'뇌혈관수술','v':'1,700만'},{'t':'허혈성','v':'1,000만'},{'t':'심장','v':'700만'}]},
      {'name':'운전자','status':'full','items':[{'t':'합의금','v':'2억','blue':True},{'t':'변호사','v':'5,000만','blue':True},{'t':'대인','v':'3,000만','blue':True}]},
      {'name':'입원·일당','status':'gap','items':[{'t':'간병인','v':'15만','blue':True},{'t':'상해일당','v':'1만','blue':True},{'t':'질병일당 없음','none':True}]},
      {'name':'실손·일배책','status':'full','items':[{'t':'실손입원','v':'5,000만','blue':True},{'t':'통원','v':'20만','blue':True},{'t':'일상배상','v':'1억','blue':True}]},
      {'name':'골절·화상','status':'part','items':[{'t':'골절','v':'20만'},{'t':'화상진단','v':'40만'},{'t':'깁스','v':'10만'}]},
      {'name':'응급실·독감','status':'gap','items':[{'t':'응급실 없음','none':True},{'t':'독감 없음','none':True}]},
     ],
     'strength':[{'h':'암 보장','d':'표적항암 8,000만 포함 진단·치료비 두텁게 구성'},
                 {'h':'운전자','d':'합의금 2억·변호사 5,000만, 운전 리스크 종합 대비'},
                 {'h':'실손·배상','d':'실손입원 5,000만 + 일상배상 1억 안전망 확보'},
                 {'h':'수술비','d':'뇌혈관·허혈성·심장 수술비 별도 보유'}],
     'weak':[{'h':'심장 진단','d':'협심증·심부전·부정맥·빈맥 진단계열 전무'},
             {'h':'질병사망·후유','d':'질병 기인 사망·후유 보장 0'},
             {'h':'입원 일당','d':'질병일당·간호통합병동·중환자실 미가입'},
             {'h':'응급실·독감','d':'일상 빈발 담보 공백'}],
     'renew_list':[{'nm':'라이나생명','v':'41,900원'},{'nm':'AIA생명','v':'48,611원'},{'nm':'삼성화재','v':'41,672원'},{'nm':'현대해상','v':'27,450원'},{'nm':'흥국실손','v':'갱신'}],
     'nonrenew_list':[{'nm':'동양생명','v':'230,003원'},{'nm':'롯데손해','v':'134,605원'},{'nm':'흥국(0809)','v':'6,313원'}],
     'premium_bars':[{'nm':'동양생명','amt':230003,'renew':False},{'nm':'롯데손해','amt':134605,'renew':False},
                     {'nm':'AIA생명','amt':48611,'renew':True},{'nm':'라이나생명','amt':41900,'renew':True},
                     {'nm':'삼성화재','amt':41672,'renew':True},{'nm':'현대해상','amt':27450,'renew':True},{'nm':'흥국(0809)','amt':6313,'renew':False}],
     'donuts':[{'name':'암','pct':90},{'name':'운전자','pct':88},{'name':'실손·배상','pct':85},{'name':'수술비','pct':80},{'name':'뇌혈관','pct':55},
               {'name':'사망·후유','pct':50},{'name':'골절·화상','pct':45},{'name':'심장','pct':25},{'name':'입원·일당','pct':20},{'name':'응급실·독감','pct':5}],
    }
    build_report_pdf(sample,'정기철_weasy.pdf')
    print('PDF 생성 완료')
