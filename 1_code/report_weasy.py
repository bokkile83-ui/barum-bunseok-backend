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

_SCOPE_HEART = [
    ('I20  협심증', [0,1,1,1,1]),
    ('I21  급성심근경색증', [1,1,1,1,1]),
    ('I22  후속 심근경색증', [1,1,1,1,1]),
    ('I23  급성심근경색 후 합병증', [1,1,1,1,1]),
    ('I24  기타 급성 허혈성 심장병', [0,1,1,1,1]),
    ('I25  만성 허혈성 심장병', [0,1,1,1,1]),
    ('I30·31  급성 심장막염', [0,0,1,1,1]),
    ('I33  급성·아급성 심내막염', [0,0,1,1,1]),
    ('I34~37  심장 판막질환', [0,0,1,1,1]),
    ('I40·41  심근염', [0,0,1,1,1]),
    ('I42·43  심근병증', [0,0,0,1,1]),
    ('I47  발작성 빈맥', [0,0,1,1,1]),
    ('I48  심방세동·조동', [0,0,1,1,1]),
    ('I49  기타 심장 부정맥', [0,0,0,1,1]),
    ('I50  심부전', [0,0,1,1,1]),
    ('I70·71  대동맥 죽상경화·동맥류', [0,0,0,1,1]),
    ('I00~09  류마티스 심장질환', [0,0,0,1,1]),
    ('Q20~25  선천성 심장기형', [0,0,0,0,1]),
]
_SCOPE_BRAIN = [
    ('I60  거미막하 출혈', [1,1,1,1]),
    ('I61  뇌내출혈', [1,1,1,1]),
    ('I62  기타 비외상성 두개내출혈', [1,1,1,1]),
    ('I63  뇌경색증', [0,1,1,1]),
    ('I65  뇌전동맥 폐색·협착', [0,1,1,1]),
    ('I66  대뇌동맥 폐색·협착', [0,1,1,1]),
    ('I64  출혈·경색 미분류 뇌졸중', [0,0,1,1]),
    ('I67  기타 뇌혈관질환', [0,0,1,1]),
    ('I68  달리분류 뇌혈관장애', [0,0,1,1]),
    ('I69  뇌혈관질환 후유증', [0,0,1,1]),
    ('Q28.0~28.3  순환계 선천기형', [0,0,0,1]),
    ('S06  두개내 손상', [0,0,0,1]),
]
_HCOLS = [('급성심근경색','#C0392B','#F7E0DC'), ('허혈성','#B9540B','#FBEADB'), ('2대주요','#1E7A46','#E4F0EA'), ('순환계','#1F5FA8','#E6F1FB'), ('산정특례','#9A7A12','#FBF1D8')]
_BCOLS = [('뇌출혈','#C0392B','#F7E0DC'), ('뇌졸중','#1E7A46','#E4F0EA'), ('뇌혈관·순환계','#1F5FA8','#E6F1FB'), ('산정특례','#9A7A12','#FBF1D8')]

def _scope_table(title, rows, cols):
    th = ''.join(f'<th style="background:{bg};color:{fg}">{_html.escape(nm)}</th>' for nm, fg, bg in cols)
    body = ''
    for nm, marks in rows:
        tds = ''
        for i, m in enumerate(marks):
            dot = f'<span style="color:{cols[i][1]};font-size:8pt;font-weight:700">●</span>' if m else ''
            tds += f'<td>{dot}</td>'
        body += f'<tr><td class="nm">{_html.escape(nm)}</td>{tds}</tr>'
    _colg = '<col class="scn">' + ('<col>' * len(cols))
    return (f'<div class="smxh">{_html.escape(title)}</div>'
            f'<table class="smx"><colgroup>{_colg}</colgroup>'
            f'<thead><tr><th style="text-align:left">질병분류 (KCD 코드)</th>{th}</tr></thead><tbody>{body}</tbody></table>')

def _ws_ch(rep, key):
    for c in rep.get('chiryo',[]):
        if c.get('name')==key:
            return ('<b style="color:#C0444C">✘ 미가입</b>' if c.get('value')=='미가입'
                    else f'<b style="color:#1F7A4D">✔ {_html.escape(str(c.get("value")))}</b>')
    return '□ ___ 만원'

def _ws_amt(rep, name):
    return '✔ 가입 · ___ 만원'

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

    # ── CI 선지급 분석 블록 (rep['ci'], 없으면 빈 문자열) ──
    ci=rep.get('ci',{'present':False})
    if ci.get('present'):
        _chips=''.join(f'<span class="ci-it"><b>{_html.escape(i["t"])}</b> {_html.escape(i["v"])}</span>' for i in ci['items'])
        _rate=ci.get('rate',0); _rem=max(0,100-_rate)
        ci_html=(f'<div class="sect">CI · 생명보험 선지급 분석 <span>CRITICAL ILLNESS</span></div>'
                 f'<div class="ci-wrap">'
                 f'<div class="ci-top"><div class="ci-rate">선지급 <b>{_rate}%</b> 형</div>'
                 f'<div class="ci-desc">CI 사망보장 <b>{_html.escape(ci["samang"])}</b> · 중대질병 진단 시 {_rate}% 선지급, 잔여 {_rem}%는 사망 시 지급</div></div>'
                 f'<div class="ci-bar"><div class="ci-fill" style="width:{_rate}%"></div>'
                 f'<span class="ci-l">선지급 {_rate}%</span><span class="ci-r">잔여 {_rem}%</span></div>'
                 f'<div class="ci-items">{_chips}</div>'
                 f'<div class="ci-res">진단 후 잔여 사망보장<b>{_html.escape(ci["residual"])}</b></div>'
                 f'</div>')
    else:
        ci_html=''

    # ── Plan B: 비CI 진단비 정액 지급 구조 (CI 미보유 시 P3 상단 대체) ──
    noci=rep.get('noci',{'present':False})
    if (not ci.get('present')) and noci.get('present') and noci.get('items'):
        _nchips=''.join(f'<span class="ci-it"><b>{_html.escape(i["t"])}</b> {_html.escape(i["v"])}</span>' for i in noci['items'])
        noci_html=(f'<div class="sect">진단비 정액 지급 구조 <span>DIAGNOSIS · FIXED PAYOUT</span></div>'
                   f'<div class="ci-wrap">'
                   f'<div class="ci-top"><div class="ci-rate">비CI · <b>정액형</b></div>'
                   f'<div class="ci-desc">CI(선지급형) 미보유 · 암·뇌·심 진단비는 진단 즉시 <b>정액 100%</b> 지급(선지급·차감·잔여 없음)</div></div>'
                   f'<div class="ci-items" style="padding-top:3mm">{_nchips}</div>'
                   f'<div class="ci-res">CI형과 달리 사망보험금 차감이 없어 <b>진단금 전액</b>이 치료비로 쓰이고, 사망보장은 별도로 유지됩니다.</div>'
                   f'</div>')
    else:
        noci_html=''

    # ── 보장 진단 코멘트 (P3 하단 채움 + 설명서 성격) ──
    _sh=[_html.escape(s['h']) for s in rep.get('strength',[])][:4]
    _cm=f'{cust} 고객님은 보유 <b>{n_contract}건</b> · 월 <b>{premium:,}원</b>의 보장을 운용하고 있습니다. '
    if _sh: _cm+='특히 <b>'+' · '.join(_sh)+'</b> 영역의 핵심담보를 충실히 보유했습니다. '
    if ci.get('present'):
        _cm+=f'CI 선지급형 보유로 암·뇌졸증·급성심근 등 중대질병 진단 시 진단자금이 즉시 지급되며, 진단 후에도 잔여 사망보장 <b>{_html.escape(ci.get("residual","-"))}</b>이 유지됩니다. '
    _cm+='보강이 필요한 공백 영역은 <b>'+f'{gap_cnt}개</b>로 '+('전반적으로 보장 균형이 양호합니다.' if gap_cnt==0 else '상담을 통한 보완을 권장합니다.')
    comment_html=f'<div class="sect" style="margin-top:5mm">보장 진단 코멘트 <span>SUMMARY</span></div><div class="cmt">{_cm}</div>'

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
.sbody {{ padding:5mm 8mm; }}
.jc {{ display:flex; gap:4mm; }}
.jcol {{ flex:1; }}
.jhd {{ font-size:11pt; font-weight:800; color:{NAVY}; border-bottom:2pt solid {GOLD}; padding-bottom:1mm; margin-bottom:2mm; }}
.jhd .sub {{ float:right; font-size:7.5pt; color:{MUT}; font-weight:600; margin-top:2mm; }}
.jstages {{ display:flex; gap:1.4mm; margin-bottom:2.5mm; }}
.jstage {{ flex:1; border:0.5pt solid {LINE}; border-radius:1.4mm; padding:1.2mm; text-align:center; font-size:5.8pt; line-height:1.25; color:{INK}; }}
.jstage b {{ display:block; font-size:6.8pt; color:{NAVY}; }}
.jstage .dt {{ color:{GOLD}; font-weight:700; font-size:5.6pt; }}
.jtwo {{ display:flex; gap:1.6mm; margin-bottom:2mm; }}
.jt {{ width:100%; border-collapse:collapse; font-size:6.3pt; }}
.jt th {{ background:{NAVY}; color:#fff; padding:0.7mm 1.4mm; font-weight:700; }}
.jt td {{ border:0.4pt solid {LINE}; padding:0.7mm 1.4mm; }}
.jt td.p {{ color:{GAP}; font-weight:800; text-align:right; }}
.jclause {{ background:#F6F8FB; border:0.5pt solid {LINE}; border-left:2.4pt solid {NAVY}; border-radius:1.2mm; padding:1.4mm 2mm; margin-bottom:1.8mm; font-size:6.2pt; line-height:1.4; }}
.jclause b {{ color:{NAVY}; font-size:6.6pt; }}
.jclause li {{ margin:0.5mm 0 0.5mm 3mm; }}
.jwarn {{ background:#FBECEC; border:0.5pt solid #E4B4B4; border-radius:1.4mm; padding:1.6mm 2mm; margin-bottom:1.8mm; font-size:6.2pt; line-height:1.4; }}
.jwarn b {{ color:{GAP}; }}
.jwarn .h {{ font-size:7pt; font-weight:800; color:{GAP}; }}
.jtalk {{ background:#EAF2FB; border-left:2.4pt solid {BLUE}; border-radius:1.4mm; padding:1.6mm 2mm; font-size:6.3pt; line-height:1.45; color:{NAVY}; font-style:italic; }}
.jtalk .h {{ font-style:normal; font-weight:800; color:{BLUE}; font-size:7pt; }}
.jstep {{ display:flex; gap:1.6mm; align-items:flex-start; border:0.5pt solid {LINE}; border-radius:1.4mm; padding:1.2mm 1.6mm; margin-bottom:1.4mm; }}
.jstep .n {{ flex:0 0 auto; width:4.2mm; height:4.2mm; border-radius:50%; background:{GOLD}; color:#fff; font-size:6.5pt; font-weight:800; text-align:center; line-height:4.2mm; }}
.jstep .b {{ flex:1; font-size:6.2pt; line-height:1.35; }}
.jstep .b .t {{ font-weight:800; color:{NAVY}; font-size:6.8pt; }}
.jstep .b .d {{ float:right; color:{GOLD}; font-weight:700; font-size:6pt; }}
.jstep .b .g {{ color:{GOOD}; font-weight:700; }}
.wsname {{ text-align:center; margin:3mm 0 4mm; }}
.wsname .nm {{ display:inline-block; font-size:15pt; font-weight:800; color:#fff; background:linear-gradient(135deg,{NAVY},{NAVY2}); padding:2mm 10mm; border-radius:6mm; border:1.4pt solid {GOLD}; }}
.wsname .q {{ display:block; font-size:8pt; color:{MUT}; margin-top:1.5mm; }}
.ws2 {{ display:flex; gap:4mm; }}
.wscol {{ flex:1; }}
.wscap {{ font-size:10pt; font-weight:800; padding:1.4mm 3mm; border-radius:1.4mm 1.4mm 0 0; color:#fff; }}
.wscap.c {{ background:#B9540B; }} .wscap.h {{ background:#1F5FA8; }}
.wsr {{ border:0.5pt solid {LINE}; border-top:none; padding:1.6mm 3mm; font-size:7.4pt; line-height:1.35; }}
.wsr .t {{ font-weight:800; color:{NAVY}; font-size:8pt; }}
.wsr .t.on {{ color:{GOOD}; }} .wsr .t.off {{ color:{GAP}; }}
.wsr .d {{ color:{MUT}; font-size:6.8pt; }}
.wsr .amt {{ float:right; font-size:7.6pt; font-weight:800; color:{INK}; }}
.wslack {{ background:#FBF6E6; border:0.5pt solid #E6D08A; border-radius:1.4mm; padding:2mm 3mm; margin:3mm 0 2mm; font-size:7.4pt; }}
.wslack .h {{ font-weight:800; color:{GOLDD}; font-size:8pt; }}
.wslack .bl {{ display:inline-block; border-bottom:0.6pt solid {INK}; width:48mm; height:3mm; }}
.wstalk {{ background:#EAF2FB; border-left:2.6pt solid {BLUE}; border-radius:1.4mm; padding:2mm 3mm; font-size:7.6pt; line-height:1.5; color:{NAVY}; font-style:italic; }}
.wstalk .h {{ font-style:normal; font-weight:800; color:{BLUE}; }}
.ws3 {{ display:flex; gap:2.5mm; align-items:stretch; }}
.wsmid {{ flex:0 0 30mm; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; }}
.wsmid .lb {{ font-size:6.5pt; color:{MUT}; letter-spacing:1px; font-weight:700; }}
.wsmid .nm {{ font-size:12pt; font-weight:800; color:#fff; background:linear-gradient(135deg,{NAVY},{NAVY2}); border:1.2pt solid {GOLD}; border-radius:4mm; padding:2mm 4mm; margin:1.5mm 0; }}
.wsmid .cnt {{ font-size:13pt; font-weight:800; color:{NAVY}; }}
.wsmid .cnt small {{ display:block; font-size:6.2pt; color:{MUT}; font-weight:600; }}
.wssj {{ margin-top:2mm; }}
.wssj .cap {{ font-size:7.4pt; font-weight:800; color:{GOLDD}; padding:1mm 0; border-top:0.6pt solid {GOLD}; }}



.hcnote {{ background:{NAVY}; color:#EAF0F8; font-size:7.4pt; line-height:1.45; padding:2.4mm 3.5mm; border-radius:2mm; margin-bottom:3mm; }}
.hcnote b {{ color:{GOLDL}; }}
.coband {{ border:0.5pt solid {LINE}; border-radius:1.6mm; margin-bottom:2.6mm; overflow:hidden; }}
.colabel {{ background:linear-gradient(135deg,{NAVY},{NAVY2}); color:#fff; font-size:8.5pt; font-weight:800; padding:1.3mm 3mm; }}
.colabel span {{ color:{GOLDL}; font-weight:700; font-size:6.8pt; margin-left:2mm; }}
.grow {{ display:flex; flex-wrap:wrap; padding:1.4mm; gap:1.4mm; align-items:stretch; }}
.gcard {{ flex:1 1 0; min-width:33mm; border:0.4pt solid {LINE}; border-radius:1.2mm; overflow:hidden; }}
.gh {{ font-size:6.6pt; font-weight:800; padding:0.9mm 1.6mm; color:#fff; }}
.dl {{ font-size:5.6pt; line-height:1.28; padding:0.5mm 1.6mm; border-top:0.3pt solid #EEF1F5; color:{INK}; }}
.gcard.isch .gh {{ background:#1F5FA8; }} .gcard.isch {{ background:#F3F8FD; }}
.gcard.ami  .gh {{ background:#B9540B; }} .gcard.ami  {{ background:#FDF4EC; }}
.gcard.myo  .gh {{ background:#5B7A2E; }} .gcard.myo  {{ background:#F4F8EE; }}
.gcard.inf  .gh {{ background:#1E7A46; }} .gcard.inf  {{ background:#EEF7F1; }}
.gcard.arr  .gh {{ background:#9A7A12; }} .gcard.arr  {{ background:#FBF6E6; }}
.gcard.val  .gh {{ background:#6A4A9A; }} .gcard.val  {{ background:#F5F1FB; }}
.gcard.cir  .gh {{ background:#2A3F63; }} .gcard.cir  {{ background:#EEF1F6; }}
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
.smx {{ width:100%; border-collapse:collapse; table-layout:fixed; margin:0 0 2mm; }}
.smx th,.smx td {{ border:0.5pt solid {LINE}; padding:0.5mm 0.4mm; text-align:center; font-size:6.7pt; color:{INK}; }}
.smx td.nm {{ text-align:left; padding-left:1.5mm; color:{INK}; font-size:6.7pt; line-height:1.05; }}
.smx thead th {{ font-weight:700; font-size:6.3pt; line-height:1.1; }}
.smx .scn {{ width:34%; }}
.smxh {{ font-size:8.5pt; font-weight:800; color:{NAVY}; margin:1.5mm 0 1mm; }}
.smxcap {{ margin-top:1.5mm; font-size:6.5pt; color:{MUT}; line-height:1.35; }}
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
.ci-wrap {{ border:0.8pt solid {GOLD}; border-radius:4pt; overflow:hidden; margin-bottom:5mm; }}
.ci-top {{ background:{NAVY}; color:#fff; padding:3mm 4mm; }}
.ci-top .ci-rate {{ font-size:13pt; font-weight:800; }}
.ci-top .ci-rate b {{ color:{GOLDL}; font-size:17pt; }}
.ci-top .ci-desc {{ font-size:8.5pt; color:#C9D5E4; margin-top:1mm; }}
.ci-top .ci-desc b {{ color:{GOLDL}; }}
.ci-bar {{ position:relative; height:6mm; background:{NAVY2}; }}
.ci-bar .ci-fill {{ height:6mm; background:linear-gradient(90deg,{GOLD},{GOLDD}); }}
.ci-bar .ci-l {{ position:absolute; left:3mm; top:1.5mm; color:{NAVY}; font-size:8pt; font-weight:800; }}
.ci-bar .ci-r {{ position:absolute; right:3mm; top:1.6mm; color:#fff; font-size:8pt; font-weight:700; }}
.ci-items {{ padding:3mm 4mm 1mm; }}
.ci-items .ci-it {{ display:inline-block; font-size:9.5pt; padding:1.5mm 3mm; margin:0 2mm 2mm 0; border-radius:2pt; background:#FBF1D8; color:{NAVY}; }}
.ci-items .ci-it b {{ color:{GOLDD}; }}
.ci-res {{ padding:2.5mm 4mm 3mm; font-size:9.5pt; color:{INK}; border-top:0.5pt solid {LINE}; }}
.ci-res b {{ color:{NAVY}; font-size:12pt; margin-left:1mm; }}
.cmt {{ padding:4mm 4.5mm; background:#F6F8FB; border:0.5pt solid {LINE}; border-left:2.5pt solid {NAVY}; border-radius:3pt; font-size:9.5pt; line-height:1.7; color:{INK}; }}
.cmt b {{ color:{NAVY}; }}
.ft {{ position:absolute; bottom:0; left:0; width:100%; padding:4mm 11mm; background:{NAVY}; color:#9FB0C6; font-size:8pt; overflow:hidden; }}
.ft b {{ color:{GOLD}; font-weight:700; }}
.ft .r {{ float:right; }}
'''

    advice_html=''
    for _a in rep.get('advice',[]):
        advice_html += (f'<div class="note" style="margin-top:3mm;border-left-color:{GAP}">'
                        f'<b style="color:{NAVY}">■ {_html.escape(_a["t"])}</b><br>'
                        f'<span style="font-size:8.5pt">{_html.escape(_a["d"])}</span></div>')
    scope_heart = _scope_table('심장 — 질병코드별 담보 커버', _SCOPE_HEART, _HCOLS)
    scope_brain = _scope_table('뇌 — 질병코드별 담보 커버', _SCOPE_BRAIN, _BCOLS)
    _HEART_FULL = [
     ('한화손해보험','4가지',[
       ('심혈관Ⅰ','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','발작성 빈맥 I47','심방세동 및 조동 I48','기타 심장부정맥 I49','심부전 I50']),
       ('심혈관Ⅱ','ami',['급성 심근경색증 I21','후속 심근경색증 I22','급성 심근경색 후 특정 현존 합병증 I23','인공소생에 성공한 심장정지 I46.0']),
       ('심근병증','myo',['심근병증 I42','확장성 심근병증 I42.0','비후성 심근병증 I42.1','제한성 심근병증 I42.2','기타 심근병증 I42.8','상세불명의 심근병증 I42.9']),
       ('심혈관특정질환','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50','※ 기타 심장부정맥(I49) 제외']),
     ]),
     ('DB손해보험','4가지',[
       ('특정Ⅰ','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','급성 심장막염 I30','심장막의 기타질환 I31','달리 분류된 질환의 심장막염 I32','급성·아급성 심내막염 I33','상세불명 판막 심내막염 I38.9','급성 심근염 I40','달리 분류된 질환의 심근염 I41']),
       ('특정Ⅱ','ami',['급성 심근경색증 I21','후속 심근경색증 I22','급성 심근경색 후 특정 현존 합병증 I23','인공소생에 성공한 심장정지 I46.0']),
       ('특정Ⅲ','val',['판막질환 —','발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50']),
       ('순환계 3대질환','cir',['인공소생에 성공한 심장정지 I46.0','부정맥 I47~I49','심부전 I50','※ 세부 범위는 약관 참조']),
     ]),
     ('KB손해보험','5가지',[
       ('특정Ⅰ','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50']),
       ('특정Ⅱ','ami',['급성 심근경색증 I21','후속 심근경색증 I22','급성 심근경색 후 특정 현존 합병증 I23','인공소생에 성공한 심장정지 I46.0']),
       ('심근병증','myo',['심근병증 I42','확장성 심근병증 I42.0','비후성 심근병증 I42.1','제한성 심근병증 I42.2','기타 심근병증 I42.8','상세불명의 심근병증 I42.9']),
       ('심장판막질환','val',['판막질환 —','심내막염 I33','상세불명의 판막질환 I38','※ 세부 범위는 약관 참조']),
       ('기타심장부정맥(I49)','arr',['기타 심장부정맥 I49','※ I47·I48은 해당 담보 대상 아님']),
     ]),
     ('현대해상','6가지',[
       ('허혈성 심장질환','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25']),
       ('특정허혈 심장질환','ami',['급성 심근경색증 I21','후속 심근경색증 I22','급성 심근경색 후 특정 현존 합병증 I23']),
       ('특정Ⅰ','arr',['발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50']),
       ('특정Ⅱ','ami',['급성 심근경색증 I21','후속 심근경색증 I22','급성 심근경색 후 특정 현존 합병증 I23','인공소생에 성공한 심장정지 I46.0']),
       ('주요 심장염증','inf',['급성 심장막염 I30','심장막의 기타질환 I31','달리 분류된 질환의 심장막염 I32','급성·아급성 심내막염 I33','상세불명 판막 심내막염 I38.9','급성 심근염 I40','달리 분류된 질환의 심근염 I41']),
       ('특정 2대 + 기타심장부정맥(I49)','arr',['방실차단 2도 I44.1','완전방실차단 I44.2','기타·상세불명 심방실차단 I44.3','이중섬유속차단 I45.2','삼중섬유속차단 I45.3','기타 심장부정맥 I49','※ I47·I48은 대상 아님']),
     ]),
     ('NH농협손해보험','5가지',[
       ('심혈관질환특정Ⅰ','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','발작성 빈맥 I47','심방세동 및 조동 I48','기타 심장부정맥 I49','심부전 I50']),
       ('특정Ⅰ (기타부정맥 제외)','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50','※ I49 제외']),
       ('심근병증','myo',['심근병증 I42','달리 분류된 질환의 심근병증 I43']),
       ('주요 심장염증질환','inf',['급성 심낭염 I30','기타 심장막염 I31','달리 분류된 질환의 심장막염 I32','급성·아급성 심내막염 I33','상세불명의 심내막염 I38','급성 심근염 I40','달리 분류된 질환의 심근염 I41']),
       ('기타 심장부정맥','arr',['기타 심장부정맥 I49']),
     ]),
     ('흥국화재','실제 약관',[
       ('특정심혈관질환(기타부정맥제외)','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25','발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50']),
       ('특정심혈관질환(기타부정맥)','arr',['기타 심장부정맥 I49']),
       ('심근병증(허혈성제외)','myo',['심근병증 I42','달리 분류된 질환의 심근병증 I43']),
       ('주요심장염증질환','inf',['급성 심장막염 I30','심장막의 기타질환 I31','달리 분류된 질환의 심장막염 I32','급성·아급성 심내막염 I33','상세불명 판막의 심내막염 I38','급성 심근염 I40','달리 분류된 질환의 심근염 I41']),
       ('허혈성심질환진단비','isch',['협심증 I20','기타 급성 허혈심장질환 I24','만성 허혈성심장병 I25']),
     ]),
     ('롯데손해보험','5가지',[
       ('특정심장질환Ⅰ','ami',['급성 심근경색증 I21','후속 심근경색증 I22','급성 심근경색 후 특정 현존 합병증 I23','인공소생에 성공한 심장정지 I46.0']),
       ('특정심장질환Ⅱ','isch',['협심증 I20','기타 급성 허혈성심장질환 I24','만성 허혈성심장병 I25','주요 심장염증 I30~I41']),
       ('특정 15대 심장질환','val',['판막질환 I05~I09·I34~I39','심근병증 I42~I43','발작성 빈맥 I47','심방세동 및 조동 I48','심부전 I50']),
       ('기타 심장부정맥','arr',['기타 심장부정맥 I49']),
       ('방실차단 및 전도장애','arr',['방실차단 2도 I44.1','방실차단 3도 I44.2','기타·상세불명 방실차단 I44.3','이중섬유속차단 I45.2','삼중섬유속차단 I45.3']),
     ]),
     ('삼성화재 (허혈성심장질환)','6가지',[
       ('허혈성 심장질환','isch',['급성기: 급성 심근경색증(STEMI/NSTEMI) I21','후속기: 후속 심근경색증 I22','합병증: 급성 심근경색 후 합병증 I23','협심증 I20','기타급성: 기타 급성 허혈성심장질환 I24','만성: 만성 허혈성심장병 I25']),
     ]),
     ('메리츠화재 (허혈성심장질환)','6가지',[
       ('허혈성 심장질환','isch',['급성기: 급성 심근경색증(STEMI/NSTEMI) I21','후속기: 후속 심근경색증 I22','합병증: 급성 심근경색 후 합병증 I23','협심증 I20','기타급성: 기타 급성 허혈성심장질환 I24','만성: 만성 허혈성심장병 I25']),
       ('＋ 기존 심장진단','ami',['심장질환진단Ⅰ → 허혈성','심장질환진단Ⅱ → 급성심근','(별도 상품 병존)']),
     ]),
    ]
    _FMAP={c[0]:c for c in _HEART_FULL}
    def _fullco(name):
        co=_FMAP[name]
        cards=''
        for gname,cat,lines in co[2]:
            items=''.join(f'<div class="dl">{_html.escape(l)}</div>' for l in lines)
            cards+=f'<div class="gcard {cat}"><div class="gh">{_html.escape(gname)}</div>{items}</div>'
        return f'<div class="coband"><div class="colabel">{_html.escape(co[0])}<span>{_html.escape(co[1])}</span></div><div class="grow">{cards}</div></div>'
    def _fullpage(pgno, sub, cos, note):
        body=''.join(_fullco(n) for n in cos)
        return f'''<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>{pgno}</b>심혈관 담보 분류</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">보험사별 심혈관 담보 분류 <span>상담용 · BARUM 실제 판매담보 기준 {sub}</span></div>
  {note}
  {body}
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · {pgno} / 9</span></div>
</div>'''
    _n8='<div class="hcnote">★ <b>"특정Ⅰ·Ⅱ"는 회사마다 뜻이 다릅니다 — 라벨 말고 질병코드로 확인.</b> 흥국·롯데 특정Ⅰ=급성심근경색 / 한화·NH 특정Ⅰ=협심증·허혈·빈맥·부정맥·심부전 / DB 특정Ⅰ=협심증·허혈·염증 / KB 특정Ⅰ=협심증·허혈·빈맥·심부전 / 현대 특정Ⅰ=빈맥·심부전. 빈맥(I47·48)과 부정맥(I49)은 별개.</div>'
    _n9='<div class="hcnote">★ 삼성·메리츠는 허혈성심장질환을 6가지로 세분(급성기·후속·합병증·협심증·기타급성·만성). 롯데 특정심장Ⅰ=급성심근경색 / 흥국은 특정심혈관질환(기타부정맥제외)=협심·허혈·빈맥·심부전(급성심근 아님). 색: <b style="color:#1F5FA8">허혈·협심</b> / <b style="color:#B9540B">급성심근</b> / <b style="color:#5B7A2E">심근병</b> / <b style="color:#1E7A46">염증</b> / <b style="color:#9A7A12">부정맥·전도</b> / <b style="color:#6A4A9A">판막</b>.</div>'
    heart_chart = _fullpage(8,'① 손해보험', ['한화손해보험','DB손해보험','KB손해보험','현대해상'], _n8) + '\n' + \
                  _fullpage(9,'② 손해보험', ['NH농협손해보험','흥국화재','롯데손해보험','삼성화재 (허혈성심장질환)','메리츠화재 (허혈성심장질환)'], _n9)
    doc=f'''<!DOCTYPE html><html><head><meta charset="utf-8"><style>{css}</style></head><body>
<!-- P0: 표지 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm" style="font-size:23pt">{cust} <b>고객님</b> 보장 설명서</div>
  <div class="bar"></div></div>
 <div class="body" style="padding:0 11mm">
  <div style="margin-top:185mm;border-top:1.2pt solid {GOLD};padding-top:6mm">
   <div style="font-size:15pt;font-weight:800;color:{GOLDD};letter-spacing:0.5pt">MAKEONE</div>
   <div style="margin-top:9mm;color:{MUT};font-size:9.5pt;line-height:3.2">
    <span style="display:inline-block;border-bottom:0.6pt solid {INK};width:62mm">&nbsp;</span>&nbsp;&nbsp;( 이름 )<br>
    <span style="display:inline-block;border-bottom:0.6pt solid {INK};width:62mm">&nbsp;</span>&nbsp;&nbsp;( 직책 )
   </div>
  </div>
 </div>
</div>
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
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 1 / 7</span></div>
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
   <td><div class="dc w"><div class="h">! 보장 공백</div><ul>{li(rep["weak"]) if rep["weak"] else '<li style="color:#1F7A4D"><b>주요 공백 없음</b> — 핵심담보 균형이 양호합니다.</li>'}</ul></div></td>
  </tr></table>
  <div class="sect" style="margin-top:4mm">갱신 / 비갱신 구조 <span>RENEWAL</span></div>
  <table class="ren"><tr>
   <td><div class="rbox b"><div class="rh">갱신형 {renew}건<small>보험료 인상 가능</small></div>{prem_rows(rep["renew_list"],True)}</div></td>
   <td><div class="rbox k"><div class="rh">비갱신형 {nonrenew}건<small>만기까지 고정</small></div>{prem_rows(rep["nonrenew_list"],False)}</div></td>
  </tr></table>
  <div class="sect" style="margin-top:4mm">월 보험료 구성 <span>PREMIUM</span></div>
  <table class="pbar">{bars}</table>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 2 / 7</span></div>
</div>
<!-- P3: 핵심 보장 분석 (CI 선지급 + 주요 치료비) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>3</b>핵심 보장 분석</div><div class="bar"></div></div>
 <div class="body">
  {ci_html}{noci_html}
  <div class="sect"{' style="margin-top:5mm"' if (ci_html or noci_html) else ''}>주요 치료비 정리 <span>TREATMENT BENEFITS</span></div>
  <table class="ctab">{crows}</table>
  {comment_html}
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 3 / 7</span></div>
</div>
<!-- P4 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>4</b>부위별 충족률</div><div class="bar"></div></div>
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
  {advice_html}
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 4 / 7</span></div>
</div>
<!-- P5: 보장범위 안내 (뇌·심·순환계·산정특례) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>5</b>보장범위 안내</div><div class="bar"></div></div>
 <div class="body">
  <div class="sect">담보별 보장범위 <span>DISEASE-CODE COVERAGE</span></div>
  {scope_heart}
  {scope_brain}
  <div class="smxcap">● = 보장. 오른쪽 담보일수록 커버 범위가 넓습니다 (허혈성·뇌출혈 &lt; 2대·뇌졸중 &lt; 순환계·뇌혈관). <b>산정특례</b> = 진단 기반 <b>별개 개별 담보</b>(각각 보상). 대상 코드범위 <b>[심] I20~50·판막 / [뇌] I60~69·Q28·S06</b> 전체. 지급조건·기간은 회사·약관별 [확인]. 근거 현대해상 교육자료.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 5 / 7</span></div>
</div>

<!-- P6: 주요치료비 변천사 (juyo_a4_v7 상세본) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>6</b>주요치료비 변천사</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">3대 주요치료비 담보의 변천사 <span>병원비형 → 정액형 → 확대형 → 생활비 · 정액형 가입금액 100만원부터</span></div>
  <div class="jc">
   <div class="jcol">
    <div class="jhd">■ 암 주요치료비 <span class="sub">4세대 진화</span></div>
    <div class="jstages">
     <div class="jstage"><span class="dt">~24.11</span><b>①구간형</b>병원비 구간별·단종</div>
     <div class="jstage"><span class="dt">24.11~</span><b>②정액형</b>수술·방사선·약물 100만~</div>
     <div class="jstage"><span class="dt">25.03~</span><b>③하이클래스</b>비급여 전액본인</div>
     <div class="jstage"><span class="dt">25.08~</span><b>④생활비</b>치료+소득보상</div>
    </div>
    <div class="jtwo">
     <table class="jt"><tr><th>치료비 발생</th><th>지급</th></tr>
      <tr><td>300~500만</td><td class="p">300만</td></tr><tr><td>500~700만</td><td class="p">500만</td></tr>
      <tr><td>700~1,000만</td><td class="p">700만</td></tr><tr><td>1,000~2,000만</td><td class="p">1,000만</td></tr>
      <tr><td>2,000~3,000만</td><td class="p">2,000만</td></tr><tr><td>3,000~4,000만</td><td class="p">3,000만</td></tr></table>
     <table class="jt"><tr><th>치료비 발생</th><th>지급</th></tr>
      <tr><td>4,000~5,000만</td><td class="p">4,000만</td></tr><tr><td>5,000~6,000만</td><td class="p">5,000만</td></tr>
      <tr><td>6,000~7,000만</td><td class="p">6,000만</td></tr><tr><td>7,000~8,000만</td><td class="p">7,000만</td></tr>
      <tr><td>8,000~9,000만</td><td class="p">8,000만</td></tr><tr><td>9,000~1억</td><td class="p">9천~1억</td></tr></table>
    </div>
    <div class="jclause"><b>① 구간형 약관 요점</b><ul><li>실제 치료비 <b>지출액이 구간 하한 도달 시</b> 그 구간 정액 지급(지출 없으면 미지급)</li><li>최소 하한 미달분 미지급(삼성 1,000만~ / 농협 300만까지)</li><li>실손과 <b>중복 수령 가능</b> · 2024.11 단종(재가입 불가)</li></ul></div>
    <div class="jclause"><b>② 정액형 · ③ 하이클래스 요점</b><ul><li>정액형 = 치료 사실만으로 약정금 정액(실비 무관). 면책 90일·감액 일부. 현재 신규는 정액형만</li><li>하이클래스 = 급여 전액본인부담 + 비급여 정액. <b>양성자·중입자·표적·면역</b> 커버. 2천만×10년=2억(삼성 최초)</li></ul></div>
    <div class="jwarn"><div class="h">▲ 실손으론 부족 — 5세대 기준</div>비중증 비급여: 자기부담 50%·통원 회당 5만·한도 <b>1,000만</b>(4세대 5천만→축소). 중증(암·산정특례): 자기부담 30%·연 5천만이나 표적·면역 고가엔 부족 → <b>주요치료비·하이클래스가 메움</b>.</div>
    <div class="jtalk"><span class="h">● 상담 화법</span> &nbsp;"표적항암 한 달 수백만 원이 전액 본인부담이에요. 5세대 실손은 비급여가 1천만으로 줄었습니다. 구간형은 이제 못 드는 담보라 유지가 답이고, 없으면 정액형·하이클래스로 채워야 합니다."</div>
   </div>
   <div class="jcol">
    <div class="jhd">■ 뇌·심장 주요치료비 <span class="sub">암과 동일 진화</span></div>
    <div class="jstep"><div class="n">1</div><div class="b"><span class="d">~2024.11</span><span class="t">병원비형</span><br>뇌·허 급여본인부담 구간별. 뇌혈관·허혈성 병원비 비례(메리츠 100만↑/농협 50만↑) → <b>2024.11 단종</b></div></div>
    <div class="jstep"><div class="n">2</div><div class="b"><span class="d">2024.11~</span><span class="t">정액형(2대)</span><br><span class="g">수술·혈전용해·중환자실</span>(100만~). 메리츠 최초(일주일 1억). 연 2천만×5년=1억. 2세대=뇌혈관·허혈성</div></div>
    <div class="jstep"><div class="n">3</div><div class="b"><span class="d">2025.01~</span><span class="t">순환계</span><br><span class="g">부정맥·심부전·동맥류 확대</span>. 2대보다 넓음. 메리츠 통합 52질환·연 8천만·100세(2025.06)</div></div>
    <div class="jstep"><div class="n">4</div><div class="b"><span class="d">2025.06~</span><span class="t">순환계 생활비</span><br><span class="g">치료 하나만 받아도 생활비</span>. 수술·혈전용해·중환자실 중 하나만 받아도 연 1회 지급</div></div>
    <table class="jt" style="margin-bottom:2mm"><tr><th>유형</th><th>회사</th><th>가입</th></tr>
     <tr><td>포괄형</td><td>삼성·현대·메리츠·한화·롯데·동양</td><td class="p">2~3천만</td></tr>
     <tr><td>순환계(넓음)</td><td>DB·KB·흥국</td><td class="p">2~3천만</td></tr>
     <tr><td>특정순환계</td><td>NH농협</td><td class="p">2천만</td></tr></table>
    <div class="jclause"><b>① 뇌·심 구간형 약관 요점(급여 본인부담 기준)</b><ul><li>100만원 미만=미지급(면책) / 100만~3,000만=본인부담금 / 3,000만 이상=<b>3,000만 한도</b></li><li>담보 예) 종합병원 2대질병 주요치료비 = 급여 본인부담 연 100만↑ → 3,000만 한도·10년·맞춤간편고지</li><li>암(치료비 발생 기준)과 달리 뇌·심은 <b>급여 본인부담금 기준</b> · 회사·약관별 상이 → 약관 확인 · 2024.11 단종</li></ul></div>
    <div class="jwarn" style="background:#FDECEC"><div class="h">▲ 재발률 — 한 번으로 안 끝난다</div>뇌경색(허혈성): 5년 내 <b>20~30% 재발</b>(10명 중 2~3명). 심장 스텐트: 5년 재협착 <b>15%</b>(수명 85%), 주로 1~2년 내. 재발 시 또 치료비 → 주요치료비는 <b>재발·장기치료 반복 보장</b>.</div>
    <div class="jtalk"><span class="h">● 상담 화법</span> &nbsp;"뇌·심장은 30일은 국가가 돕지만, 스텐트 후 1년 약물, 뇌졸중 재활은 몇 년씩 갑니다. 게다가 5년 내 2~3명은 재발해요. 그 긴 치료비를 주요치료비가 반복해 채웁니다."</div>
   </div>
  </div>
  <div class="chnote">※ 세대 판별 = 계약일 기준. ~24.11 병원비형(단종·유지) / 24.11~ 정액형(100만~) / 25~ 확대형·생활비 · <b>정액형 가입금액 100만원부터</b> · 구간형(~24.11) 보유자 = 단종 담보, <b>해지 금지</b>. 근거 교육자료 2607 + 보험저널·뱅크샐러드·금융위(2026.07) · 치료비 담보 설명 전용(진단비 축과 별개) · 구간 지급액 회사별 상이 [확인].</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 6 / 7</span></div>
</div>
<!-- P7: 상담 워크시트 (FINAL版·3단 중앙이름+산정특례) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>7</b>상담 워크시트</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">지금 고객의 3대 주요치료비는? <span>3대 = 진단비 + 주요치료비 + 생활비 · 진단비와 주요치료비는 별개 축</span></div>
  <div class="ws3">
   <div class="wscol">
    <div class="wscap c">■ 암 주요치료비</div>
    <div class="wsr"><span class="amt">{_ws_amt(rep,'일반암')}</span><span class="t on">✔ 암 진단비</span><br><span class="d">걸렸을 때 일시금(기본)</span></div>
    <div class="wsr"><span class="amt">{_ws_ch(rep,'암주요치료비')}</span><span class="t off">✘ 암 주요치료비</span><br><span class="d">수술·방사선·약물 정액(100만~)</span></div>
    <div class="wsr"><span class="amt">{_ws_ch(rep,'비급여주요치료비')}</span><span class="t">하이클래스(비급여)</span><br><span class="d">표적·면역·중입자 전액본인</span></div>
    <div class="wsr"><span class="amt">□ 만원</span><span class="t">암 생활비</span><br><span class="d">치료 중 소득보상</span></div>
   </div>
   <div class="wsmid">
    <div class="lb">OUR CLIENT</div>
    <div class="nm">{cust}<br>고객님</div>
    <div class="cnt">2/8<small>3대 치료비 보유</small></div>
   </div>
   <div class="wscol">
    <div class="wscap h">■ 뇌·심장 주요치료비</div>
    <div class="wsr"><span class="amt">{_ws_amt(rep,'급성심근경색')}</span><span class="t on">✔ 뇌·심 진단비</span><br><span class="d">뇌혈관·허혈성 일시금</span></div>
    <div class="wsr"><span class="amt">{_ws_ch(rep,'순환계주요치료비')}</span><span class="t off">✘ 2대 주요치료비</span><br><span class="d">수술·혈전용해·중환자실(100만~)</span></div>
    <div class="wsr"><span class="amt">□ 만원</span><span class="t">순환계 주요치료비</span><br><span class="d">부정맥·심부전 확대</span></div>
    <div class="wsr"><span class="amt">□ 만원</span><span class="t">순환계 생활비</span><br><span class="d">치료 중 소득보상</span></div>
   </div>
  </div>
  <div class="wssj">
   <div class="cap">산정특례 — 뇌·심 각각 개별 담보 · 진단만으로 지급</div>
   <div class="ws2">
    <div class="wscol"><div class="wsr"><span class="amt">{_ws_ch(rep,'산정특례(뇌혈관)')}</span><span class="t">산정특례(뇌)</span><br><span class="d">뇌혈관질환 I60~69 전체 · Q28 · S06</span></div></div>
    <div class="wscol"><div class="wsr"><span class="amt">{_ws_ch(rep,'산정특례(심장)')}</span><span class="t">산정특례(심장)</span><br><span class="d">심혈관질환 I20~50 · 판막 전체</span></div></div>
   </div>
  </div>
  <div class="wslack"><span class="h">■ 부족한 담보 → 보완 추천</span> &nbsp; 보완 후보 추가: <span class="bl">&nbsp;</span><br>· 암 주요치료비 → 보완 권유 &nbsp;·&nbsp; 2대 주요치료비 → 보완 권유</div>
  <div class="wstalk"><span class="h">● 상담 화법</span> &nbsp;"{cust} 고객님, 진단비는 '걸렸을 때' 한 번, 주요치료비는 '치료받는 내내' 나옵니다. 비어 있는 주요치료비부터 채우면 3대가 완성됩니다."</div>
  <div class="chnote">✔=보유 / ✘=미가입 → 보완 추천. 정액형 가입금액 100만원부터 · 근거 교육자료 2607 + 보험저널 등 · 진단비 축과 별개.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 7 / 7</span></div>
</div>
</body></html>'''
    HTML(string=doc).write_pdf(out)
    return True

# ── 정기철 샘플 데이터 (실제는 build_excel 결과에서 매핑) ──
if __name__=='__main__':
    sample={
     'client':'김진구','branch':'온빛센터 바름지점','manager':'최은혜','title':'지점장','phone':'010-XXXX-XXXX',
     'n_contract':6,'premium':324470,'renew':1,'nonrenew':0,'gap_count':0,
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
     'donuts':[{'name':'암','pct':100},{'name':'운전자','pct':100},{'name':'실손·배상','pct':100},{'name':'수술비','pct':100},{'name':'뇌혈관','pct':100},
               {'name':'사망·후유','pct':100},{'name':'골절·화상','pct':100},{'name':'심장','pct':100},{'name':'입원·일당','pct':100},{'name':'응급실·독감','pct':50}],
     'donut_detail':[
       {'name':'암','have':'2억4,910만','rec':'1억','pct':100},
       {'name':'운전자','have':'5개','rec':'4개','pct':100},
       {'name':'실손·배상','have':'3개','rec':'3개','pct':100},
       {'name':'수술비','have':'6개','rec':'4개','pct':100},
       {'name':'뇌혈관','have':'5,200만','rec':'5,000만','pct':100},
       {'name':'사망·후유','have':'4억3,300만','rec':'1억5,000만','pct':100},
       {'name':'골절·화상','have':'4개','rec':'3개','pct':100},
       {'name':'심장','have':'7,200만','rec':'3,000만','pct':100},
       {'name':'입원·일당','have':'3개','rec':'3개','pct':100},
       {'name':'응급실·독감','have':'1개','rec':'2개','pct':50}],
     'band_label':'40대','age_known':False,
    }
    build_report_pdf(sample,'김진구_weasy.pdf')
    print('PDF 생성 완료')
