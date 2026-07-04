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
.chsub {{ font-size:9.5pt; font-weight:800; color:{NAVY}; margin:3mm 0 1.5mm; }}
.chsub span {{ font-size:7.5pt; color:{GOLDD}; font-weight:700; margin-left:2mm; }}
.chv {{ width:100%; border-collapse:collapse; margin-bottom:1mm; }}
.chv td {{ border:0.4pt solid {LINE}; padding:1.6mm 2.2mm; font-size:8pt; color:{INK}; line-height:1.35; vertical-align:top; }}
.chv td.g {{ width:32mm; background:#F4F1E8; font-weight:800; color:{NAVY}; }}
.chv td.g span {{ font-weight:700; color:{GOLDD}; font-size:7pt; }}
.chv2 {{ width:100%; border-collapse:collapse; margin:1mm 0; }}
.chv2 td {{ border:0.4pt solid {LINE}; padding:1.4mm 2.2mm; font-size:8pt; }}
.chv2 td.v {{ text-align:right; font-weight:800; color:{NAVY}; }}
.chnote {{ margin-top:2.5mm; font-size:7pt; color:{MUT}; line-height:1.4; background:#F4F1E8; border-left:1.5pt solid {GOLD}; padding:2mm 3mm; }}
.wsg {{ font-size:9pt; font-weight:800; color:#fff; background:{NAVY}; padding:1.6mm 3mm; margin:3mm 0 0; border-radius:1mm 1mm 0 0; }}
.ws {{ width:100%; border-collapse:collapse; margin-bottom:1mm; }}
.ws td {{ border:0.4pt solid {LINE}; padding:2mm 2.5mm; font-size:8.5pt; }}
.ws td.wl {{ width:34mm; font-weight:800; color:{NAVY}; }}
.ws td.wd {{ color:{MUT}; font-size:7.5pt; }}
.ws td.wc {{ width:34mm; text-align:right; }}
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
  <div class="smxcap">● = 보장. 오른쪽 담보일수록 커버 범위가 넓습니다 (허혈성·뇌출혈 &lt; 2대·뇌졸중 &lt; 순환계·뇌혈관). <b>빈맥(I47·48) = 마스터 무행·전 묶음 제외 → 본 표 미기재</b>(고정사실). <b>산정특례</b> = 진단 기반 <b>별개 개별 담보</b>(각각 보상). 대상 코드범위 <b>[심] I20~50·판막 / [뇌] I60~69·Q28·S06</b> 전체. 지급조건·기간은 회사·약관별 [확인]. 근거 현대해상 교육자료.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 5 / 7</span></div>
</div>

<!-- P6: 주요치료비 변천사 (교육) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>6</b>주요치료비 변천사</div><div class="bar"></div></div>
 <div class="body">
  <div class="sect">3대 주요치료비 담보의 변천사 <span>①비례형 → ②정액형 → ③비급여 → 생활비</span></div>
  <div class="chsub">■ 암 주요치료비 · 4세대 진화</div>
  <table class="chv"><tbody>
   <tr><td class="g">① 비례형(구간형)</td><td>치료비 발생 구간 하한 도달 시 그 구간 정액 지급(지출 없으면 미지급). <b>2024.11 단종</b>·재가입 불가 — 보유자 해지 금지. 실손과 중복 수령 가능.</td></tr>
   <tr><td class="g">② 정액형 <span>24.11~</span></td><td>치료 사실만으로 약정금 정액 지급(실비 무관). 암수술·항암방사선·항암약물 각 보장. 면책 90일·감액 일부. <b>가입금액 100만원부터</b>.</td></tr>
   <tr><td class="g">③ 비급여(하이클래스) <span>25.03~</span></td><td>급여 전액본인부담금 + 비급여 치료비 정액. 산정특례가 안 주는 <b>양성자·중입자·표적·면역</b> 커버. 2천만×10년=2억. 25.08~ 생활비형 추가.</td></tr>
   <tr><td class="g">④ 왜 필요한가</td><td>5세대 실손 비급여 = 자기부담 50%·통원 회당 5만·한도 1,000만(축소). 표적·면역 고가 약값엔 부족 → 주요치료비·하이클래스가 메운다.</td></tr>
  </tbody></table>
  <div class="chsub">■ 뇌·심장 주요치료비 · 암과 동일 진화</div>
  <table class="chv"><tbody>
   <tr><td class="g">① 비례형(병원비형)</td><td>급여 본인부담 기준(100만 미만 면책). 종합병원 2대질병 주요치료비 = 본인부담 연 100만↑ → 3,000만 한도·10년. <b>~24.11 단종</b>.</td></tr>
   <tr><td class="g">② 정액형(2대) <span>24.11~</span></td><td>수술·혈전용해·중환자실 치료 사실만으로 정액(100만~). 메리츠 최초, 연 2천만×5년=1억.</td></tr>
   <tr><td class="g">③ 확대(순환계)·생활비 <span>25.01~</span></td><td>부정맥·심부전·동맥류 확대(52질환·연 8천만·100세). 생활비형 = 치료 하나만 받아도 연 1회.</td></tr>
  </tbody></table>
  <div class="chsub">■ 재발률 · 실제 수술비용 <span style="font-size:8pt;color:{MUT}">(심평원·건보공단)</span></div>
  <table class="chv2"><tbody>
   <tr><td>뇌경색 재발</td><td>1년 10% · 5년 20~30%</td><td>심장수술(개흉)</td><td class="v">2,832만</td></tr>
   <tr><td>급성심근경색 재발</td><td>1년 30% · 3년 50%</td><td>관상동맥우회술</td><td class="v">2,690만</td></tr>
   <tr><td>스텐트 5년 재협착</td><td>15%</td><td>뇌동맥류 코일색전술</td><td class="v">1,100~1,600만</td></tr>
  </tbody></table>
  <div class="chnote">세대 판별 = 계약일 기준. ~24.11 비례형(단종·유지) / 24.11~ 정액형(100만~) / 25~ 비급여·생활비. <b>비례형 보유자 = 단종 담보, 해지 금지</b>. 치료비 담보 설명 전용(진단비 축과 별개)·구간액 회사별 상이 [확인]. 근거 교육자료 2607 + 심평원·건보공단(2026.07).</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 6 / 7</span></div>
</div>

<!-- P7: 상담 워크시트 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>7</b>상담 워크시트</div><div class="bar"></div></div>
 <div class="body">
  <div class="sect">지금 고객의 3대 주요치료비는? <span>WORKSHEET</span></div>
  <div class="wsg">암</div>
  <table class="ws"><tbody>
   <tr><td class="wl">진단비</td><td class="wd">걸렸을 때 일시금(기본)</td><td class="wc">{_ws_amt(rep,'일반암')}</td></tr>
   <tr><td class="wl">암 주요치료비</td><td class="wd">수술·방사선·약물 정액(100만~)</td><td class="wc">{_ws_ch(rep,'암주요치료비')}</td></tr>
   <tr><td class="wl">하이클래스(비급여)</td><td class="wd">표적·면역·중입자 전액본인</td><td class="wc">{_ws_ch(rep,'비급여주요치료비')}</td></tr>
   <tr><td class="wl">암 생활비</td><td class="wd">치료 중 소득보상</td><td class="wc">□ 만원</td></tr>
  </tbody></table>
  <div class="wsg">뇌·심장</div>
  <table class="ws"><tbody>
   <tr><td class="wl">뇌·심 진단비</td><td class="wd">뇌혈관·허혈성 일시금</td><td class="wc">{_ws_amt(rep,'급성심근경색')}</td></tr>
   <tr><td class="wl">2대 주요치료비</td><td class="wd">수술·혈전용해·중환자실(100만~)</td><td class="wc">{_ws_ch(rep,'순환계주요치료비')}</td></tr>
   <tr><td class="wl">순환계 주요치료비</td><td class="wd">부정맥·심부전 확대</td><td class="wc">□ 만원</td></tr>
   <tr><td class="wl">순환계 생활비</td><td class="wd">치료 중 소득보상</td><td class="wc">□ 만원</td></tr>
  </tbody></table>
  <div class="wsg">산정특례 — 뇌·심 각각 개별 담보 · 진단만으로 지급</div>
  <table class="ws"><tbody>
   <tr><td class="wl">산정특례(뇌)</td><td class="wd">뇌혈관질환 I60~69 · Q28 · S06</td><td class="wc">{_ws_ch(rep,'산정특례(뇌혈관)')}</td></tr>
   <tr><td class="wl">산정특례(심장)</td><td class="wd">심혈관질환 I20~50 · 판막 전체</td><td class="wc">{_ws_ch(rep,'산정특례(심장)')}</td></tr>
  </tbody></table>
  <div class="chnote">✔ = 보유 / ✘ = 미가입 → 보완 추천. 진단비(일시금)와 주요치료비(치료 실비·정액)는 <b>별개 축</b>이라 둘 다 필요하다. 산정특례는 진단만으로 지급되는 개별 담보다.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 7 / 7</span></div>
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
