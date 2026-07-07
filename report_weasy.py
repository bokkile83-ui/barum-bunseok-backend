# -*- coding: utf-8 -*-
"""weasyprint 기반 보장진단 리포트 PDF 생성 (chromium 불필요)"""
# fontTools OS/2 unicode-range 안전 패치 (색이모지 subset 크래시 방지)
try:
    import fontTools.ttLib.tables.O_S_2f_2 as _os2mod
    if not getattr(_os2mod.table_O_S_2f_2, '_barum_patched', False):
        _os2_orig_setur = _os2mod.table_O_S_2f_2.setUnicodeRanges
        _os2mod.table_O_S_2f_2.setUnicodeRanges = lambda self, bits: _os2_orig_setur(self, {b for b in bits if 0 <= b <= 122})
        _os2mod.table_O_S_2f_2._barum_patched = True
except Exception:
    pass
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

_BRAIN_TBL=[('grp','출혈성 뇌혈관 (I60~62)',None,None,None),('row','뇌출혈','I60~62','hem',[1,1,1]),('grp','허혈성 뇌혈관 (I63~66)',None,None,None),('row','뇌졸증·뇌경색','I63·65·66','infarct',[1,1,1]),('grp','기타 뇌혈관 (I64·67~69)',None,None,None),('row','기타 뇌혈관질환','I64·67·68·69','other',[1,1,1]),('grp','순환계 확장·선천',None,None,None),('row','뇌동맥류·정맥류','I71·72','aneur',[0,1,0]),('row','선천 뇌혈관기형','Q28.0~28.3','congen',[0,0,1]),('row','외상성 뇌출혈','S06','trauma',[0,0,1])]
_HEART_TBL=[('grp','허혈성 심장질환 (I20~25)',None,None,None),('row','급성심근경색','I21~23','ami',[1,1,1,1]),('row','협심증','I20','angina',[1,1,1,1]),('row','기타·만성 허혈','I24·25','chronic',[1,1,1,1]),('grp','심장특정 (판막·염증·부정맥·심근)',None,None,None),('row','심장판막','I05·I34~37','valve',[0,1,1,1]),('row','심근·심내막 염증','I30~33·I40','inflam',[0,1,1,1]),('row','부정맥','I49','arrhy',[0,1,1,1]),('row','심부전','I50','hf',[0,1,1,1]),('row','심근병증','I42~45','cardiomyo',[0,1,1,1]),('grp','순환계 확장 (2대+동맥류·정맥류 등)',None,None,None),('row','대동맥류·죽상경화','I70·71','aorta',[0,0,1,1]),('row','동맥류·정맥류 등','[확인]','aneur2',[0,0,1,1]),('row','선천 심장기형','Q20~25','congenh',[0,0,0,1])]
def _scv_build(tbl, headers, held, amounts=None):
    held=set(held or []); amounts=amounts or {}; ncol=len(headers)+1
    th=''.join('<th>'+h+'</th>' for h in headers)
    out=['<table class="scvt"><tr><th class="dl">질병 (코드)</th>'+th+'</tr>']
    for kind,label,code,key,cells in tbl:
        if kind=='grp':
            out.append('<tr class="grp"><td colspan="'+str(ncol)+'">'+_html.escape(label)+'</td></tr>'); continue
        own = key in held
        amt = amounts.get(key)
        rowcls=' class="own"' if own else ''
        def cell(c): return '<td><span class="on">●</span></td>' if c else '<td><span class="off">○</span></td>'
        if own:
            _amt=(' <span class="amt">'+_html.escape(amt)+'</span>') if amt else ''
            first=('<span class="on">●</span>' if cells[0] else '<span class="off">○</span>')+' <span class="chip">보유</span>'+_amt
            tds='<td>'+first+'</td>'+''.join(cell(c) for c in cells[1:])
        else:
            tds=''.join(cell(c) for c in cells)
        out.append('<tr'+rowcls+'><td class="dl"><b>'+_html.escape(label)+'</b><br><span class="cd">'+_html.escape(code)+'</span></td>'+tds+'</tr>')
    out.append('</table>'); return ''.join(out)


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
            v=c.get('value')
            if not v or v=='미가입':
                return '<b style="color:#C0444C">✘ 미가입</b>'
            return f'<b style="color:#1F7A4D">✔ 가입 · {_html.escape(str(v))}</b>'
    return '<b style="color:#C0444C">✘ 미가입</b>'

def _cov_val(rep, cat, *names):
    for c in rep.get('coverage',[]):
        nm=c.get('name','')
        if cat in nm:
            for it in c.get('items',[]):
                if any(n in it.get('t','') for n in names):
                    return it.get('v')
    return None

def _ws_amt(rep, kind):
    if kind=='암':
        v=_cov_val(rep,'암','일반암')
    else:
        v=_cov_val(rep,'뇌혈관','뇌졸증진단','뇌혈관진단') or _cov_val(rep,'심장','급성심근','허혈성')
    if v:
        return f'<b style="color:#1F7A4D">✔ 가입 · {_html.escape(str(v))}</b>'
    return '<b style="color:#C0444C">✘ 미가입</b>'

def _wc_status(rep, lookup):
    """(status, value) — status: 'on'(가입)/'off'(미가입)"""
    if lookup.startswith('diag:'):
        kind=lookup[5:]
        v=_cov_val(rep,'암','일반암') if kind=='암' else (_cov_val(rep,'뇌혈관','뇌졸증진단','뇌혈관진단') or _cov_val(rep,'심장','급성심근','허혈성'))
        return ('on', v) if v else ('off','')
    for c in rep.get('chiryo',[]):
        if c.get('name')==lookup:
            val=c.get('value')
            if not val or val=='미가입': return ('off','')
            return ('on', val)
    return ('off','')

def _wcard(rep, title, desc, lookup, mode):
    # mode: 'diag'(초록·✔·가입/미가입) / 'main'(빨강·✘·미가입/가입) / 'na'(회색·상태). 박스=빈칸(워크시트).
    if mode=='na':
        return (f'<div class="wcard plain"><div class="wct">{_html.escape(title)}</div>'
                f'<div class="wcd">{_html.escape(desc)}</div>'
                f'<div class="wcf"><span class="wchip n">상태</span><span class="wbox"></span><span class="wunit">만원</span></div></div>')
    st,_=_wc_status(rep, lookup)
    variant,check=('green','✔') if mode=='diag' else ('red','✘')
    chip=('<span class="wchip g">가입</span>' if st=='on' else '<span class="wchip r">미가입</span>')
    return (f'<div class="wcard {variant}"><div class="wct">{check} {_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>'
            f'<div class="wcf">{chip}<span class="wbox"></span><span class="wunit">만원</span></div></div>')

def _wcard_sj(rep, title, desc, lookup):
    # 산정특례 = 나란히 2칸, 회색 '상태' 칩, 빈 박스 (레퍼런스 그대로)
    return (f'<div class="wcard plain sj"><div class="wct">{_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>'
            f'<div class="wcf"><span class="wchip n">상태</span><span class="wbox"></span><span class="wunit">만원</span></div></div>')

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
        _rate=ci.get('rate',0); _rem=max(0,100-_rate)
        _cards=''.join(f'<div class="cicard"><div class="t">{_html.escape(i["t"])}</div><div class="v">{_html.escape(i["v"])}</div><div class="s">진단 시 선지급</div></div>' for i in ci['items'])
        _pre=ci.get('items',[{}])[0].get('v','')
        ci_html=(f'<div class="sect">CI 선지급 분석 <span>CRITICAL ILLNESS · PRE-PAYMENT</span></div>'
                 f'<div class="ci2">'
                 f'<div class="ci2-hd">선지급 <b>{_rate}%</b> 형</div>'
                 f'<div class="ci2-desc">{_html.escape(ci.get("product",""))} — CI 주계약 사망보험금 <b>{_html.escape(ci["samang"])}</b>. 암·뇌졸증·급성심근 등 중대질병 진단 시 <b>{_rate}% 선지급({_pre})</b>, 잔여 {_rem}%({_html.escape(ci["residual"])})는 사망 시 지급된다.</div>'
                 f'<div class="ci2-split"><div class="a"><div class="lb">선지급 {_rate}%</div><div class="amt">{_pre} · 진단 시</div></div>'
                 f'<div class="b"><div class="lb">잔여 {_rem}%</div><div class="amt">{_html.escape(ci["residual"])} · 사망 시</div></div></div>'
                 f'<div class="ci2-cards">{_cards}</div>'
                 f'<div class="ci2-note">■ <b>CI 선지급형이란</b> — 중대질병 진단 시 사망보험금의 일부를 미리(선지급) 받아 치료비로 쓰고, 사망 시 잔여분이 지급되는 구조다. 선지급률은 상품별로 50% 또는 80%만 존재한다. 진단 후에도 잔여 사망보장 {_html.escape(ci["residual"])}이 유지된다.</div>'
                 f'</div>')
    else:
        ci_html=''

    # ── CI 3상태: ci(선지급) / check(판정 [확인] 회색지대) / none(비CI) ──
    noci=rep.get('noci',{'present':False})
    _cist=ci.get('status') or ('ci' if ci.get('present') else 'none')
    if _cist=='check':
        noci_html=('<div class="sect">핵심 보장 분석 <span>CI 판정 보류 · CONFIRM</span></div>'
                   '<div class="noci-wrap">'
                   '<div class="noci-big" style="color:#9C7C32">CI 판정 [확인]</div>'
                   '<div class="noci-one" style="border-left-color:#9C7C32;background:#FBF6E6">'
                   '<div class="t" style="color:#8A6D1E">■ 판정 보류 — 수기 확인 필요</div>'
                   '<div class="d">상품명에 <b>CI · GI · 리빙케어</b>가 있으나 <b>중대한OO 담보 금액이 추출되지 않았습니다</b>(선지급률·기준액 미상 등). CI 여부와 선지급률(50%/80%)을 <b>약관으로 직접 확인</b>한 뒤 3장을 확정하세요. 추측 금지 · [확인].</div>'
                   '</div></div>')
    elif _cist=='none':
        noci_html=('<div class="sect">핵심 보장 분석 <span>GENERAL INSURANCE · NON-CI</span></div>'
                   '<div class="noci-wrap">'
                   '<div class="noci-big">CI보험 아님</div>'
                   '<div class="noci-one">'
                   '<div class="t">■ 일반 보험기준</div>'
                   '<div class="d">이 고객은 <b>CI(중대질병 선지급형) 계약이 아닙니다.</b> 암·뇌·심 진단비는 진단 즉시 <b>정액 100%</b> 지급되는 일반형이며, CI형과 달리 <b>사망보험금 차감·선지급·잔여</b> 개념이 없습니다. 진단금 전액이 치료비로 쓰이고, 사망보장은 별도로 유지됩니다.</div>'
                   '</div>'
                   '</div>')
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
.scv2 {{ display:flex; gap:5mm; }}
.scvcol {{ flex:1; }}
.scvhd {{ font-size:11pt; font-weight:800; padding-bottom:1.5mm; border-bottom:2pt solid {GOLD}; margin-bottom:2mm; }}
.scvhd.brain {{ color:#1F5FA8; }} .scvhd.heart {{ color:{GAP}; }}
.scvt {{ width:100%; border-collapse:collapse; font-size:7pt; }}
.scvt th {{ background:{NAVY}; color:#fff; padding:1.2mm 1mm; font-size:6.4pt; text-align:center; font-weight:700; }}
.scvt th.dl {{ text-align:left; }}
.scvt th:last-child {{ background:{GOLDD}; }}
.scvt td {{ border-top:0.4pt solid {LINE}; padding:1.1mm 1.4mm; text-align:center; }}
.scvt td.dl {{ text-align:left; }}
.scvt td.dl .cd {{ font-size:5.8pt; color:{MUT}; }}
.scvt tr.grp td {{ background:#EEF1F6; font-size:6.4pt; font-weight:800; color:{NAVY}; text-align:left; padding:1mm 1.4mm; }}
.scvt tr.own td {{ background:#FBF6E6; }}
.scvt td:last-child {{ background:#F2F6F1; }}
.scvt tr.own td:last-child {{ background:#F0EFD8; }}
.on {{ color:#1F7A4D; font-weight:800; }} .off {{ color:#B9C2CE; }}
.hold {{ color:{GOLDD}; font-weight:700; font-size:6pt; }}
.chip {{ background:{GOLD}; color:#fff; font-size:5.4pt; font-weight:800; padding:0.2mm 1mm; border-radius:2mm; }}
.scvt .amt {{ color:{NAVY}; font-weight:800; font-size:6.6pt; margin-left:0.6mm; }}
.scvleg {{ font-size:6.2pt; color:{MUT}; margin:1.5mm 0; }} .own2 {{ color:{GOLDD}; font-weight:700; }}
.scvnote {{ font-size:6.4pt; line-height:1.5; color:{INK}; background:#F6F8FB; border-left:2.2pt solid {NAVY}; padding:2mm 2.5mm; border-radius:1.4mm; }}
.scvnote b {{ color:{NAVY}; }} .scvnote b.r {{ color:{GAP}; }}
.scvbot {{ font-size:6.6pt; line-height:1.55; color:{INK}; background:#FBF7EE; border:0.6pt solid {GOLD}; border-radius:1.6mm; padding:2.4mm 3mm; margin-top:3mm; }}
.scvbot .h {{ font-size:8pt; font-weight:800; color:{GOLDD}; margin-bottom:1mm; }}
.scvbot b {{ color:{NAVY}; }}

.ci2 {{ border:0.6pt solid {LINE}; border-radius:2mm; padding:4mm; }}
.ci2-hd {{ display:inline-block; background:{NAVY}; color:#fff; font-size:10pt; font-weight:800; padding:1.4mm 4mm; border-radius:4mm; }}
.ci2-hd b {{ color:{GOLDL}; }}
.ci2-desc {{ font-size:8.6pt; line-height:1.5; color:{INK}; margin:3mm 0; }}
.ci2-desc b {{ color:{NAVY}; }}
.ci2-split {{ display:flex; gap:3mm; margin-bottom:3.5mm; }}
.ci2-split .a, .ci2-split .b {{ flex:1; border-radius:1.6mm; padding:2.5mm 3mm; }}
.ci2-split .a {{ background:{NAVY}; color:#fff; }} .ci2-split .b {{ background:#EEF1F6; color:{NAVY}; }}
.ci2-split .lb {{ font-size:8pt; font-weight:700; opacity:0.85; }}
.ci2-split .amt {{ font-size:11pt; font-weight:800; margin-top:1mm; }}
.ci2-split .a .lb {{ color:{GOLDL}; opacity:1; }}
.ci2-cards {{ display:flex; gap:3mm; margin-bottom:3.5mm; }}
.cicard {{ flex:1; border:0.6pt solid {GOLD}; border-radius:1.6mm; padding:2.5mm; text-align:center; background:#FBF7EE; }}
.cicard .t {{ font-size:8.4pt; font-weight:800; color:{NAVY}; }}
.cicard .v {{ font-size:13pt; font-weight:800; color:{GOLDD}; margin:1mm 0; }}
.cicard .s {{ font-size:6.6pt; color:{MUT}; }}
.ci2-note {{ font-size:7.6pt; line-height:1.5; color:{MUT}; background:#F6F8FB; border-left:2.4pt solid {NAVY}; padding:2.4mm 3mm; border-radius:1.4mm; }}
.ci2-note b {{ color:{NAVY}; }}
.noci-wrap {{ margin-top:2mm; }}
.noci-big {{ font-size:30pt; font-weight:800; color:{GAP}; text-align:center; padding:10mm 0 7mm; letter-spacing:3px; }}
.noci-one {{ border:1pt solid {LINE}; border-left:4pt solid {NAVY}; border-radius:2mm; background:#F6F8FB; padding:5mm 6mm; }}
.noci-one .t {{ font-size:12.5pt; font-weight:800; color:{NAVY}; margin-bottom:2.5mm; }}
.noci-one .d {{ font-size:9.8pt; line-height:1.65; color:{INK}; }}
.noci-one .d b {{ color:{NAVY}; }}

.cvbar {{ height:9mm; background:{NAVY}; }}
.cvbody {{ padding:16mm 16mm 12mm; height:271mm; display:flex; flex-direction:column; }}
.cvbrand {{ font-size:14pt; font-weight:800; color:{NAVY}; letter-spacing:5px; }}
.cvbrand .ln {{ width:24mm; height:1.6mm; background:{GOLD}; margin-top:2.5mm; }}
.cvtitle {{ font-size:31pt; font-weight:800; color:{NAVY}; line-height:1.28; margin-top:44mm; }}
.cvtitle .g {{ color:{GOLD}; }}
.cvsub {{ font-size:11.5pt; font-weight:700; color:{NAVY}; margin-top:5mm; }}
.cvhr {{ border-top:1.6pt solid {GOLD}; margin:6mm 0 5mm; }}
.cvstats {{ display:flex; gap:6mm; }}
.cvst {{ flex:1; }}
.cvst .k {{ font-size:9pt; color:{MUT}; font-weight:700; }}
.cvst .v {{ font-size:17pt; font-weight:800; color:{NAVY}; margin-top:2mm; }}
.cvst .v small {{ font-size:9pt; color:{MUT}; font-weight:600; }}
.cvspacer {{ flex:1; }}
.cvhr2 {{ border-top:0.5pt solid {LINE}; margin-bottom:4mm; }}
.cvfoot {{ text-align:center; font-size:9.5pt; font-weight:700; color:{NAVY}; }}

.jc {{ display:flex; gap:5mm; }}
.jcol {{ flex:1; }}
.jhd {{ font-size:11pt; font-weight:800; color:{NAVY}; border-bottom:2pt solid {GOLD}; padding-bottom:1mm; margin-bottom:2mm; }}
.jhd .sub {{ float:right; font-size:9pt; color:{MUT}; font-weight:600; margin-top:2mm; }}
.jstages {{ display:flex; gap:1.4mm; margin-bottom:2.5mm; }}
.jstage {{ flex:1; border:0.5pt solid {LINE}; border-radius:1.4mm; padding:1.2mm; text-align:center; font-size:5.8pt; line-height:1.25; color:{INK}; }}
.jstage b {{ display:block; font-size:6.8pt; color:{NAVY}; }}
.jstage .dt {{ color:{GOLD}; font-weight:700; font-size:5.6pt; }}
.jtwo {{ display:flex; gap:1.6mm; margin-bottom:2mm; }}
.jt {{ width:100%; border-collapse:collapse; font-size:8pt; table-layout:fixed; }}
.jt td, .jt th {{ overflow:hidden; }}
.jtsm {{ font-size:6.6pt; }}
.jtsm th {{ font-size:6.6pt; }}
.jt4 col, .jt4 {{ }}
.jt4 th:nth-child(1), .jt4 td:nth-child(1), .jt4 th:nth-child(3), .jt4 td:nth-child(3) {{ width:29%; }}
.jt4 th:nth-child(2), .jt4 td:nth-child(2), .jt4 th:nth-child(4), .jt4 td:nth-child(4) {{ width:21%; }}
.jt th {{ background:{NAVY}; color:#fff; padding:1.5mm 1.6mm; font-weight:700; font-size:8pt; }}
.jt td {{ border:0.4pt solid {LINE}; padding:2.3mm 1.6mm; }}
.jt td.p {{ color:{GAP}; font-weight:800; text-align:right; }}
.jclause {{ background:#F6F8FB; border:0.5pt solid {LINE}; border-left:2.4pt solid {NAVY}; border-radius:1.2mm; padding:1.4mm 2mm; margin-bottom:1.8mm; font-size:6.2pt; line-height:1.4; }}
.jclause b {{ color:{NAVY}; font-size:6.6pt; }}
.jclause li {{ margin:0.5mm 0 0.5mm 3mm; }}
.jwarn {{ background:#FBECEC; border:0.5pt solid #E4B4B4; border-radius:1.4mm; padding:2.4mm 2.2mm; margin-bottom:2.4mm; font-size:8pt; line-height:1.55; }}
.jwarn b {{ color:{GAP}; }}
.jwarn .h {{ font-size:9pt; font-weight:800; color:{GAP}; }}
.jtalk {{ background:#EAF2FB; border-left:2.4pt solid {BLUE}; border-radius:1.4mm; padding:2.4mm 2.2mm; font-size:8pt; line-height:1.6; color:{NAVY}; font-style:italic; }}
.jtalk .h {{ font-style:normal; font-weight:800; color:{BLUE}; font-size:9pt; }}
.p5own {{ display:flex; gap:2.5mm; margin:2.5mm 0 3.5mm; }}
.p5it {{ flex:1; border:0.7pt solid {LINE}; border-radius:1.8mm; padding:2.4mm 2mm; font-size:9.5pt; font-weight:800; text-align:center; }}
.p5it.on {{ background:#FBF1D8; border-color:{GOLD}; color:{NAVY}; }}
.p5it.off {{ color:{GAP}; background:#FBFBFC; }}
.p5it b {{ display:block; font-size:7pt; font-weight:700; color:{MUT}; margin-bottom:1mm; }}
.p5tbl {{ width:100%; border-collapse:collapse; margin:2.5mm 0 3.5mm; table-layout:fixed; }}
.p5tbl th {{ background:{NAVY}; color:#fff; font-size:8.4pt; font-weight:700; padding:2mm 1mm; text-align:center; border:0.5pt solid {NAVY}; }}
.p5tbl td {{ font-size:12pt; font-weight:800; text-align:center; padding:2.8mm 1mm; border:0.5pt solid {LINE}; }}
.p5tbl td.on {{ color:{NAVY}; background:#FBF6E6; }}
.p5tbl td.off {{ color:{GAP}; background:#FBFBFC; }}
.jstep {{ display:flex; gap:1.6mm; align-items:flex-start; border:0.5pt solid {LINE}; border-radius:1.4mm; padding:1.2mm 1.6mm; margin-bottom:1.4mm; }}
.jstep .n {{ flex:0 0 auto; width:4.2mm; height:4.2mm; border-radius:50%; background:{GOLD}; color:#fff; font-size:6.5pt; font-weight:800; text-align:center; line-height:4.2mm; }}
.jstep .b {{ flex:1; font-size:6.2pt; line-height:1.35; }}
.jstep .b .t {{ font-weight:800; color:{NAVY}; font-size:6.8pt; }}
.jstep .b .d {{ float:right; color:{GOLD}; font-weight:700; font-size:6pt; }}
.jstep .b .g {{ color:{GOOD}; font-weight:700; }}
.jsec {{ font-size:10.5pt; font-weight:800; color:{NAVY}; margin:1.8mm 0 1.1mm; border-bottom:0.6pt solid {LINE}; padding-bottom:0.8mm; }}
.jsec .jn {{ display:inline-block; width:4.4mm; height:4.4mm; border-radius:50%; background:{GOLD}; color:#fff; font-size:7pt; text-align:center; line-height:4.4mm; margin-right:1.2mm; }}
.jsec .jdt {{ float:right; color:{GAP}; font-weight:800; font-size:8.4pt; margin-top:1mm; }}
.jul {{ margin:0 0 3.2mm 0; padding:0; list-style:none; font-size:8.6pt; line-height:1.85; color:{INK}; }}
.jul li {{ margin:1mm 0 1mm 2.6mm; text-indent:-2.2mm; }}
.jul li:before {{ content:"· "; color:{NAVY}; font-weight:800; }}
.jnote2 {{ font-size:6.8pt; color:{MUT}; text-align:center; margin-top:1mm; line-height:1.35; }}
.wsname {{ text-align:center; margin:3mm 0 4mm; }}
.wsname .nm {{ display:inline-block; font-size:15pt; font-weight:800; color:#fff; background:linear-gradient(135deg,{NAVY},{NAVY2}); padding:2mm 10mm; border-radius:6mm; border:1.4pt solid {GOLD}; }}
.wsname .q {{ display:block; font-size:8pt; color:{MUT}; margin-top:1.5mm; }}
.ws2 {{ display:flex; gap:4mm; }}
.ws2 .wcard {{ flex:1 1 0; min-width:0; }}
.wscol {{ flex:1; }}
.wscap {{ font-size:11.5pt; font-weight:800; padding:1.6mm 3mm; border-radius:1.4mm 1.4mm 0 0; color:#fff; }}
.wscap.c {{ background:#B9540B; }} .wscap.h {{ background:#1F5FA8; }}
.wsr {{ border:0.5pt solid {LINE}; border-top:none; padding:2mm 3mm; font-size:8.8pt; line-height:1.4; }}
.wsr .t {{ font-weight:800; color:{NAVY}; font-size:9.6pt; }}
.wsr .t.on {{ color:{GOOD}; }} .wsr .t.off {{ color:{GAP}; }}
.wsr .d {{ color:{MUT}; font-size:8pt; }}
.wsr .amt {{ float:right; font-size:9pt; font-weight:800; color:{INK}; }}
.wslack {{ background:#FBF6E6; border:0.5pt solid #E6D08A; border-radius:1.4mm; padding:2.4mm 3mm; margin:3mm 0 2mm; font-size:9pt; }}
.wslack .h {{ font-weight:800; color:{GOLDD}; font-size:8pt; }}
.wslack .bl {{ display:inline-block; border-bottom:0.6pt solid {INK}; width:48mm; height:3mm; }}
.wslack .lackbox {{ height:18mm; }}
.wstalk {{ background:#EAF2FB; border-left:2.6pt solid {BLUE}; border-radius:1.4mm; padding:2mm 3mm; font-size:7.6pt; line-height:1.5; color:{NAVY}; font-style:italic; }}
.wstalk .h {{ font-style:normal; font-weight:800; color:{BLUE}; }}
.ws3 {{ display:flex; gap:2.5mm; align-items:stretch; }}
.wscol.wsmain {{ display:flex; flex-direction:column; }}
.wsmain .wscap {{ flex:0 0 auto; }}
.wsmain .wsr {{ min-height:39mm; display:flex; flex-direction:column; justify-content:center; }}
.wsr .box {{ display:block; margin-top:1.6mm; border:0.6pt solid {LINE}; border-radius:1mm; height:6mm; background:#FCFDFE; }}
.wsmid {{ flex:0 0 30mm; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; }}
.wsmid .lb {{ font-size:6.5pt; color:{MUT}; letter-spacing:1px; font-weight:700; }}
.wsmid .nm {{ font-size:32pt; font-weight:800; color:{NAVY}; line-height:1.02; letter-spacing:0; margin:2mm 0 0; }}
.wsmid .nmsub {{ font-size:12pt; font-weight:800; color:{GOLDD}; margin:1.5mm 0 2mm; }}
.wsmid .cnt {{ font-size:13pt; font-weight:800; color:{NAVY}; }}
.wsmid .cnt small {{ display:block; font-size:6.2pt; color:{MUT}; font-weight:600; }}
.wssj {{ margin-top:3mm; border:1pt solid #D8B65A; border-radius:2mm; background:#FDFAF0; padding:3mm 3.5mm 3.5mm; }}
.wcard {{ border:0.9pt solid {LINE}; border-radius:2mm; padding:2.8mm 3mm; margin-bottom:2.4mm; min-height:26mm; display:flex; flex-direction:column; }}
.wcard.green {{ border-color:#3E9A63; background:#F4FBF7; }}
.wcard.red {{ border-color:#DE9A9A; background:#FDF4F3; }}
.wcard.plain {{ border-color:{LINE}; background:#FCFDFE; }}
.wcard .wct {{ font-size:11pt; font-weight:800; color:{NAVY}; }}
.wcard.green .wct {{ color:#1F7A4D; }}
.wcard.red .wct {{ color:#C0444C; }}
.wcard .wcd {{ font-size:8.6pt; color:{MUT}; margin:1.4mm 0 auto; }}
.wcard .wcf {{ display:flex; align-items:center; gap:2mm; margin-top:2.2mm; }}
.wchip {{ font-size:9pt; font-weight:800; color:#fff; padding:1.1mm 3.2mm; border-radius:1.5mm; white-space:nowrap; }}
.wchip.g {{ background:#3E9A63; }}
.wchip.r {{ background:#C0444C; }}
.wchip.n {{ background:#9AA5B4; }}
.wbox {{ flex:1; border:0.7pt solid {LINE}; border-radius:1mm; height:7.5mm; line-height:7.5mm; padding:0 2.5mm; font-size:10.5pt; font-weight:800; color:{NAVY}; background:#fff; text-align:right; }}
.wunit {{ font-size:8.6pt; color:{MUT}; }}
.wcard.sj {{ min-height:22mm; margin-bottom:0; background:#fff; }}
.wcard.sj .wct {{ font-size:11pt; }}
.wssj .cap {{ font-size:9.4pt; font-weight:800; color:{GOLDD}; padding:0 0 2mm; }}



.hcnote {{ background:{NAVY}; color:#EAF0F8; font-size:7.4pt; line-height:1.45; padding:2.4mm 3.5mm; border-radius:2mm; margin-bottom:3mm; }}
.hcnote b {{ color:{GOLDL}; }}
.chnote {{ font-size:7pt; color:{MUT}; line-height:1.4; margin-top:1.5mm; }}
.chnote b {{ color:{NAVY}; }}
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
    # ★badge-5 담보별: 상단 박스 제거 → 뇌졸증·뇌출혈·급성심근경색 보유금액을 질병코드 표 행 안에 직접 기재(지점장 2026.07.07)
    p5box=''
    _amt_brain={}; _amt_heart={}
    _AMTKEY={'뇌출혈진단비':('b','hem'),'뇌졸증진단비':('b','infarct'),'뇌졸중진단비':('b','infarct'),'급성심근경색':('h','ami')}
    for _it in rep.get('p5_own',[]):
        _v=_it.get('v')
        if not _v or _v=='미가입': continue
        _m=_AMTKEY.get(_it.get('t',''))
        if not _m: continue
        (_amt_brain if _m[0]=='b' else _amt_heart)[_m[1]]=_v
    scv_brain=_scv_build(_BRAIN_TBL,['뇌혈관<br>진단비','순환계','산정<br>특례'],rep.get('scope_brain'),_amt_brain)
    scv_heart=_scv_build(_HEART_TBL,['허혈성<br>진단비','심장<br>(특정)','순환계','산정<br>특례'],rep.get('scope_heart'),_amt_heart)
    doc=f'''<!DOCTYPE html><html><head><meta charset="utf-8"><style>{css}</style></head><body>
<!-- P0: 표지 (최종본 260707 스펙) -->
<div class="pg">
 <div class="cvbar"></div>
 <div class="cvbody">
  <div class="cvbrand">MAKEONE<div class="ln"></div></div>
  <div class="cvtitle">보장 진단서<br><span class="g">{cust}</span> 고객님</div>
  <div class="cvsub">{rep.get('meta','')}</div>
  <div class="cvhr"></div>
  <div class="cvstats">
   <div class="cvst"><div class="k">보유 계약</div><div class="v">{rep.get('n_contract',0)} <small>건</small></div></div>
   <div class="cvst"><div class="k">월 납입보험료</div><div class="v">{rep.get('premium',0):,} <small>원</small></div></div>
   <div class="cvst"><div class="k">갱신 / 비갱신</div><div class="v">{rep.get('renew',0)} <small>/ {rep.get('nonrenew',0)}</small></div></div>
  </div>
  <div class="cvspacer"></div>
  <div class="cvhr2"></div>
  <div class="cvfoot">MAKEONE · 보장분석 자동화 리포트</div>
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
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 1 / 9</span></div>
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
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 2 / 9</span></div>
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
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 3 / 9</span></div>
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
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 4 / 9</span></div>
</div>
<!-- P5: 담보별 보장범위 (최종본 260707 스펙) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>5</b>담보별 보장범위</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">담보별 보장범위 — 질병코드 커버 <span>DISEASE-CODE COVERAGE · 각 축=개별 담보·각각 보상 · 산정특례·순환계·외상성뇌출혈=단독</span></div>
  {p5box}
  <div class="scv2">
   <div class="scvcol">
    <div class="scvhd brain">뇌 — 질병코드별 커버</div>
    {scv_brain}
    <div class="scvleg"><span class="on">●</span> 보장 &nbsp;<span class="off">○</span> 미보장 &nbsp;<span class="hold">[확인]</span> 산정특례 HOLD &nbsp;<span class="own2">노란행=보유</span></div>
    <div class="scvnote">· {cust} 보유 = <b>뇌출혈진단비(I60~62)</b> → <b class="r">뇌경색(I63)부터 공백</b>, 순환계·2대주요치료비로 확장 필요<br>· <b>외상성 뇌출혈(S06)</b> = 뇌혈관진단비 <b class="r">미보장</b>(산정특례 축only)<br>· <b>산정특례</b> = 진단 기반 별개 담보축. 대상 코드범위 = 뇌혈관질환 전체 I60~69 + Q28 선천 + S06(개별 담보·각각 보상). 지급조건은 회사·약관별 [확인]</div>
   </div>
   <div class="scvcol">
    <div class="scvhd heart">심장 — 질병코드별 커버</div>
    {scv_heart}
    <div class="scvleg"><span class="on">●</span> 보장 &nbsp;<span class="off">○</span> 미보장 &nbsp;<span class="hold">[확인]</span> 산정특례 HOLD &nbsp;<span class="own2">노란행=보유</span></div>
    <div class="scvnote">· {cust} 보유 = <b>허혈성진단비 + 부정맥(I49)</b> → 허혈성(I20~25)·부정맥 커버, <b class="r">판막·염증·심부전은 심장특정/순환계 영역</b><br>· <b>빈맥(I47·48)</b> = 마스터 무행·전 묶음 <b class="r">제외</b> → 본 표 미기재<br>· <b>산정특례</b> = 진단 기반 별개 담보축. 대상 코드범위 = 심혈관질환 전체 I20~50 + 판막(개별 담보·각각 보상). 지급조건 [확인]<br>· <b>순환계</b> = 가장 넓은 범위(2대 + 동맥류·정맥류 등 순환기 전반). 세부 대상코드 회사·약관별 상이 [확인]</div>
   </div>
  </div>
  <div class="scvbot"><div class="h">산정특례 기준 (진단 기반 · 별개 담보축)</div>· 산정특례 = 위 4범위(허혈성·2대·순환계)와 <b>축이 다른 별개 담보</b> — 마스터 '산정특례심장'·'산정특례뇌혈관' 전용행에서 진단코드 기반으로 지급. &nbsp;· 외상성 뇌출혈(S06) = 뇌혈관진단비 미보장 → <b>산정특례 축only</b>로만 커버(고정사실). &nbsp;· 대상 코드범위(확정): <b>[뇌]</b> I60~69 전체 + I67.0·1·5·6 + Q28 선천 + S06 / <b>[심]</b> I20~25·I30~41·I42~45·I46·I47~50 + 판막. <b>각각 개별 담보로 각각 보상.</b> 지급조건·기간(30일·5% 등)만 회사·약관별 [확인].</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 5 / 9</span></div>
</div>

<!-- P6: 주요치료비 변천사 (juyo_a4_v7 상세본) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>6</b>주요치료비 변천사</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">3대 주요치료비 담보의 변천사 <span>①비례형(구간) → ②정액형 → ③비급여 → 생활비 · 정액형 가입금액 100만원부터</span></div>
  <div class="jc">
   <div class="jcol">
    <div class="jhd">🎗️ 암 주요치료비 <span class="sub">4세대 진화</span></div>
    <div class="jsec"><span class="jn">1</span>비례형(구간형) — 치료비 발생 기준<span class="jdt">~24.11 단종</span></div>
    <table class="jt jt4">
     <tr><th>치료비 발생</th><th>지급</th><th>치료비 발생</th><th>지급</th></tr>
     <tr><td>300~500만</td><td class="p">300만</td><td>4,000~5,000만</td><td class="p">4,000만</td></tr>
     <tr><td>500~700만</td><td class="p">500만</td><td>5,000~6,000만</td><td class="p">5,000만</td></tr>
     <tr><td>700~1,000만</td><td class="p">700만</td><td>6,000~7,000만</td><td class="p">6,000만</td></tr>
     <tr><td>1,000~2,000만</td><td class="p">1,000만</td><td>7,000~8,000만</td><td class="p">7,000만</td></tr>
     <tr><td>2,000~3,000만</td><td class="p">2,000만</td><td>8,000~9,000만</td><td class="p">8,000만</td></tr>
     <tr><td>3,000~4,000만</td><td class="p">3,000만</td><td>9,000~1억</td><td class="p">9,000~1억</td></tr>
    </table>
    <ul class="jul"><li>치료비 지출액이 구간 하한 도달 시 그 구간 정액 지급(지출 없으면 미지급)</li><li>최소 하한 미달분 미지급(삼성 1,000만 시작 → 농협 300만까지)</li><li>실손과 중복 수령 가능 · 2024.11 단종(재가입 불가) — 보유자 해지 금지</li></ul>
    <div class="jsec"><span class="jn">2</span>정액형<span class="jdt">24.11~ 현재 신규</span></div>
    <ul class="jul"><li>치료 사실만으로 약정금 정액 지급(실비 무관, 고정금액)</li><li>암수술·항암방사선·항암약물 각 보장, 특약 시 입원·검사 추가</li><li>면책 90일, 감액기간 일부(~50%) · 가입금액 100만원부터</li></ul>
    <div class="jsec"><span class="jn">3</span>비급여(하이클래스)<span class="jdt">25.03~</span></div>
    <ul class="jul"><li>급여 전액본인부담금 + 비급여 치료비 정액 보장</li><li>산정특례가 안 주는 비급여(양성자·중입자·표적·면역) 커버</li><li>면책 90일, 연 1회 · 2천만×10년=2억(삼성 최초) · 25.08~ 생활비형(치료+소득보상) 추가</li></ul>
    <div class="jsec"><span class="jn">4</span>설명 — 왜 필요한가</div>
    <div class="jwarn"><div class="h">⚠ 실손으론 부족 — 5세대 기준</div>· 비중증 비급여: 자기부담 50%·통원 회당 5만원·한도 <b>1,000만</b>(4세대 5천만→축소)<br>· 중증(암·산정특례): 자기부담 30%·연 5천만이나 표적·면역 고가 약값엔 부족<br>· 외래(통원) 주사 반복 → <b>주요치료비·하이클래스가 메움</b></div>
    <div class="jtalk"><span class="h">🗣 상담 화법</span> &nbsp;"표적항암 한 달 수백만 원이 전액 본인부담이에요. 5세대 실손은 비급여가 1천만으로 줄었습니다. 구간형은 이제 못 드는 담보라 유지가 답이고, 없으면 정액형·하이클래스로 채워야 합니다."</div>
   </div>
   <div class="jcol">
    <div class="jhd">❤️ 뇌·심장 주요치료비 <span class="sub">암과 동일 진화</span></div>
    <div class="jsec"><span class="jn">1</span>비례형(병원비형) — 급여 본인부담 기준<span class="jdt">~24.11 단종</span></div>
    <table class="jt jtsm jt4">
     <tr><th>급여 본인부담</th><th>지급</th><th>급여 본인부담</th><th>지급</th></tr>
     <tr><td>100만원 미만</td><td class="p">미지급(면책)</td><td>1,000~2,000만</td><td class="p">1,000~2,000만</td></tr>
     <tr><td>100~500만</td><td class="p">100~500만</td><td>2,000~3,000만</td><td class="p">2,000~3,000만</td></tr>
     <tr><td>500~1,000만</td><td class="p">500~1,000만</td><td>3,000만 이상</td><td class="p">3,000만(한도)</td></tr>
    </table>
    <ul class="jul"><li>지급 = 본인부담금 전액(비례) · 암(치료비 발생 하한 정액)과 축이 다름</li><li>예) 종합병원 2대질병 주요치료비 = 급여 본인부담 연 100만↑ → 3,000만 한도·10년 지급 · 구간액 회사별 상이[확인]</li></ul>
    <div class="jsec"><span class="jn">2</span>정액형(2대)<span class="jdt">24.11~ 메리츠 최초</span></div>
    <ul class="jul"><li>수술·혈전용해·중환자실 치료 사실만으로 정액(100만~) · 메리츠 최초, 연 2천만×5년=1억</li></ul>
    <div class="jsec"><span class="jn">3</span>확대(순환계) · 생활비<span class="jdt">25.01~ / 25.06~</span></div>
    <ul class="jul"><li>부정맥·심부전·동맥류 확대(52질환·연 8천만·100세) · 생활비형: 치료 하나만 받아도 연 1회</li></ul>
    <div class="jsec"><span class="jn">4</span>재발률 · 실제 수술비용<span class="jdt">심평원·건보공단</span></div>
    <div class="jwarn"><div class="h">🔁 한 번으로 안 끝난다</div>뇌경색 재발 1년 10%·<b>5년 20~30%</b> / 급성심근경색 <b>1년 30%·3년 50%</b> / 스텐트 5년 재협착 15%</div>
    <table class="jt"><tr><th>수술</th><th>건당 총진료비(평균)</th></tr>
     <tr><td>심장수술(개흉)</td><td class="p">2,832만원</td></tr>
     <tr><td>관상동맥우회술(CABG)</td><td class="p">2,690만원</td></tr>
     <tr><td>뇌동맥류 코일색전술</td><td class="p">1,100~1,600만원</td></tr>
     <tr><td>심장 스텐트(다발혈관)</td><td class="p">1,000~1,400만원</td></tr>
     <tr><td>뇌기저부수술</td><td class="p">1,475만원</td></tr></table>
    <div class="jnote2">총진료비 기준(심평원·건보공단 통계) · 재발 시 이 비용이 반복 — 주요치료비가 채운다</div>
   </div>
  </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 6 / 9</span></div>
</div>
<!-- P7: 상담 워크시트 (FINAL版·3단 중앙이름+산정특례) -->
<div class="pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 상담 워크시트</div>
  <div class="nm">지금 고객의 <b>3대 주요치료비</b>는?</div>
  <div class="pgn"><b>7</b>상담 워크시트</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="ws3">
   <div class="wscol wsmain">
    <div class="wscap c">🎗️ 암 주요치료비</div>
    {_wcard(rep,'암 진단비','걸렸을 때 일시금 (기본)','diag:암','diag')}
    {_wcard(rep,'암 주요치료비','수술·방사선·약물 정액 (100만~)','암주요치료비','main')}
    {_wcard(rep,'하이클래스 (비급여)','표적·면역·중입자 전액본인 커버','비급여주요치료비','na')}
    {_wcard(rep,'암 생활비','치료 중 소득보상','암생활비','na')}
   </div>
   <div class="wsmid">
    <div class="lb">OUR CLIENT</div>
    <div class="nm">{"<br>".join(list(cust)[:4])}</div>
    <div class="nmsub">고객님</div>
    <div class="cnt">2/8<small>3대 치료비 보유</small></div>
   </div>
   <div class="wscol wsmain">
    <div class="wscap h">❤️ 뇌·심장 주요치료비</div>
    {_wcard(rep,'뇌·심 진단비','뇌혈관·허혈성 일시금','diag:뇌심','diag')}
    {_wcard(rep,'2대 주요치료비','수술·혈전용해·중환자실 (100만~)','2대주요치료비','main')}
    {_wcard(rep,'순환계 주요치료비','부정맥·심부전 확대','순환계주요치료비','na')}
    {_wcard(rep,'순환계 생활비','치료 중 소득보상','순환계생활비','na')}
   </div>
  </div>
  <div class="wssj">
   <div class="cap">산정특례 — 뇌·심 각각 개별 담보 · 진단만으로 지급</div>
   <div class="ws2">
    {_wcard_sj(rep,'산정특례 (뇌)','뇌혈관질환 I60~69 전체 · Q28 · S06','산정특례(뇌혈관)')}
    {_wcard_sj(rep,'산정특례 (심장)','심혈관질환 I20~50 · 판막 전체','산정특례(심장)')}
   </div>
  </div>
  <div class="wslack"><span class="h">부족한 담보 → 보완 추천</span><div class="lackbox"></div></div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 7 / 9</span></div>
</div>
{heart_chart}
</body></html>'''
    HTML(string=doc).write_pdf(out)
    return True

# ── 정기철 샘플 데이터 (실제는 build_excel 결과에서 매핑) ──
if __name__=='__main__':
    import sys
    from coverage_benchmark import map_excel_to_report
    if len(sys.argv)<2:
        print('사용법: python report_weasy.py <엑셀> [고객명]'); sys.exit(1)
    _xl=sys.argv[1]; _cn=sys.argv[2] if len(sys.argv)>2 else '고객'
    _rep=map_excel_to_report(_xl, settings={'client':_cn,'branch':'온빛센터 바름지점','manager':'최은혜','title':'지점장','phone':''})
    build_report_pdf(_rep, f'보장설명지_{_cn}.pdf'); print('PDF 생성: 보장설명지_'+_cn+'.pdf')
