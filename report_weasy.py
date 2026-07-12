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

def _donut(pct, color, size=92, sw=12, over=False):
    r=(size-sw)/2; c=size/2; C=2*math.pi*r; off=C*(1-min(100,pct)/100)
    _lbl=f'{pct}%'
    _fs=21 if len(_lbl)<=3 else (19 if len(_lbl)==4 else 16)
    return f'''<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}">
<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="#E3E8EF" stroke-width="{sw}"/>
<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="{color}" stroke-width="{sw}" stroke-linecap="round"
 stroke-dasharray="{C:.1f}" stroke-dashoffset="{off:.1f}" transform="rotate(-90 {c} {c})"/>
<text x="{c}" y="{c}" text-anchor="middle" dominant-baseline="central" dy="0.02em" font-size="{_fs}" font-weight="800" fill="{NAVY}" font-family="NanumSquareRound">{_lbl}</text>
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
def _scv_build(tbl, headers, held, amounts=None, ci_amounts=None):
    held=set(held or []); amounts=amounts or {}; ci_amounts=ci_amounts or {}; ncol=len(headers)+1
    th=''.join('<th>'+h+'</th>' for h in headers)
    out=['<table class="scvt"><tr><th class="dl">질병 (코드)</th>'+th+'</tr>']
    for kind,label,code,key,cells in tbl:
        if kind=='grp':
            out.append('<tr class="grp"><td colspan="'+str(ncol)+'">'+_html.escape(label)+'</td></tr>'); continue
        own = key in held
        amt = amounts.get(key)
        rowcls=' class="own"' if own else ''
        def cell(c): return '<td><span class="on">●</span></td>' if c else '<td><span class="off">○</span></td>'
        _amtbox=' <span class="mb amtbox">'+(_html.escape(amt) if amt else '')+'</span>'
        # ★CI 계약일 때만 유동 노출: 'CI' 미니칩 + 진단금액
        _ci = ci_amounts.get(key)
        _cibox=(' <span class="cichip">CI</span><span class="mb amtbox cibox">'+_html.escape(_ci)+'</span>') if _ci else ''
        if own:
            first=('<span class="on">●</span>' if cells[0] else '<span class="off">○</span>')+' <span class="chip">보유</span>'+_amtbox+_cibox
        else:
            first=('<span class="on">●</span>' if cells[0] else '<span class="off">○</span>')+_amtbox+_cibox
        tds='<td>'+first+'</td>'+''.join(cell(c) for c in cells[1:])
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
                return '<b style="color:#C0444C">× 미가입</b>'
            return f'<b style="color:#1F7A4D">✓ 가입 · {_html.escape(str(v))}</b>'
    return '<b style="color:#C0444C">× 미가입</b>'

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
        return f'<b style="color:#1F7A4D">✓ 가입 · {_html.escape(str(v))}</b>'
    return '<b style="color:#C0444C">× 미가입</b>'

def _wc_status(rep, lookup):
    """(status, value) — status: 'on'(가입)/'off'(미가입)"""
    if lookup.startswith('diag:'):
        kind=lookup[5:]
        _pv={_it.get('t'):_it.get('v') for _it in rep.get('p5_own',[])}
        def _pvv(t):
            v=_pv.get(t); return str(v) if (v and str(v)!='미가입') else ''
        if kind=='암':
            pairs=[('암진단비', _cov_val(rep,'암','일반암','암진단')),
                   ('통합암진단비', _cov_val(rep,'암','통합암','통합전이암')),
                   ('유사암진단비', _cov_val(rep,'암','유사암')),
                   ('고액암진단비', _cov_val(rep,'암','고액암'))]
        else:
            pairs=[('뇌혈관', _pvv('뇌혈관진단비')),
                   ('뇌졸증', _pvv('뇌졸증진단비')),
                   ('뇌출혈', _pvv('뇌출혈진단비')),
                   ('협심증', _cov_val(rep,'심장','협심증')),
                   ('심부전', _cov_val(rep,'심장','심부전')),
                   ('염증', _cov_val(rep,'심장','염증','심장염증')),
                   ('부정맥', _cov_val(rep,'심장','부정맥')),
                   ('허혈성', _pvv('허혈성 진단비')),
                   ('급성심근경색', _pvv('급성심근경색'))]
        # ★2026.07.12 지점장: 뇌·심 진단비는 '보유한 담보만' 유동 표시(빈 담보는 행 자체를 없앤다).
        #   → 빈 자리에 CI 행(중대한OO)을 따로 박을 수 있다. CI 급성심근 + 일반 급성심근은 별도 행.
        if kind!='암':
            _own=[(l,v) for l,v in pairs if v]
            if _own: pairs=_own
        return ('list', [(l,(v or '')) for l,v in pairs])
    for c in rep.get('chiryo',[]):
        if c.get('name')==lookup:
            val=c.get('value')
            if not val or val=='미가입': return ('off','')
            return ('on', val)
    return ('off','')

def _wc_raw(rep, lookup):
    """워크시트 흰칸용 원본담보명 (chiryo raw). 없으면 ''"""
    for c in rep.get('chiryo',[]):
        if c.get('name')==lookup:
            return c.get('raw') or ''
    return ''

def _wcard(rep, title, desc, lookup, mode):
    # mode: 'diag'(초록·✓·가입/미가입) / 'main'(빨강·×·미가입/가입) / 'na'(회색·상태). 박스=빈칸(워크시트).
    if mode=='na':
        _rw=_wc_raw(rep, lookup)
        _dn=(f'<div class="wdn">{_html.escape(_rw)}</div>') if _rw else ''
        _half=' half' if lookup in ('암생활비','순환계생활비','순환계주요치료비','비급여주요치료비') else ''
        return (f'<div class="wcard plain{_half}"><div class="wct">{_html.escape(title)}</div>'
                f'<div class="wcd">{_html.escape(desc)}</div>{_dn}'
                f'<div class="wcf"><span class="wchip n">상태</span><span class="wbox"></span><span class="wunit">만원</span></div></div>')
    st,val=_wc_status(rep, lookup)
    variant,check=('green','✓') if mode=='diag' else ('red','×')
    _half=' half' if lookup in ('암주요치료비','2대주요치료비') else ''
    if st=='list':
        _any=any(v for l,v in val)
        chip=('<span class="wchip g">가입</span>' if _any else '<span class="wchip r">미가입</span>')
        _dgcls=' dgcancer' if '암' in str(lookup) else (' dgheart' if any(k in str(lookup) for k in ('뇌','심')) else '')
        _rows=''.join(f'<div class="dgrow"><span class="dglab">{_html.escape(l)}</span>'
                      f'<span class="mb">{_html.escape(v)}</span><span class="dgu">만</span></div>' for l,v in val)
        # ★CI 계약일 때만 유동 노출: CI 미니칩 + 진단명 + 진단금액 (7p 워크시트)
        _cinfo=(_CURREP or {}).get('ci',{}) if _CURREP else {}
        if str(_cinfo.get('status'))=='ci':
            _lk=str(lookup)
            for _it in _cinfo.get('items',[]):
                _nm=str(_it.get('t','')); _v=str(_it.get('v',''))
                _isam = ('암' in _nm)
                _ishs = any(k in _nm for k in ('뇌','심근','심장'))
                if ('암' in _lk and _isam) or ('뇌심' in _lk and _ishs):
                    _rows += (f'<div class="dgrow cirow"><span class="dglab"><span class="cichip">CI</span> {_html.escape(_nm)}</span>'
                              f'<span class="mb cibox">{_html.escape(_v)}</span><span class="dgu">만</span></div>')
        return (f'<div class="wcard {variant} dgcard{_dgcls}"><div class="wct">{check} {_html.escape(title)}</div>'
                f'<div class="wcd">{_html.escape(desc)}</div>'
                f'<div class="wcf">{chip}<div class="dglist">{_rows}</div></div></div>')
    chip=('<span class="wchip g">가입</span>' if st=='on' else '<span class="wchip r">미가입</span>')
    _rw=_wc_raw(rep, lookup)
    _dn=(f'<div class="wdn">{_html.escape(_rw)}</div>') if _rw else ''
    _bv=_html.escape(val) if (st=='on' and val) else ''
    _wrap=' wrap' if (_bv and len(_bv)>12) else ''
    _box=(f'<span class="wbox{_wrap}">{_bv}</span>' if _bv else '<span class="wbox"></span><span class="wunit">만원</span>')
    return (f'<div class="wcard {variant}{_half}"><div class="wct">{check} {_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>{_dn}'
            f'<div class="wcf">{chip}{_box}</div></div>')

def _wcard_fix(title, desc):
    """9페이지 비갱신 추천 카드: 상담용 빈 흰칸(값은 현장에서 기입)."""
    import html as _html
    return (f'<div class="wcard plain fx"><div class="wct">{_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>'
            f'<div class="wcf"><span class="wchip b">비갱신 추천</span>'
            f'<span class="wbox"></span><span class="wunit">만원</span></div></div>')


_CURREP=None   # build_report_pdf에서 세팅 → _wcard_fix_list 자동주입용
_PPT_MODE=False  # report_pptx가 True로 세팅 → 빈 흰칸 전체에 '.' 주입(편집칸화)

def _ac(label):
    """엑셀(rep) 담보값 자동주입: 라벨→값. 확실한 것만 채우고, 없으면 '' (빈칸=편집가능 유지).
       2026.07.11 재작성: 명시적 별칭표+정확매칭. 느슨한 부분일치 제거(골절←5대·일당←실손입원 오매칭 차단)."""
    rep=_CURREP
    if not rep: return ''
    L=str(label).strip()
    def _clean(v):
        if v is None: return ''
        s=str(v).strip()
        if s in ('','-','0','미가입','없음','X','x'): return ''
        return s
    import re as _re
    def _nrm(s): return _re.sub(r'[\s·()\[\]/]','',str(s))
    dmap=rep.get('dambo',{})
    # 워크시트 라벨 → 담보명 후보(정확, 앞이 우선). 여기 있는 라벨은 이 후보로만 매칭(오매칭 방지).
    _ALIAS={
        '종신 사망':['일반사망'], '질병 사망':['질병사망(80세)','질병사망'], '상해 사망':['상해사망'],
        '종신':['일반사망'],
        '일반암':['일반암'], '유사암':['유사암(갑.기.경.제)','유사암'], '고액암':['고액암'],
        '뇌혈관 수술비':['뇌혈관수술비'], '허혈성 수술비':['허혈성수술비'], '심장 수술비':['심장수술비'],
        '상해 후유 3%':['상해후유3%'], '상해 후유 80%':['상해후유80%'],
        '질병 후유 3%':['질병후유3%'], '질병 후유 80%':['질병후유80%'],
        '골절':['골절합산'], '5대 골절':['5대골절진단비'],
        '화상진단비':['화상진단비'], '중대화상진단비':['중증화상진단비'],
        '깁스':['깁스진단비'], '응급실':['응급실(응급)','응급실'],
        '질병 입원일당':['질병일당','질병입원일당'], '상해 입원일당':['상해일당','상해입원일당'],
        '질병 중환자실':['질병중환자실'], '상해 중환자실':['상해중환자실'],
        '대인 벌금':['대인'], '대물 벌금':['대물'], '합의금':['합의금'], '6주미만 합의금':['6주미만'],
        '변호사비':['변호사'], '자동차부상위로금':['자부상'],
        '간병인지원일당':['간병인'], '간호통합병동':['간호통합병동'],
    }
    # 1) 뇌·심 진단비 (p5_own, 담보명 확정)
    _P={'뇌혈관':'뇌혈관진단비','뇌졸증':'뇌졸증진단비','뇌출혈':'뇌출혈진단비',
        '허혈성':'허혈성 진단비','급성심근경색':'급성심근경색'}
    if L in _P:
        pv={i.get('t'):i.get('v') for i in rep.get('p5_own',[])}
        r=_clean(pv.get(_P[L]))
        return r
    # 2) 별칭표: 지정 후보로만 매칭 (있으면 그 후보 외엔 채우지 않음)
    if L in _ALIAS:
        for cand in _ALIAS[L]:
            r=_clean(dmap.get(_nrm(cand)))
            if r: return r
        return ''
    # 3) chiryo(치료비 등) 정확 일치
    for c in rep.get('chiryo',[]):
        if str(c.get('name','')).replace(' ','')==L.replace(' ',''):
            r=_clean(c.get('value'))
            if r: return r
    # 4) 전체 담보맵 정확(정규화) 일치
    r=_clean(dmap.get(_nrm(L)))
    return r

def _wcard_fix_list(title, desc, rows):
    """세분화 카드: 항목별 라벨 + 개별 흰칸. 엑셀에 값 있으면 자동주입, 없으면 빈칸(현장 기입)."""
    import html as _html
    def _cell(r):
        v=_ac(r)
        inner=_html.escape(v) if v else ''
        return (f'<div class="fxsub"><span class="lbl">{_html.escape(r)}</span>'
                f'<span class="mb">{inner}</span></div>')
    # ★항목 10개 이상이면 형식(박스칸) 유지하고 가로 2열로 배치 (지점장 2026.07.12)
    if len(rows)>=10:
        _tr=''
        for i in range(0,len(rows),2):
            _l=f'<td class="fx2c">{_cell(rows[i])}</td>'
            _r=f'<td class="fx2c">{_cell(rows[i+1])}</td>' if i+1<len(rows) else '<td class="fx2c"></td>'
            _tr+=f'<tr>{_l}{_r}</tr>'
        lines=f'<table class="fx2">{_tr}</table>'
        return (f'<div class="wcard plain fx tall two"><div class="wct">{_html.escape(title)}</div>'
                f'<div class="wcd">{_html.escape(desc)}</div>{lines}</div>')
    lines=''.join(_cell(r) for r in rows)
    return (f'<div class="wcard plain fx tall"><div class="wct">{_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>'
            f'<div class="fxlist">{lines}</div></div>')


def _wcard_fix_group(title, desc, groups):
    """★진단비 카드: 암·뇌·심 그룹 구분 + 가로 2열 박스칸 (지점장 2026.07.12)"""
    import html as _html
    def _cell(r):
        v=_ac(r)
        return (f'<div class="fxsub"><span class="lbl">{_html.escape(r)}</span>'
                f'<span class="mb">{_html.escape(v) if v else ""}</span></div>')
    body=''
    for gname, rows in groups:
        tr=''
        for i in range(0,len(rows),2):
            l=f'<td class="fx2c">{_cell(rows[i])}</td>'
            r=f'<td class="fx2c">{_cell(rows[i+1])}</td>' if i+1<len(rows) else '<td class="fx2c"></td>'
            tr+=f'<tr>{l}{r}</tr>'
        body+=(f'<div class="fxg">{_html.escape(gname)}</div>'
               f'<table class="fx2">{tr}</table>')
    return (f'<div class="wcard plain fx tall two"><div class="wct">{_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>{body}</div>')


def _wcard_sj(rep, title, desc, lookup):
    # 산정특례 = 나란히 2칸, 회색 '상태' 칩, 빈 박스 (레퍼런스 그대로)
    _rw=_wc_raw(rep, lookup)
    _dn=(f'<div class="wdn">{_html.escape(_rw)}</div>') if _rw else ''
    return (f'<div class="wcard plain sj"><div class="wct">{_html.escape(title)}</div>'
            f'<div class="wcd">{_html.escape(desc)}</div>{_dn}'
            f'<div class="wcf"><span class="wchip n">상태</span><span class="wbox"></span><span class="wunit">만원</span></div></div>')

def build_report_pdf(rep, out):
    """rep: 리포트 데이터 dict (아래 sample_rep 구조). out: 저장 경로(.pdf)"""
    global _CURREP; _CURREP=rep
    try:
        from ga_tables import ga_pages_html
        _ga_html = ga_pages_html()
    except Exception:
        _ga_html = ''
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
    # ── 보험료 막대 (★10만원 초과 = 빨강 경고) ──
    mx=max((c['amt'] for c in rep['premium_bars']),default=1)
    def _bcol(c):
        if c['amt']>=100000: return GAP          # 10만원 이상 → 빨강
        return BLUE if c['renew'] else NAVY2
    def _vcol(c):
        if c['amt']>=100000: return GAP
        return BLUE if c['renew'] else NAVY
    bars=''.join(
        f'''<tr><td class="bl">{_html.escape(c["nm"])}</td>
<td class="track-td"><div class="track"><div class="fill" style="width:{c["amt"]/mx*100:.1f}%;background:{_bcol(c)}"></div></div></td>
<td class="bv" style="color:{_vcol(c)}">{c["amt"]:,}{'<span class="warn10">▲</span>' if c['amt']>=100000 else ''}</td></tr>'''
        for c in rep['premium_bars'])
    # ── 도넛 ──
    def dcolor(p): return '#1F7A4D' if p>=70 else ('#D08B1F' if p>=40 else '#C0242E')
    donuts=''
    drows=''
    per=5
    for i in range(0,len(rep['donuts']),per):
        cells=''.join(
            f'<td class="dcell"><div>{_donut(d["pct"],dcolor(d["pct"]),over=d.get("over",False))}</div><div class="dn">{_html.escape(d["name"])}</div></td>'
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
    # ★2026.07.12 지점장: 3p 카드 = 중대한OO 4종 고정
    _JUNG=[('중대한 암진단','중대한 암'),('중대한 뇌졸증','중대한 뇌졸증'),
           ('중대한 뇌출혈','중대한 뇌출혈'),('중대한 급성심근경색','중대한 급성심근')]
    def _jval(key):
        for c in rep.get('chiryo',[]):
            if str(c.get('name','')).strip()==key: return str(c.get('value','미가입'))
        for cat in rep.get('coverage',[]):
            for it in cat.get('items',[]):
                if str(it.get('name','')).strip()==key:
                    v=it.get('value') or it.get('amt')
                    if v: return str(v)
        return '미가입'
    crows=''.join(f'<td class="cc"><div class="cl">{_html.escape(lb)}</div><div class="cval">{_cv(_jval(k))}</div></td>'
                  for lb,k in _JUNG)
    crows=f'<tr>{crows}</tr>'

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

    # ── 보장 진단 코멘트 (숫자 카드형 · 60대 가독) ──
    _sh=[_html.escape(s['h']) for s in rep.get('strength',[])][:4]
    _ok = '충실' if gap_cnt==0 else '보완 필요'
    _okc = 'ok' if gap_cnt==0 else 'ng'
    _cards=(f'<div class="sumcard"><div class="sv">{n_contract}<small>건</small></div><div class="sk">보유 계약</div></div>'
            f'<div class="sumcard"><div class="sv">{premium:,}<small>원</small></div><div class="sk">월 납입보험료</div></div>'
            f'<div class="sumcard {_okc}"><div class="sv">{gap_cnt}<small>개</small></div><div class="sk">보장 공백</div></div>'
            f'<div class="sumcard {_okc}"><div class="sv sm">{_ok}</div><div class="sk">종합 판정</div></div>')
    _strong=''
    if _sh:
        _chips=''.join(f'<span class="stchip">{h}</span>' for h in _sh)
        _strong=f'<div class="sumrow"><span class="sumlb ok">✓ 강점</span><span class="sumtx">{_chips}</span></div>'
    _cirow=''
    if ci.get('present'):
        _cirow=(f'<div class="sumrow"><span class="sumlb ci">CI 선지급</span>'
                f'<span class="sumtx">중대질병 진단 시 <b>즉시 지급</b> · 진단 후에도 사망보장 '
                f'<b>{_html.escape(ci.get("residual","-"))}</b> 유지</span></div>')
    _gaprow=(f'<div class="sumrow"><span class="sumlb {_okc}">{"✓ 결론" if gap_cnt==0 else "! 결론"}</span>'
             f'<span class="sumtx">{"핵심담보 균형이 <b>양호</b>합니다. 지금 보장을 <b>유지</b>하십시오." if gap_cnt==0 else f"공백 <b>{gap_cnt}개</b> — 상담을 통한 <b>보완</b>이 필요합니다."}</span></div>')
    _citab=('<div class="sect" style="margin-top:4.5mm">CI보험 — 장점 vs 단점 <span>CRITICAL ILLNESS</span></div>'
            '<table class="citab">'
            '<tr><th class="cig">■ 장점</th><th class="cir">■ 단점 — 이것 때문에 못 받는다</th></tr>'
            '<tr><td class="cig">'
            '<div class="cil"><b>주계약 = 중대한 담보가 종신보장</b> — 만기 없이 평생 유지</div>'
            '<div class="cil"><b>원스톱</b> — 사망 + 중대질병 한 상품</div>'
            '<div class="cil"><b>선지급</b> — 진단 즉시 사망보험금 50~80% 지급</div>'
            '<div class="cil"><b>잔여 사망보장 유지</b> — 선지급 후에도 남은 사망금 지급</div>'
            '<div class="cil"><b>주계약 중대한 담보가 종신보장</b> — 만기 없이 평생 유지</div>'
            '<div class="cil"><b>납입면제</b> 특약 결합 용이</div>'
            '</td><td class="cir">'
            '<div class="cil r"><b>암·뇌·심 중 딱 1가지만</b> — 먼저 발생한 1회만 선지급, 이후 다른 중증질환은 <b>중복 보장 불가</b></div>'
            '<div class="cil r"><b>[뇌] 25% 후유장해가 영구적</b>이어야 지급 — 재활 후에도 신경계 장해(ADLs) 25% 이상 <b>영구 지속</b> + 수시간호 상태. 일과성 허혈발작(TIA)·외상성 출혈 제외</div>'
            '<div class="cil r"><b>[심장] 효소검사 3가지 전부 충족</b> — ①전형적 흉통 ②새 심전도 변화(ST상승·Q파·T파역전) ③심근효소 상승(CK-MB·트로포닌 I·T). <b>모든 협심증 제외</b></div>'
            '<div class="cil r"><b>[암] 침윤·파괴적 증식만</b> — 상피내암·경계성·초기전립선암·피부암 제외</div>'
            '<div class="cil r"><b>선지급 = 사망보험금 감소</b> · 보험료 비쌈 · 분쟁 다발</div>'
            '</td></tr></table>'
            '<div class="cint">※ 출처: 26년 바름 교육자료 — CI보험 완벽가이드(기초약관·장점·단점). 가입 시기·보험사별 약관 상이 → 원문 확인 필수.</div>')
    comment_html=_citab

    css=f'''
@page {{ size:A4; margin:0; }}
* {{ margin:0; padding:0; box-sizing:border-box; font-family:'NanumSquareRound','Noto Sans CJK KR',sans-serif; }}
body {{ color:{INK}; }}
.pg {{ width:210mm; height:297mm; position:relative; page-break-after:always; background:#fff; }}
.pg:last-child {{ page-break-after:auto; }}
.top {{ background:#fff; color:{NAVY}; padding:9mm 11mm 8mm; position:relative; }}
.top .bar {{ position:absolute; left:0; bottom:0; width:100%; height:1.4mm; background:linear-gradient(90deg,{GOLD},{GOLDD} 55%,transparent); }}
.top .eb {{ font-size:9pt; letter-spacing:2px; color:{GOLDD}; font-weight:700; }}
.top .nm {{ font-size:20pt; font-weight:800; margin-top:2mm; color:{NAVY}; }}
.top .nm b {{ color:{GOLDD}; }}
.top .pgn {{ position:absolute; right:11mm; top:9mm; text-align:right; font-size:9pt; color:{MUT}; }}
.top .pgn b {{ display:block; font-size:22pt; color:{NAVY}; font-weight:800; }}
.body {{ padding:7mm 11mm; }}
.sbody {{ padding:3mm 8mm 4mm; }}
.scv2 {{ display:flex; gap:4mm; }}
.scvcol {{ flex:1; }}
.scvhd {{ font-size:9.5pt; font-weight:800; padding-bottom:1mm; border-bottom:2pt solid {GOLD}; margin-bottom:1.4mm; }}
.scvhd.brain {{ color:#1F5FA8; }} .scvhd.heart {{ color:{GAP}; }}
.scvt {{ width:100%; border-collapse:collapse; font-size:8.4pt; }}
.scvt th {{ background:{NAVY}; color:#fff; padding:1.1mm 0.7mm; font-size:7.6pt; text-align:center; font-weight:800; }}
.scvt th.dl {{ text-align:left; }}
.scvt th:last-child {{ background:{GOLDD}; }}
.scvt td {{ border-top:0.4pt solid {LINE}; padding:0.9mm 0.7mm; text-align:center; }}
.scvt td.dl {{ text-align:left; }}
.scvt td.dl .cd {{ font-size:6.8pt; color:{MUT}; }}
.scvt tr.grp td {{ background:#EEF1F6; font-size:7pt; font-weight:800; color:{NAVY}; text-align:left; padding:0.8mm 0.7mm; }}
.scvt tr.own td {{ background:#FBF6E6; border-top:1.4pt solid {GOLD}; border-bottom:1.4pt solid {GOLD}; }}
.scvt tr.own td:first-child {{ border-left:1.4pt solid {GOLD}; }}
.scvt tr.own td:last-of-type {{ border-right:1.4pt solid {GOLD}; }}
.scvt td:last-child {{ background:#F2F6F1; }}
.scvt tr.own td:last-child {{ background:#F0EFD8; }}
.on {{ color:#1F7A4D; font-weight:800; }} .off {{ color:#B9C2CE; }}
.hold {{ color:{GOLDD}; font-weight:700; font-size:6pt; }}
.chip {{ background:{GOLD}; color:#fff; font-size:5.4pt; font-weight:800; padding:0.2mm 1mm; border-radius:2mm; }}
.cichip {{ background:{NAVY}; color:#fff; font-size:5.4pt; font-weight:800; padding:0.2mm 1.2mm; border-radius:2mm; margin-left:1.2mm; }}
.scvt .amtbox.cibox {{ background:#F4F7FB; border-color:{NAVY}; min-width:9mm; }}
.scvt .amt {{ color:{NAVY}; font-weight:800; font-size:8.6pt; margin-left:0.8mm; }}
.scvt .amtbox {{ display:inline-block; min-width:9mm; height:6.6mm; line-height:6.6mm; border:0.6pt solid #C9CFD8; border-radius:1mm; padding:0 1.2mm; margin-left:0.8mm; background:#fff; color:{NAVY}; font-weight:800; font-size:8.4pt; text-align:right; vertical-align:middle; white-space:nowrap; }}
.scvleg {{ font-size:6.2pt; color:{MUT}; margin:1.5mm 0; }} .own2 {{ color:{GOLDD}; font-weight:700; }}
.scvnote {{ font-size:5.6pt; line-height:1.35; color:{INK}; background:#F6F8FB; border-left:2.2pt solid {NAVY}; padding:1.5mm 2mm; border-radius:1.4mm; }}
.scvnote b {{ color:{NAVY}; }} .scvnote b.r {{ color:{GAP}; }}
.rngbox {{ border:0.8pt solid {LINE}; border-radius:1.8mm; background:#FBFCFD; padding:2mm 2.4mm; margin-top:2mm; }}
.rngt {{ font-size:7.4pt; font-weight:800; color:{NAVY}; margin-bottom:1.4mm; }}
.rngtab {{ width:100%; border-collapse:collapse; }}
.rgcol {{ width:27mm; vertical-align:middle; }}
.rgtxt {{ vertical-align:middle; padding-left:2mm; }}
.rngc {{ position:relative; height:26mm; }}
.rgo {{ position:absolute; bottom:0; border-radius:50%; border:1pt solid; text-align:center; }}
.rgo .rgp {{ position:absolute; left:0; right:0; top:1.2mm; font-size:6.6pt; font-weight:900; }}
.rgo.o1 {{ left:0; width:26mm; height:26mm; background:#E6F1FB; border-color:#1F5FA8; }}
.rgo.o1 .rgp {{ color:#1F5FA8; }}
.rgo.o2 {{ left:4mm; width:18mm; height:18mm; background:#C9DFF5; border-color:#1F5FA8; }}
.rgo.o2 .rgp {{ color:#14406F; }}
.rgo.o3 {{ left:9mm; width:8mm; height:8mm; background:#1F5FA8; border-color:#14406F; }}
.rgo.o3 .rgp {{ color:#fff; top:1mm; font-size:5.6pt; }}
.rgo.o1.h {{ background:#FBEBEB; border-color:{GAP}; }} .rgo.o1.h .rgp {{ color:{GAP}; }}
.rgo.o2.h {{ background:#F2C9C9; border-color:{GAP}; }} .rgo.o2.h .rgp {{ color:#8E2B2B; }}
.rgo.o3.h {{ background:{GAP}; border-color:#8E2B2B; }} .rgo.o3.h .rgp {{ color:#fff; }}
.rgl {{ font-size:6.8pt; font-weight:700; color:{INK}; line-height:1.35; padding:1mm 0; border-bottom:0.4pt solid #EDF0F3; white-space:nowrap; }}
.rgl:last-child {{ border-bottom:none; }}
.rgl .cdx {{ font-size:5.8pt; color:{MUT}; }}
.rgl .pc {{ color:#1F5FA8; margin-left:1mm; }} .rgl .pc.r {{ color:{GAP}; }}
.rgl b {{ color:{NAVY}; }}
.rgl .d {{ display:inline-block; width:2mm; height:2mm; border-radius:50%; margin-right:1.2mm; }}
.rgl .d1 {{ background:#E6F1FB; border:0.6pt solid #1F5FA8; }}
.rgl .d2 {{ background:#C9DFF5; border:0.6pt solid #1F5FA8; }}
.rgl .d3 {{ background:#1F5FA8; }}
.rgl .d1h {{ background:#FBEBEB; border:0.6pt solid {GAP}; }}
.rgl .d2h {{ background:#F2C9C9; border:0.6pt solid {GAP}; }}
.rgl .d3h {{ background:{GAP}; }}
.scvbot {{ font-size:6pt; line-height:1.45; color:{INK}; background:#FBF7EE; border:0.6pt solid {GOLD}; border-radius:1.6mm; padding:1.6mm 2.4mm; margin-top:1.6mm; }}
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

.cvpg {{ position:relative; overflow:hidden; }}
.cvbar {{ height:14mm; background:{NAVY}; border-bottom:2.4mm solid {GOLD}; }}
.cvfootbar {{ position:absolute; left:0; bottom:0; width:100%; height:6mm; background:{NAVY}; }}
.cvmark {{ position:absolute; right:10mm; top:46mm; font-size:60pt; font-weight:900; color:#F2F4F7; letter-spacing:-3px; }}
.cveyebrow {{ font-size:9.5pt; font-weight:800; color:{GOLDD}; letter-spacing:4px; margin-top:26mm; }}
.cvrule {{ width:34mm; height:1.8mm; background:{GOLD}; margin:5mm 0 6mm; }}
.cvnamebox {{ margin-top:9mm; padding:5mm 6mm; background:#F4F7FB; border-left:4pt solid {NAVY}; border-radius:0 2.5mm 2.5mm 0; }}
.cvnamebox .g {{ font-size:44pt; font-weight:900; color:{NAVY}; letter-spacing:-1px; }}
.cvnamebox .s {{ font-size:17pt; font-weight:800; color:{GOLDD}; margin-left:3mm; }}
.cvbody {{ padding:14mm 16mm 12mm; height:277mm; display:flex; flex-direction:column; position:relative; }}
.cvbrand {{ font-size:14pt; font-weight:800; color:{NAVY}; letter-spacing:5px; }}
.cvbrand .ln {{ width:24mm; height:1.6mm; background:{GOLD}; margin-top:2.5mm; }}
.cvtitle {{ font-size:60pt; font-weight:900; color:{NAVY}; line-height:1.05; margin-top:3mm; letter-spacing:-2px; position:relative; }}
.cvname {{ font-size:30pt; font-weight:800; color:{NAVY}; line-height:1.1; margin-top:12mm; text-align:right; }}
.cvname .g {{ color:#C9A15A; font-size:72pt; letter-spacing:-1px; }}
.cvsub {{ font-size:13pt; font-weight:700; color:{NAVY}; margin-top:6mm; }}
.cvhr {{ border-top:1.6pt solid {GOLD}; margin:3mm 0 5mm; }}
.cvstats {{ display:flex; gap:4mm; margin-top:12mm; }}
.cvst {{ flex:1; border:1pt solid {LINE}; border-top:2.4pt solid {GOLD}; border-radius:2mm; padding:3.4mm 3.6mm; background:#fff; }}
.cvst .k {{ font-size:9pt; color:{MUT}; font-weight:700; }}
.cvst .v {{ font-size:17pt; font-weight:800; color:{NAVY}; margin-top:2mm; }}
.cvst .v small {{ font-size:9pt; color:{MUT}; font-weight:600; }}
.cvspacer {{ flex:1; }}
.cvname2 {{ font-size:26pt; font-weight:900; color:{NAVY}; line-height:1.15; text-align:right; margin-top:10mm; margin-bottom:4mm; }}
.cvname2 .g {{ color:#C9A15A; font-size:52pt; font-weight:900; letter-spacing:-1px; line-height:1.1; }}
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
.wslack .lackbox {{ height:30mm; }}
.wstalk {{ background:#EAF2FB; border-left:2.6pt solid {BLUE}; border-radius:1.4mm; padding:2mm 3mm; font-size:7.6pt; line-height:1.5; color:{NAVY}; font-style:italic; }}
.wstalk .h {{ font-style:normal; font-weight:800; color:{BLUE}; }}
.ws3 {{ display:flex; gap:1.8mm; align-items:stretch; }}
.wscol.wsmain {{ display:flex; flex-direction:column; }}
.wsmain .wscap {{ flex:0 0 auto; }}
.wsmain .wsr {{ min-height:39mm; display:flex; flex-direction:column; justify-content:center; }}
.wsr .box {{ display:block; margin-top:1.6mm; border:0.6pt solid {LINE}; border-radius:1mm; height:6mm; background:#FCFDFE; }}
.wsmid {{ flex:0 0 14mm; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; }}
.wsmid .lb {{ font-size:6.5pt; color:{MUT}; letter-spacing:1px; font-weight:700; }}
.wsmid .nm {{ font-size:22pt; font-weight:800; color:{NAVY}; line-height:1.02; letter-spacing:0; margin:2mm 0 0; }}
.wsmid .nmsub {{ font-size:9pt; font-weight:800; color:{GOLDD}; margin:1.5mm 0 2mm; }}
.wsmid .cnt {{ font-size:10pt; font-weight:800; color:{NAVY}; }}
.wsmid .cnt small {{ display:block; font-size:6.2pt; color:{MUT}; font-weight:600; }}
.wssj {{ margin-top:1mm; border:1pt solid #D8B65A; border-radius:2mm; background:#FDFAF0; padding:2mm 3mm 2.4mm; }}
.wcard {{ border:0.9pt solid {LINE}; border-radius:2mm; padding:1.4mm 2mm; margin-bottom:1.2mm; min-height:12mm; display:flex; flex-direction:column; }}
.wcard.half {{ min-height:26mm; padding:1.8mm 2.2mm; }}
.wcard.half .wcd {{ margin:0.8mm 0 auto; }}
.wcard.half .wbox {{ height:11mm; line-height:11mm; }}
.wcard.green {{ border-color:#3E9A63; background:#F4FBF7; }}
.wcard.red {{ border-color:#DE9A9A; background:#FDF4F3; }}
.wcard.plain {{ border-color:{LINE}; background:#FCFDFE; }}
.wcard .wct {{ font-size:11.5pt; font-weight:800; color:{NAVY}; }}
.wcard.green .wct {{ color:#1F7A4D; }}
.wcard.red .wct {{ color:#C0444C; }}
.wcard .wcd {{ font-size:8.8pt; color:{MUT}; margin:1.2mm 0 auto; }}
.wcard .wdn {{ font-size:9pt; font-weight:700; color:#2B3A52; margin:1.4mm 0 0.8mm; line-height:1.25; word-break:break-all; }}
.wcard .wcf {{ display:flex; align-items:center; gap:2mm; margin-top:1.2mm; }}
.wchip {{ font-size:9pt; font-weight:800; color:#fff; padding:1.1mm 3.2mm; border-radius:1.5mm; white-space:nowrap; }}
.wchip.g {{ background:#3E9A63; }}
.wchip.r {{ background:#C0444C; }}
.wchip.n {{ background:#9AA5B4; }}
.wchip.b {{ background:#2E5A88; }}
.wscap.n2 {{ background:#2E5A88; color:#fff; }}
.wcard.fx {{ min-height:46mm; }}
.wcard.fx.tall {{ min-height:auto; }}
.fxlist {{ margin-top:1.6mm; }}
.fxsub {{ display:flex; align-items:center; gap:2mm; margin:0.7mm 0; font-size:8.2pt; }}
.fxsub .lbl {{ flex:0 0 30mm; font-weight:700; color:#2B3A52; font-size:8pt; white-space:normal; word-break:keep-all; line-height:1.15; }}
.fxsub .mb {{ flex:1; border:0.7pt solid {LINE}; border-radius:1mm; height:4.8mm; line-height:4.8mm; padding:0 2mm; background:#fff; text-align:right; font-size:8.2pt; font-weight:800; color:{NAVY}; white-space:nowrap; overflow:hidden; }}
.fxg {{ font-size:6.8pt; font-weight:900; color:{NAVY}; background:#EEF1F6; border-radius:1mm; padding:0.4mm 1.2mm; margin:0.8mm 0 0.3mm; }}
/* ★2열형: 형식(박스칸) 유지 + 가로 2개씩 */
.fx2 {{ width:100%; border-collapse:separate; border-spacing:1mm 0; table-layout:fixed; margin-top:0; }}
.fx2 td.fx2c {{ width:50%; vertical-align:middle; overflow:hidden; }}
.wcard.two .fxsub {{ display:table; width:100%; margin:0.1mm 0; }}
.wcard.two .fxsub .lbl {{ display:table-cell; width:44%; flex:none; font-size:6.4pt; white-space:nowrap; overflow:hidden; vertical-align:middle; padding-right:1mm; }}
.wcard.two .fxsub .mb {{ display:table-cell; width:56%; flex:none; height:3.8mm; line-height:3.8mm; font-size:6.8pt; padding:0 1mm; overflow:hidden; vertical-align:middle; }}
.wsectcap {{ margin-bottom:2.6mm; border-radius:1.4mm; }}
.wsectcap.gap2 {{ margin-top:0.9mm; }}
.wstalk2 {{ background:#EAF2FB; border-left:2.6pt solid {BLUE}; border-radius:1.4mm; padding:2.2mm 3.2mm; font-size:8pt; line-height:1.5; color:{NAVY}; margin-top:2.6mm; }}
.wstalk2 b {{ color:{BLUE}; }}
.jgbig .wstalk2 {{ font-size:11.3pt; line-height:1.5; margin-top:2.0mm; padding:2.3mm 3.2mm; }}
.jgbig .fixnote .fxrow {{ font-size:10.9pt; line-height:1.48; margin:1.4mm 0; }}
.steprow {{ display:flex; gap:1.8mm; margin-top:1.6mm; }}
.stepc {{ flex:1; border:0.7pt solid {LINE}; border-radius:2mm; background:#FCFDFE; padding:1.5mm 2mm 1.8mm; font-size:6.8pt; line-height:1.32; color:{MUT}; }}
.steprow.big .stepc {{ font-size:6.8pt; line-height:1.3; padding:1.5mm 2mm 1.6mm; }}
.steprow.big .stepc b {{ font-size:8pt; }}
.note24 {{ text-align:center; font-size:7.8pt; font-weight:700; color:{NAVY}; background:#FBF3E6; border:0.7pt solid {GOLD}; border-radius:1.4mm; padding:1mm; margin-top:0.7mm; }}
.note24 b {{ color:{GOLDD}; }}
.memobox {{ border:0.8pt solid {LINE}; border-radius:2mm; height:46mm; margin-top:1.5mm; background:repeating-linear-gradient(#fff, #fff 7mm, #EDF1F6 7mm, #EDF1F6 7.4mm); }}
.gentab td:last-child, .gentab th:last-child {{ text-align:left; }}
.silgenpg .st td, .silgenpg .st th {{ font-size:7pt; padding:0.5mm 1.3mm; }}
.silgenpg .note {{ font-size:10pt; padding:2.2mm 3mm; line-height:1.45; }}
.silgenpg .silverd {{ padding:4mm 4.5mm; margin:4mm 0; }}
.silgenpg .silverd .svhead {{ font-size:18pt; }}
.silgenpg .silverd .svbody {{ font-size:12.5pt; line-height:1.5; margin-top:2mm; }}
.silgenpg .wssj .fxrow {{ font-size:8.4pt; line-height:1.42; }}
.g5tab td, .g5tab th {{ font-size:8.2pt; padding:1.05mm 1.6mm; }}
.g5tab td:last-child, .g5tab th:last-child, .g5tab td:nth-child(2), .g5tab th:nth-child(2) {{ text-align:left; }}
.silgenpg .memobox {{ height:10mm; }}
.silverd .svtx {{ display:inline-block; margin-top:1.2mm; font-size:10pt; font-weight:700; color:{NAVY}; }}
.silverd.g5 .svtx {{ color:{GAP}; }}
.genhit td {{ background:#FFF3D6 !important; font-weight:800; color:{NAVY} !important;
   border-top:2.4pt solid {GAP} !important; border-bottom:2.4pt solid {GAP} !important; }}
.genhit td:first-child {{ border-left:2.4pt solid {GAP} !important; }}
.genhit td:last-child {{ border-right:2.4pt solid {GAP} !important; }}
.g5tab .gyeo td {{ background:#EAF2FB; }}
.g5tab .jung td {{ background:#FBF7EE; }}
.g5tab td, .g5tab th {{ font-size:10.6pt; line-height:1.45; padding:2.0mm 2.2mm; }}
.silgenpg .fixnote .fxrow {{ font-size:10.6pt; line-height:1.5; margin:1.5mm 0; }}
.silgenpg .fixnote .cap {{ font-size:11.5pt; }}
.silgenpg .note.bigfact {{ font-size:10.4pt; line-height:1.5; }}
.gentab .smn, .g5tab .smn {{ font-size:10pt; color:#6B7686; font-weight:700; }}
.gentab .smn {{ color:#C0444C; font-size:10pt; }}
.gentab .disc {{ display:inline-block; margin-top:0.5mm; font-size:7.4pt; font-weight:800; color:#1456B0; background:#EAF2FB; border-radius:1mm; padding:0.4mm 1.2mm; }}
.barttl {{ font-size:11pt; font-weight:800; color:{NAVY}; margin-top:2mm; }}
.barttl span {{ font-size:8.5pt; color:{MUT}; font-weight:700; }}
.g5wrap {{ display:table; width:100%; border-spacing:2mm 0; margin:2mm 0; }}
.g5card {{ display:table-cell; width:50%; border-radius:2.5mm; padding:2.2mm 3mm; border:1.2pt solid; vertical-align:top; }}
.g5card.jungc {{ border-color:#9C7C32; background:#FBF7EE; }}
.g5card.bic {{ border-color:#C0444C; background:#FCF3F3; }}
.g5h {{ font-size:14pt; font-weight:800; color:{NAVY}; }}
.g5h span {{ font-size:10pt; color:{MUT}; font-weight:700; margin-left:1.5mm; }}
.g5sub {{ font-size:9pt; color:{MUT}; font-weight:700; margin-bottom:1.5mm; }}
.g5big {{ font-size:13pt; font-weight:800; color:#9C7C32; margin:0.8mm 0 1.4mm; }}
.g5big.bad {{ color:{GAP}; }}
.g5big b {{ font-size:21pt; }}
.g5line {{ font-size:10.5pt; font-weight:700; color:{INK}; margin-top:0.8mm; }}
.g5line b {{ color:{NAVY}; font-size:13.5pt; }}
.g5line b.bad {{ color:{GAP}; }}
.g5nt {{ font-size:8.5pt; color:{MUT}; font-weight:700; }}
.g5nt {{ font-size:8.5pt; color:{MUT}; font-weight:700; }}
.barwrap {{ display:table; width:100%; border-spacing:2mm 0; margin:1.2mm 0 0.6mm; }}
.barcol {{ display:table-cell; width:20%; vertical-align:bottom; text-align:center; }}
.barlb.vt {{ writing-mode:vertical-rl; text-orientation:upright; letter-spacing:0.2pt; margin:0.8mm auto 0; font-size:8.4pt; line-height:1; white-space:nowrap; }}
.barlb.vt.bad {{ color:{GAP}; }}
.hbar {{ border-radius:1.5mm 1.5mm 0 0; padding:0.8mm 0; text-align:center; }}
.hbar span {{ font-size:11pt; font-weight:800; color:#fff; }}
.b30 {{ height:6mm; background:#8FA8C8; }}
.b40 {{ height:8mm; background:#C9A15A; }}
.b50 {{ height:10mm; background:#D08B4A; }}
.b60 {{ height:12mm; background:#C0444C; }}
.b90 {{ height:18mm; background:#8B1A1A; }}
.barlb {{ font-size:6.6pt; font-weight:800; color:{NAVY}; margin-top:0.4mm; line-height:1.15; }}
.g5warn {{ margin-top:0.5mm; padding:0.9mm 2mm; border-radius:2mm; background:#FCF3F3; border-left:3.5pt solid {GAP}; font-size:7.6pt; font-weight:700; color:{INK}; line-height:1.3; }}
.g5warn b {{ color:{GAP}; }}
.myeon {{ margin-top:2mm; border:1.2pt solid {GAP}; border-radius:2mm; background:#FCF3F3; padding:2mm 2.5mm; }}
.mytt {{ font-size:9.4pt; font-weight:800; color:{GAP}; margin-bottom:1mm; }}
.mytab {{ display:table; width:100%; border-spacing:1.5mm 0; }}
.myc {{ display:table-cell; width:20%; vertical-align:top; background:#fff; border:0.8pt solid #E8B4B4; border-radius:1.6mm; padding:1.8mm 1.6mm; text-align:center; }}
.myh {{ font-size:8.6pt; font-weight:800; color:{GAP}; margin-bottom:1mm; line-height:1.25; }}
.myd {{ font-size:7pt; font-weight:700; color:{INK}; line-height:1.28; }}
.mynt {{ font-size:8pt; font-weight:600; color:{INK}; margin-top:1.5mm; }}
.myeon {{ border:1.2pt solid {GAP}; border-radius:2mm; padding:1.6mm 2.2mm; margin-top:1.2mm; background:#FCF3F3; }}
.mytt {{ font-size:9.4pt; font-weight:800; color:{GAP}; margin-bottom:1mm; }}
.mytab {{ display:table; width:100%; border-spacing:1.5mm 0; }}
.myc {{ display:table-cell; width:20%; background:#fff; border:0.8pt solid #E8B4B4; border-radius:1.6mm; padding:1.3mm 1.2mm; vertical-align:top; text-align:center; }}
.myh {{ font-size:7.6pt; font-weight:800; color:{GAP}; margin-bottom:0.6mm; }}
.myd {{ font-size:7pt; font-weight:700; color:{INK}; line-height:1.28; }}
.mynt {{ font-size:7pt; color:{MUT}; font-weight:600; margin-top:1mm; }}
.renote {{ border:1pt solid {LINE}; border-left:3.5pt solid {BLUE}; border-radius:2mm; padding:1.5mm 2.4mm; margin-top:1.2mm; background:#EAF2FB; }}
.rentt {{ font-size:9pt; font-weight:800; color:{NAVY}; margin-bottom:0.8mm; }}
.renrow {{ font-size:8.6pt; font-weight:700; color:{INK}; margin:2.2mm 0; }}
.renrow b {{ color:{BLUE}; }}
.rennt {{ font-size:7.4pt; color:{MUT}; font-weight:600; margin-top:2mm; }}
.sumwrap {{ display:table; width:100%; border-spacing:3mm 0; margin:2.5mm 0; }}
.sumcard {{ display:table-cell; width:25%; border:1pt solid {LINE}; border-radius:2.5mm; padding:4mm 3mm; text-align:center; background:#F7F9FC; }}
.sumcard.ok {{ border-color:{GOOD}; background:#F1F8F4; }}
.sumcard.ng {{ border-color:{GAP}; background:#FCF3F3; }}
.sumcard .sv {{ font-size:26pt; font-weight:800; color:{NAVY}; line-height:1.1; }}
.sumcard .sv small {{ font-size:12pt; font-weight:700; color:{MUT}; margin-left:1mm; }}
.sumcard .sv.sm {{ font-size:20pt; }}
.sumcard.ok .sv {{ color:{GOOD}; }}
.sumcard.ng .sv {{ color:{GAP}; }}
.sumcard .sk {{ font-size:11pt; font-weight:700; color:{MUT}; margin-top:2mm; }}
.citab {{ width:100%; border-collapse:separate; border-spacing:2.6mm 0; table-layout:fixed; margin-top:1.5mm; }}
.citab th {{ font-size:12pt; font-weight:900; padding:2.6mm 3.2mm; border-radius:2mm 2mm 0 0; color:#fff; text-align:left; }}
.citab th.cig {{ background:#1F7A4D; }} .citab th.cir {{ background:{GAP}; }}
.citab td {{ vertical-align:top; padding:3.6mm 3.6mm; border-radius:0 0 2mm 2mm; border:1pt solid; }}
.citab td.cig {{ background:#F4FBF7; border-color:#3E9A63; }}
.citab td.cir {{ background:#FDF4F3; border-color:#DE9A9A; }}
.cil {{ font-size:8.4pt; font-weight:700; color:{INK}; line-height:1.35; padding:1.4mm 0 1.4mm 3mm; text-indent:-3mm; border-bottom:0.6pt solid #D8DDE4; }}
.citab td .cil:last-child {{ border-bottom:none; }}
.cil:before {{ content:'● '; color:#1F7A4D; }}
.cil.r:before {{ content:'× '; color:{GAP}; }}
.cil b {{ color:{NAVY}; }} .cil.r b {{ color:{GAP}; }}
.cint {{ font-size:8pt; color:{MUT}; font-weight:600; margin-top:3mm; padding-top:2.4mm; border-top:0.6pt solid {LINE}; }}
.sumrow {{ display:table; width:100%; margin-top:3mm; }}
.sumlb {{ display:table-cell; width:24mm; font-size:12pt; font-weight:800; color:#fff; background:{NAVY}; border-radius:2mm; padding:2.5mm 2mm; text-align:center; vertical-align:middle; }}
.sumlb.ok {{ background:{GOOD}; }}
.sumlb.ng {{ background:{GAP}; }}
.sumlb.ci {{ background:{GOLDD}; }}
.sumtx {{ display:table-cell; padding-left:4mm; font-size:13pt; font-weight:600; color:{INK}; vertical-align:middle; line-height:1.5; }}
.sumtx b {{ color:{NAVY}; font-weight:800; }}
.stchip {{ display:inline-block; font-size:13pt; font-weight:800; color:{NAVY}; background:#EAF6EF; border:0.8pt solid {GOOD}; border-radius:2mm; padding:1.2mm 3mm; margin-right:2mm; }}
.ghwrap {{ display:table; width:100%; border-spacing:3mm 0; margin-top:1.2mm; }}
.ghcell {{ display:table-cell; width:50%; vertical-align:top; border:1pt solid {LINE}; border-radius:2mm; padding:1.6mm 2.2mm; }}
.ghttl {{ font-size:8pt; font-weight:800; color:{NAVY}; margin-bottom:0.6mm; }}
.ghnt {{ font-size:6.6pt; font-weight:700; color:{INK}; margin-top:0.6mm; text-align:center; }}
.hz1 {{ height:5mm; background:#1F7A4D; }}
.hz2 {{ height:7mm; background:#8FA8C8; }}
.hz3 {{ height:11mm; background:#C9A15A; }}
.hz4 {{ height:13mm; background:#C0444C; }}
.hz5 {{ height:17mm; background:#8B1A1A; }}
.gen1b {{ height:5mm; background:#2E7D4F; }}
.gen2b {{ height:5mm; background:#8FA8C8; }}
.gen3b {{ height:5mm; background:#8FA8C8; }}
.gen4b {{ height:6mm; background:#C9A15A; }}
.gen5b {{ height:16mm; background:#8B1A1A; }}
.gen2c {{ height:5mm; background:#8FA8C8; }}
.gen3c {{ height:6mm; background:#A9A0C0; }}
.gen4c {{ height:7mm; background:#C9A15A; }}
.gen5c {{ height:11mm; background:{GAP}; }}
.g4b {{ height:6mm; background:#B8C0CC; }}
.gjb {{ height:6mm; background:#C9A15A; }}
.gbb {{ height:10mm; background:{GAP}; }}
.gxb {{ height:18mm; background:#8B1A1A; }}
.gzb {{ height:19mm; background:#5A0F0F; }}
.g4wrap {{ display:table; width:100%; border-spacing:2mm 2mm; }}
.g4cell {{ display:table-cell; width:50%; border:1.2pt solid {LINE}; border-radius:2mm; padding:3mm 3.5mm; vertical-align:top; background:#FBFCFD; }}
.g4tt {{ font-size:14pt; font-weight:800; color:{NAVY}; margin-bottom:1.8mm; }}
.g4nt {{ font-size:9.4pt; font-weight:700; color:{INK}; margin-top:2mm; text-align:center; }}
.q1 {{ height:4mm; background:#9CB8A0; }}
.q2 {{ height:6mm; background:#8FA8C8; }}
.q3 {{ height:9mm; background:#6E90BC; }}
.q4 {{ height:11mm; background:#C9A15A; }}
.q5 {{ height:26mm; background:#8B1A1A; }}
.q5b {{ height:17mm; background:{GAP}; }}
.w1 {{ height:26mm; background:#9CB8A0; }}
.w2 {{ height:15mm; background:#8FA8C8; }}
.w5 {{ height:5mm; background:{GAP}; }}
.extab td, .extab th {{ font-size:7.6pt; line-height:1.25; padding:0.9mm 1.2mm; }}
.extab td.good, .extab .good {{ color:{GOOD}; }}
/* ── MRI 입원 vs 통원 그래프 (얇은형) ── */
.mricap {{ font-size:10pt; font-weight:800; color:{NAVY}; margin:2.4mm 0 1.2mm; }}
.mrig {{ width:92%; border-collapse:separate; border-spacing:3mm 1.2mm; table-layout:fixed; margin:1.2mm auto 0; }}
.mrc {{ width:50%; vertical-align:middle; text-align:left; }}
.mrbw {{ background:#EDF0F4; border-radius:1.2mm; height:6.6mm; overflow:hidden; }}
.mrfill {{ width:58%; margin:0 auto; border-radius:1mm 1mm 0 0; }}
.mrfill.g {{ background:#C9A15A; }} .mrfill.r {{ background:#E29A9A; }}
.mrv {{ font-size:11pt; font-weight:900; margin-top:1mm; }}
.mrv.g {{ color:#9C7C32; }} .mrv.r {{ color:{GAP}; }}
.mrcap {{ font-size:8.4pt; font-weight:800; color:{NAVY}; line-height:1.25; }}
.mrsub {{ font-size:7.4pt; font-weight:700; color:{MUT}; }}
.mrcut {{ font-size:7.4pt; font-weight:800; color:{GAP}; }}
.mrbar {{ height:6.6mm; line-height:6.6mm; border-radius:1.2mm; color:#fff; font-size:9.6pt; font-weight:900; text-align:right; padding-right:2mm; }}
.mrbar.g {{ background:#B98C33; }} .mrbar.r {{ background:{GAP}; }}
.dx3 {{ display:table; width:100%; border-spacing:2.4mm 0; table-layout:fixed; margin-top:2.4mm; }}
.dxc {{ display:table-cell; width:33.33%; border:1pt solid {LINE}; border-radius:2mm; background:#fff; padding:2.4mm 2.6mm; vertical-align:top; }}
.dxh {{ font-size:9.8pt; font-weight:800; color:{NAVY}; margin-bottom:1.4mm; }}
.dxg {{ font-size:9.4pt; font-weight:800; color:#9C7C32; line-height:1.5; }}
.dxr {{ font-size:9.4pt; font-weight:800; color:{GAP}; line-height:1.5; }}
.dxs {{ font-size:7.6pt; font-weight:700; color:{MUT}; margin-top:1mm; }}
.dstab {{ display:table; width:100%; border-spacing:2.6mm 0; table-layout:fixed; margin-bottom:2mm; }}
.dscell {{ display:table-cell; vertical-align:top; border:1pt solid {LINE}; border-radius:2mm; background:#fff; padding:2.4mm 2.6mm; }}
.dsh {{ font-size:11pt; font-weight:900; color:{NAVY}; margin-bottom:1.6mm; }}
.dsh2 {{ font-size:10pt; font-weight:900; color:{NAVY}; margin-bottom:1.6mm; }}
.dsg {{ width:100%; border-collapse:collapse; }}
.dsg th {{ background:{NAVY}; color:#fff; font-size:8.4pt; font-weight:800; padding:1.2mm 1mm; text-align:center; }}
.dsg td {{ border:0.5pt solid {LINE}; font-size:9pt; font-weight:700; color:{INK}; padding:1.6mm 1mm; text-align:center; }}
.dsg td.a {{ font-weight:900; font-size:10pt; }}
.dsg td.a.good {{ color:#1F5FA8; }} .dsg td.a.bad {{ color:{GAP}; }}
.dsg td.bad {{ color:{GAP}; font-weight:900; }}
.dsg tr.hl td {{ background:#FCF3F3; }}
.dsg td.g1 {{ background:#1B3E6F; color:#fff; font-weight:800; }}
.dsg td.g2 {{ background:#2F6BAF; color:#fff; font-weight:800; }}
.dsg td.g3 {{ background:#2E9E96; color:#fff; font-weight:800; }}
.dsg td.g4 {{ background:#6B4FA8; color:#fff; font-weight:800; }}
.dsg td.g5 {{ background:{GAP}; color:#fff; font-weight:800; }}
.dsn {{ font-size:7.8pt; font-weight:700; color:{MUT}; text-align:center; margin-bottom:2.4mm; }}
.dsl {{ font-size:8.6pt; font-weight:700; color:{INK}; line-height:1.4; margin-bottom:1.4mm; }}
.dsl.r {{ color:{GAP}; }}
.dsl b {{ color:{NAVY}; }}
.hjtab {{ width:100%; border-collapse:collapse; margin-top:1.2mm; table-layout:fixed; }}
.hjtab th {{ background:{NAVY}; color:#fff; font-size:8pt; font-weight:800; padding:0.9mm 1mm; text-align:center; border:0.5pt solid #fff; }}
.hjtab td {{ border:0.6pt solid {LINE}; text-align:center; padding:1mm 1mm; }}
.hjtab td.k {{ font-size:8pt; font-weight:700; color:{INK}; background:#F6F8FB; }}
.hjtab td.v {{ font-size:9.6pt; font-weight:900; }}
.hjtab td.v.ok {{ color:{GOOD}; background:#EFF7F2; }} .hjtab td.v.n {{ color:{MUT}; }}
.hjtab td.v.w {{ color:#B5750C; background:#FDF6EA; }} .hjtab td.v.w2 {{ color:#C2531E; background:#FBEDE7; }}
.hjtab td.v.b {{ color:{GAP}; background:#FBEBEB; }}
.hjtag {{ display:inline-block; font-size:7.4pt; font-weight:800; padding:0.5mm 1.4mm; border-radius:1mm; margin:1mm 1mm 0 0; }}
.hjtag.ok {{ background:#E4F0EA; color:{GOOD}; }} .hjtag.n {{ background:#EEF1F5; color:{MUT}; }}
.hjtag.w {{ background:#FBEEDA; color:#B5750C; }} .hjtag.w2 {{ background:#FADFD6; color:#C2531E; }}
.hjtag.b {{ background:#F7DCDC; color:{GAP}; }}
.exn {{ font-size:7.4pt; font-weight:700; color:{INK}; margin-top:1mm; padding:1.2mm 1.8mm; background:#F1F4F8; border-radius:1.5mm; line-height:1.45; }}
.silbox {{ border:1.2pt solid {GOLD}; border-radius:2mm; padding:2mm 2.5mm; margin-top:1.6mm; background:#FBF7EE; }}
.silbox.none {{ border-color:{GAP}; background:#FCF3F3; }}
.sbt {{ font-size:10.5pt; font-weight:800; color:{NAVY}; margin-bottom:1.2mm; }}
.sbt b {{ color:{GOLDD}; }}
.sltab {{ width:100%; border-collapse:collapse; }}
.sltab th {{ background:{NAVY}; color:#fff; font-size:9.4pt; font-weight:800; padding:2.4mm 1.5mm; text-align:center; }}
.sltab td {{ border:0.6pt solid {LINE}; background:#fff; font-size:10pt; font-weight:800; color:{INK}; padding:6.4mm 1.5mm; text-align:center; }}
.sltab td.slp {{ text-align:left; font-size:8.4pt; }}
.sltab td.slc {{ font-weight:800; color:{NAVY}; }}
.sltab td.sla {{ font-weight:800; }}
.sbn {{ font-size:7pt; color:{MUT}; font-weight:600; margin-top:0.6mm; line-height:1.25; }}
.ezbox {{ border:1.2pt solid {LINE}; border-radius:2.5mm; padding:2.5mm 3mm; margin-bottom:2.5mm; background:#FBFCFD; }}
.eztt {{ font-size:13pt; font-weight:800; color:{NAVY}; margin-bottom:2mm; }}
.eznum {{ display:inline-block; width:6mm; height:6mm; line-height:6mm; text-align:center; border-radius:50%; background:{NAVY}; color:#fff; font-size:10pt; margin-right:1.5mm; }}
.eztab {{ display:table; width:100%; border-spacing:2.5mm 0; }}
.ezc {{ display:table-cell; width:50%; border-radius:2mm; padding:2.5mm 3mm; vertical-align:top; border:1pt solid; }}
.ezblue {{ background:#EAF2FB; border-color:#8FA8C8; }}
.ezred {{ background:#FCF3F3; border-color:#E8B4B4; }}
.ezgold {{ background:#FBF7EE; border-color:#C9A15A; }}
.ezpink {{ background:#FCF3F3; border-color:{GAP}; }}
.ezh {{ font-size:15pt; font-weight:800; color:{NAVY}; margin-bottom:1.2mm; }}
.ezsub {{ font-size:8.5pt; color:{MUT}; font-weight:700; }}
.ezbig {{ font-size:13pt; font-weight:800; color:#9C7C32; margin:1mm 0 1.5mm; }}
.ezbig b {{ font-size:22pt; }}
.ezbig.bad {{ color:{GAP}; }}
.ezd {{ font-size:10pt; font-weight:700; color:{INK}; line-height:1.5; }}
.ezd .good {{ color:{GOOD}; }}
.ezd .bad {{ color:{GAP}; }}
.ezn {{ font-size:9.6pt; font-weight:700; color:{INK}; margin-top:2mm; padding:1.8mm 2.5mm; background:#F1F4F8; border-radius:1.6mm; line-height:1.45; }}
.ezhtab {{ display:table; width:100%; }}
.ezhc {{ display:table-cell; text-align:center; vertical-align:middle; }}
.ezarrow {{ display:table-cell; text-align:center; vertical-align:middle; font-size:12pt; font-weight:800; color:{MUT}; width:4mm; }}
.ezhb {{ font-size:15pt; font-weight:800; color:#fff; border-radius:2mm; padding:2.5mm 0; }}
.ezhb.h1 {{ background:#8FA8C8; }}
.ezhb.h2 {{ background:#6E90BC; }}
.ezhb.h3 {{ background:#C9A15A; }}
.ezhb.h4 {{ background:#D08B4A; }}
.ezhb.h5 {{ background:{GAP}; }}
.ezhl {{ font-size:9.5pt; font-weight:800; color:{NAVY}; margin-top:1.2mm; line-height:1.25; }}
.ezsum {{ border:1.5pt solid {NAVY}; border-radius:2.5mm; padding:3mm 3.5mm; background:#F4F7FB; }}
.ezst {{ font-size:12pt; font-weight:800; color:{NAVY}; }}
.ezsl {{ font-size:14pt; font-weight:800; color:{GAP}; margin:1.5mm 0; line-height:1.4; }}
/* ── 급여 통원 5카드(그림형) ── */
.gvcap {{ font-size:11pt; font-weight:800; color:{NAVY}; margin:3mm 0 1.6mm; }}
.gvcap b {{ color:{GOLDD}; }}
.gvtab {{ display:table; width:97%; border-spacing:1.2mm 0; table-layout:fixed; margin:0 auto; }}
.gvc {{ display:table-cell; width:20%; vertical-align:top; border-radius:2mm; border:1.4pt solid; overflow:hidden; }}
.gvc.c1 {{ border-color:#1E9E5A; }} .gvc.c2 {{ border-color:#E8A317; }} .gvc.c3 {{ border-color:#DE7B2B; }}
.gvc.c4 {{ border-color:#D9433F; }} .gvc.c5 {{ border-color:#8A4FBF; }}
.gvh {{ color:#fff; font-size:9.4pt; font-weight:800; text-align:center; padding:2.2mm 0.5mm; line-height:1.2; white-space:nowrap; }}
.gvc.c1 .gvh {{ background:#1E9E5A; }} .gvc.c2 .gvh {{ background:#E8A317; }} .gvc.c3 .gvh {{ background:#DE7B2B; }}
.gvc.c4 .gvh {{ background:#D9433F; }} .gvc.c5 .gvh {{ background:#8A4FBF; }}
.gvsm {{ font-size:7.6pt; font-weight:700; }}
.gvb {{ background:#fff; text-align:center; padding:2mm 0.5mm 2.4mm; }}
.gvk {{ font-size:7.4pt; font-weight:700; color:{MUT}; margin-top:0.8mm; }}
.gvp {{ font-size:17pt; font-weight:900; line-height:1.15; }}
.gvc.c1 .gvp {{ color:#1E9E5A; }} .gvc.c2 .gvp {{ color:#E8A317; }} .gvc.c3 .gvp {{ color:#DE7B2B; }}
.gvc.c4 .gvp {{ color:#D9433F; }} .gvc.c5 .gvp {{ color:#8A4FBF; }}
.gvhr {{ border-top:1.4pt solid #E3E7EC; margin:1.6mm 2mm; }}
.gvm {{ font-size:12.5pt; font-weight:900; color:{GAP}; line-height:1.2; }}
.gvr {{ font-size:12.5pt; font-weight:900; color:{GOOD}; line-height:1.2; }}
.gvtag {{ font-size:7.6pt; font-weight:800; margin-top:1.6mm; white-space:nowrap; }}
.gvtag.ok {{ color:{GOOD}; }} .gvtag.wn {{ color:#C9820A; }} .gvtag.no {{ color:{GAP}; }}
.gvfoot {{ text-align:center; font-size:9pt; font-weight:700; color:{MUT}; margin-top:2.2mm; }}
.gvfoot b {{ color:{INK}; }} .gvfoot .bad {{ color:{GAP}; }}
.ezsn {{ font-size:9.8pt; font-weight:700; color:{INK}; line-height:1.5; }}
/* ── 중증 vs 비중증 (그림형) ── */
.svdef {{ border:1.5pt solid {NAVY}; border-radius:2.5mm; background:#F4F7FB; padding:2.4mm 3mm; margin-bottom:2.6mm; }}
.svdt {{ font-size:12.5pt; font-weight:800; color:{NAVY}; margin-bottom:1mm; }}
.svdt b {{ color:{GOLDD}; }}
.svdd {{ font-size:9.8pt; font-weight:700; color:{INK}; line-height:1.5; }}
.svdd .bad {{ color:{GAP}; }}
.svtab {{ display:table; width:100%; border-spacing:2.6mm 0; margin-bottom:2.6mm; }}
.svc {{ display:table-cell; width:50%; vertical-align:top; border-radius:2.5mm; border:1.6pt solid; overflow:hidden; }}
.svc.g {{ border-color:#C9A15A; }}
.svc.r {{ border-color:{GAP}; }}
.svhd {{ color:#fff; padding:2.2mm 3mm; }}
.svc.g .svhd {{ background:#B98C33; }}
.svc.r .svhd {{ background:{GAP}; }}
.svh1 {{ font-size:15pt; font-weight:800; line-height:1.2; }}
.svh2 {{ font-size:8.8pt; font-weight:700; opacity:0.92; }}
.svbd {{ padding:2.8mm 3mm 3.2mm; }}
.svc.g .svbd {{ background:#FBF7EE; }}
.svc.r .svbd {{ background:#FCF3F3; }}
.svpct {{ font-size:10.5pt; font-weight:800; color:{INK}; margin-bottom:1.6mm; }}
.svpct b {{ font-size:26pt; }}
.svc.g .svpct b {{ color:#9C7C32; }}
.svc.r .svpct b {{ color:{GAP}; }}
.svchips {{ margin-bottom:1.8mm; }}
.svchip {{ display:inline-block; font-size:8.6pt; font-weight:800; padding:0.9mm 2mm; border-radius:1.4mm; margin:0 1mm 1mm 0; border:0.8pt solid; }}
.svc.g .svchip {{ background:#fff; color:#8A6A1E; border-color:#C9A15A; }}
.svc.r .svchip {{ background:#fff; color:{GAP}; border-color:#E8B4B4; }}
.svrow {{ display:table; width:100%; border-top:0.7pt solid #DCD3C2; padding:2.7mm 0; }}
.svc.r .svrow {{ border-top-color:#E8CACA; }}
.svk {{ display:table-cell; width:30%; font-size:8.8pt; font-weight:800; color:{MUT}; vertical-align:middle; }}
.svv {{ display:table-cell; font-size:9.6pt; font-weight:800; color:{INK}; line-height:1.4; vertical-align:middle; }}
.svv .ok {{ color:{GOOD}; }}
.svv .no {{ color:{GAP}; }}
.svex {{ display:table; width:100%; border-spacing:2.2mm 0; margin-bottom:2.4mm; }}
.svec {{ display:table-cell; width:33.33%; border:1pt solid {LINE}; border-radius:2mm; background:#fff; padding:2.8mm 2.6mm; vertical-align:top; }}
.svet {{ font-size:9.6pt; font-weight:800; color:{NAVY}; margin-bottom:1.4mm; }}
.sveg {{ font-size:11pt; font-weight:800; color:#9C7C32; line-height:1.5; }}
.sver {{ font-size:11pt; font-weight:800; color:{GAP}; line-height:1.5; }}
.svesm {{ font-size:8.2pt; font-weight:700; color:{MUT}; margin-top:1mm; line-height:1.35; }}
.cmptab td, .cmptab th {{ font-size:7.2pt; line-height:1.25; padding:0.6mm 1.4mm; }}
.cmptab th.h4g {{ background:{MUT} !important; }}
.cmptab th.h5g {{ background:{NAVY} !important; }}
.cmptab .secrow td {{ background:{NAVY} !important; color:#fff; font-weight:800; font-size:7.8pt; text-align:left; padding:0.8mm 1.8mm; }}
.cmptab .secrow .smn {{ color:#C9D4E2; }}
.cmptab .same {{ display:inline-block; font-size:7.6pt; font-weight:800; color:{GOOD}; background:#EAF6EF; border-radius:1mm; padding:0.3mm 1.2mm; margin-left:1mm; }}
.cmptab td.good, .cmptab .good {{ color:{GOOD}; font-weight:700; }}
.cmppg .memobox {{ height:16mm; }}
.g5tab th.hjung {{ background:#9C7C32 !important; }}
.g5tab th.hbi {{ background:#C0444C !important; }}
.g5tab .hosp {{ display:inline-block; margin-top:0.8mm; font-size:8.4pt; font-weight:800; color:#C0444C; background:#FFF0F0; border:0.6pt solid #E8B4B4; border-radius:1.2mm; padding:0.8mm 1.4mm; line-height:1.35; }}
.hitmk {{ display:inline-block; color:#fff; background:{GAP}; font-weight:800; font-size:8.5pt; margin-left:1.5mm; padding:0.4mm 1.6mm; border-radius:1.2mm; }}
.silverd {{ border:1pt solid {NAVY}; border-left:3.5pt solid {NAVY}; border-radius:2mm; padding:3mm 3.5mm; margin:2.6mm 0; background:#F4F7FB; }}
.silverd .svhead {{ font-size:13pt; font-weight:800; color:{NAVY}; }}
.silverd .svhead b {{ color:{GAP}; font-size:15pt; }}
.silverd .svmeta {{ font-size:9pt; font-weight:600; color:{MUT}; margin-left:2mm; }}
.silverd .svbody {{ font-size:10.5pt; line-height:1.55; color:{INK}; margin-top:1.8mm; }}
.silverd.g1 {{ border-color:{GOOD}; border-left-color:{GOOD}; background:#F1F8F4; }}
.silverd.g5 {{ border-color:{GAP}; border-left-color:{GAP}; background:#FCF3F3; }}
.silverd.chk {{ border-color:{GOLDD}; border-left-color:{GOLDD}; background:#FBF7EE; }}
.stepc .stepn {{ width:5mm; height:5mm; border-radius:50%; background:{NAVY}; color:#fff; font-size:7.5pt; font-weight:800; text-align:center; line-height:5mm; }}
.stepc b {{ display:block; font-size:7.8pt; color:{NAVY}; margin:1mm 0 0.8mm; }}
.wgcols {{ display:flex; gap:3mm; margin-top:1.6mm; }}
.wgcol {{ flex:1; border:0.7pt solid {LINE}; border-radius:2mm; background:#FCFDFE; padding:2.2mm 3mm 3mm; }}
.wgcol .wgh {{ font-size:8.6pt; font-weight:800; color:#2E5A88; margin-bottom:1.6mm; padding-bottom:1.2mm; border-bottom:0.6pt solid {LINE}; }}
.wgcol .wgi {{ font-size:8pt; color:#2B3A52; line-height:1.62; }}
.wgcol .wgi:before {{ content:"· "; color:#2E5A88; font-weight:800; }}
.st {{ width:100%; border-collapse:collapse; margin:1.5mm 0 2mm; font-size:8.4pt; }}
.st th {{ background:{NAVY}; color:#fff; font-weight:800; padding:1.8mm 2mm; text-align:center; font-size:8.2pt; }}
.st td {{ border:0.5pt solid {LINE}; padding:1.6mm 2mm; text-align:center; }}
.st.cmp td {{ padding:1mm 1.6mm; }}
.st.cmp th {{ padding:1.2mm 1.6mm; }}
.st.cmp.mini {{ font-size:7pt; margin:1mm 0 1.5mm; }}
.st.cmp.mini td {{ padding:0.7mm 1.4mm; }}
.st.cmp.mini th {{ padding:0.8mm 1.4mm; font-size:7pt; }}
.st td.g {{ font-weight:800; }}
.st tr.gen1 td {{ background:#EAF2FB; }} .st tr.gen5 td {{ background:#FBEBEC; }}
.st td.bad {{ color:#C0444C; font-weight:800; }} .st td.ok {{ color:#1F7A4D; font-weight:700; }}
.stt {{ font-size:10.5pt; font-weight:800; color:{NAVY}; margin:1mm 0; }}
.stt small {{ font-size:7.6pt; color:{MUT}; font-weight:600; }}
.st.cmp td {{ text-align:center; font-size:8pt; padding:1.4mm 2mm; }}
.st.cmp td.g {{ background:#F2F6FB; font-weight:800; color:{NAVY}; }}
.st.wg th {{ font-size:8pt; padding:1.4mm; }} .st.wg td {{ font-size:8pt; padding:1.2mm 2mm; }}
.st.wg td.n {{ text-align:left; font-weight:700; color:#2B3A52; }}
.wssj.fixnote {{ background:#EEF4FB; border-color:#9BBBD8; }}
.fixnote .cap {{ color:#2E5A88; }}
.fxrow {{ font-size:9pt; line-height:1.5; margin:1.5mm 0; color:#2B3A52; }}
.fxrow b {{ color:#2E5A88; }}
.wbox {{ flex:1; border:0.7pt solid {LINE}; border-radius:1mm; height:24mm; line-height:24mm; padding:0 3mm; font-size:9.5pt; font-weight:800; color:{NAVY}; background:#fff; text-align:right; white-space:nowrap; }}
.wbox.wrap {{ white-space:normal; word-break:keep-all; line-height:1.3; height:auto; min-height:15mm; display:flex; flex-wrap:wrap; align-items:center; justify-content:flex-end; padding:1.4mm 2.6mm; font-size:8pt; }}
.dglist {{ flex:1; }}
.wcard.dgcard {{ min-height:100mm; display:flex; flex-direction:column; }}
.wcard.dgcard .wcf {{ flex:1; align-items:flex-start; }}
.wcard.dgcard .dglist {{ flex:1; align-self:stretch; display:flex; flex-direction:column; justify-content:flex-start; }}
.dgrow {{ display:flex; align-items:center; gap:1mm; margin:0.4mm 0; }}
.dgcancer .dgrow .mb {{ height:7.4mm; line-height:7.4mm; font-size:9pt; }}
.dgcancer .dglab {{ flex:0 0 21mm; font-size:6.2pt; line-height:1.05; white-space:nowrap; }}
.dgcancer .dgrow {{ margin:1.2mm 0; }}
.dgheart .dgrow .mb {{ height:7.4mm; line-height:7.4mm; font-size:9pt; }}
.dgheart .dglab {{ flex:0 0 21mm; font-size:6.2pt; line-height:1.05; white-space:nowrap; }}
.dgheart .dgrow {{ margin:1.2mm 0; }}
.dgrow .dglab {{ flex:0 0 11.5mm; font-size:6.6pt; font-weight:700; color:{NAVY}; }}
.dgrow .mb {{ flex:1; min-width:0; border:0.6pt solid {LINE}; border-radius:1mm; height:8.8mm; line-height:8.8mm; padding:0 1mm; background:#fff; text-align:right; font-size:6.8pt; font-weight:800; color:{NAVY}; white-space:nowrap; }}
.dgrow .dgu {{ flex:0 0 auto; font-size:6pt; color:{MUT}; }}
.dgrow.cirow .dglab {{ color:{NAVY}; font-weight:800; }}
.dgrow .mb.cibox {{ background:#F4F7FB; border-color:{NAVY}; color:{NAVY}; }}
.wunit {{ font-size:8.6pt; color:{MUT}; }}
.wcard.sj {{ min-height:19mm; margin-bottom:0; background:#fff; }}
.wcard.sj .wct {{ font-size:11pt; }}
.wssj .cap {{ font-size:9.4pt; font-weight:800; color:{GOLDD}; padding:0 0 1.2mm; }}



.hcnote {{ background:{NAVY}; color:#EAF0F8; font-size:8.4pt; line-height:1.45; padding:2.4mm 3.5mm; border-radius:2mm; margin-bottom:3mm; }}
.hcnote b {{ color:{GOLDL}; }}
.chnote {{ font-size:7pt; color:{MUT}; line-height:1.4; margin-top:1.5mm; }}
.chnote b {{ color:{NAVY}; }}
.coband {{ border:0.5pt solid {LINE}; border-radius:1.6mm; margin-bottom:2.6mm; overflow:hidden; }}
.colabel {{ background:linear-gradient(135deg,{NAVY},{NAVY2}); color:#fff; font-size:13pt; font-weight:800; padding:1.6mm 3.5mm; }}
.colabel span {{ color:{GOLDL}; font-weight:700; font-size:8pt; margin-left:2.5mm; }}
.grow {{ display:flex; flex-wrap:wrap; padding:1.8mm; gap:1.8mm; align-items:stretch; }}
.gcard {{ flex:1 1 0; min-width:40mm; border:0.4pt solid {LINE}; border-radius:1.2mm; overflow:hidden; }}
.gh {{ font-size:10pt; font-weight:800; padding:1.6mm 2.5mm; color:#fff; }}
.dl {{ font-size:8.6pt; line-height:1.5; padding:1.1mm 2.5mm; border-top:0.3pt solid #EEF1F5; color:{INK}; }}
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
.cov-cell {{ width:50%; vertical-align:top; padding:2mm 5mm 2mm 0; border-bottom:0.5pt solid {LINE}; }}
.cov-h {{ margin-bottom:2mm; }}
.cov-h .cn {{ font-size:12.5pt; font-weight:800; color:{NAVY}; }}
.cov-h .bd {{ font-size:10pt; font-weight:800; padding:0.8mm 3mm; border-radius:8pt; float:right; }}
.items .it {{ display:inline-block; font-size:9.8pt; font-weight:600; padding:1.2mm 2.6mm; margin:0.9mm 1.2mm 0.9mm 0; border-radius:2.5pt; background:#EEF1F5; color:{INK}; }}
.items .it b {{ color:{NAVY}; }}
.items .it.bl b {{ color:{BLUE}; }}
.items .it.r {{ color:{MUT}; background:transparent; border:0.5pt dashed {LINE}; }}
.diag {{ width:100%; border-collapse:collapse; }}
.diag td {{ width:50%; vertical-align:top; padding-right:5mm; }}
.diag td:last-child {{ padding-right:0; padding-left:5mm; }}
.dc {{ border:0.5pt solid {LINE}; border-radius:4pt; overflow:hidden; }}
.dc .h {{ padding:3.5mm 4.5mm; color:#fff; font-weight:800; font-size:12.5pt; }}
.dc.g .h {{ background:{GOOD}; }} .dc.w .h {{ background:{GAP}; }}
.dc ul {{ list-style:none; padding:3.5mm 4.5mm; }}
.dc li {{ font-size:10pt; line-height:1.45; margin-bottom:2mm; padding-left:3.5mm; position:relative; }}
.dc li b {{ color:{NAVY}; }}
.ren {{ width:100%; border-collapse:collapse; margin-top:2mm; }}
.ren td {{ width:50%; vertical-align:top; padding-right:5mm; }}
.ren td:last-child {{ padding-right:0; padding-left:5mm; }}
.rbox {{ border:0.5pt solid {LINE}; border-radius:4pt; overflow:hidden; }}
.rbox .rh {{ padding:2.6mm 4.5mm; color:#fff; font-weight:800; font-size:11.5pt; }}
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
.warn10 {{ font-size:7pt; margin-left:0.8mm; }}
.gdcrit {{ border:0; border-radius:2.6mm; padding:0; margin:0 0 3.4mm; background:linear-gradient(0deg,#FBF8F1,#FBF8F1); border:1.2pt solid {GOLD}; overflow:hidden; }}
.gdhead {{ background:{NAVY}; padding:1.8mm 4mm; }}
.gdtt {{ font-size:11.5pt; font-weight:900; color:#fff; letter-spacing:-0.3px; }}
.gdtt b {{ color:{GOLD}; }}
.gdeng {{ font-size:7pt; font-weight:700; color:#9FB0C6; letter-spacing:3px; margin-top:0.6mm; }}
.gdwrap {{ padding:2.2mm 3mm 0; }}
.gdtab {{ width:100%; border-collapse:collapse; table-layout:fixed; margin:0; }}
.gdtab td {{ padding:0 1.2mm; }}
.gdtab td:first-child {{ padding-left:0; }}
.gdtab td:last-child {{ padding-right:0; }}
.gdtab td {{ width:33.33%; vertical-align:top; text-align:center; padding:0 1.2mm; }}
.gdtab td:first-child {{ padding-left:0; }}
.gdtab td:last-child {{ padding-right:0; }}
.gdn {{ width:6.4mm; height:6.4mm; line-height:6.4mm; margin:0 auto 1.2mm; text-align:center; border-radius:50%; background:{NAVY}; color:#fff; font-size:9.6pt; font-weight:900; }}
.gdk {{ font-size:7.6pt; font-weight:700; color:{MUT}; line-height:1.3; }}
.gdv {{ font-size:9.6pt; font-weight:900; color:{NAVY}; line-height:1.3; margin-top:0.8mm; }}
.gdnt2 {{ font-size:7.4pt; font-weight:700; color:{MUT}; padding:1.8mm 4mm 2mm; }}
.gdnt2 b {{ color:{GAP}; }}
/* ── 보험 인포메이션 간지 (화이트) ── */
.infopg {{ background:#fff; }}
.ibar {{ height:8mm; background:{NAVY}; border-bottom:2mm solid {GOLD}; }}
.ibody {{ padding:34mm 22mm 14mm; height:287mm; position:relative; }}
.ieyebrow {{ font-size:9pt; font-weight:800; color:{GOLDD}; letter-spacing:6px; }}
.ititle2 {{ font-size:50pt; font-weight:900; color:{NAVY}; line-height:1.12; margin-top:4mm; letter-spacing:-1.5px; }}
.ititle2 b {{ color:{GOLDD}; }}
.irule {{ width:32mm; height:1.8mm; background:{GOLD}; margin:7mm 0 6mm; }}
.isub {{ font-size:11.5pt; font-weight:700; color:{MUT}; line-height:1.6; margin-bottom:22mm; }}
.isub b {{ color:{NAVY}; }}
.irow {{ display:table; width:100%; padding:7mm 0; border-top:0.8pt solid #E3E7EC; }}
.irow:last-of-type {{ border-bottom:0.8pt solid #E3E7EC; }}
.inum {{ display:table-cell; width:22mm; vertical-align:middle; font-size:20pt; font-weight:900; color:{GOLD}; letter-spacing:-1px; }}
.itx {{ display:table-cell; vertical-align:middle; }}
.ih {{ font-size:17pt; font-weight:900; color:{NAVY}; line-height:1.25; }}
.ip {{ font-size:9.6pt; font-weight:700; color:{MUT}; margin-top:1.6mm; }}
.ipg {{ display:table-cell; width:26mm; vertical-align:middle; text-align:right; font-size:10pt; font-weight:800; color:{NAVY}; }}
.ifoot {{ position:absolute; left:22mm; right:22mm; bottom:14mm; text-align:center; font-size:8.4pt; font-weight:700; color:{MUT}; padding-top:3mm; border-top:0.6pt solid {LINE}; }}
.flexnote {{ font-size:8pt; font-weight:700; color:{MUT}; background:#F4F6F9; border-left:2.6pt solid {NAVY}; border-radius:1.4mm; padding:1.6mm 2.4mm; margin-top:2.4mm; line-height:1.4; }}
.flexnote b {{ color:{NAVY}; }}
.gdrow {{ font-size:11pt; font-weight:700; color:{INK}; margin:1.3mm 0; }}
.gdrow b {{ color:{NAVY}; }}
.gdtab td b {{ color:{NAVY}; }}
.pbar {{ margin-top:3mm; width:100%; border-collapse:collapse; }}
.pbar td {{ padding:1.2mm 0; vertical-align:middle; }}
.pbar td {{ padding:0.95mm 0; }}
.pbar .bl {{ width:28mm; text-align:right; font-size:10.5pt; font-weight:700; padding-right:3mm; }}
.pbar .track-td {{ width:130mm; }}
.pbar .track {{ height:6mm; background:#EEF1F5; border-radius:3mm; }}
.pbar .fill {{ height:6mm; border-radius:3mm; }}
.pbar .bv {{ width:24mm; text-align:right; font-size:9pt; font-weight:700; padding-left:3mm; }}
.dtab {{ width:100%; border-collapse:collapse; }}
.dcell {{ text-align:center; padding:2.6mm 0; }}
.memo4 {{ border:1.2pt solid {LINE}; border-radius:2mm; padding:3mm 4mm; background:#FCFCFD; }}
.memo4 .ml {{ height:12mm; border-bottom:0.6pt dashed #C9CFD8; }}
.memo4 .ml:last-child {{ border-bottom:none; }}

.dcell .dn {{ font-size:9pt; font-weight:700; color:{NAVY}; margin-top:1mm; }}
.legend {{ display:flex; justify-content:center; gap:12mm; margin:3mm 0 2mm; font-size:11pt; font-weight:700; color:{INK}; }}
.legend span {{ display:flex; align-items:center; gap:2.5mm; white-space:nowrap; padding:1.2mm 3mm; border-radius:3mm; background:#F4F6F9; }}
.legend i {{ width:4.5mm; height:4.5mm; border-radius:1.2mm; display:inline-block; flex:0 0 auto; }}
.sect2 {{ font-size:11.5pt; font-weight:800; color:{NAVY}; margin:3.5mm 0 1.5mm; border-bottom:1.5pt solid {GOLD}; padding-bottom:1.5mm; }}
.sect2 span {{ font-size:8.5pt; font-weight:600; color:{MUT}; letter-spacing:.5px; margin-left:2mm; }}
.btab {{ width:100%; border-collapse:collapse; font-size:9.5pt; }}
.btab th {{ background:{NAVY}; color:#fff; padding:2mm 3mm; text-align:center; font-weight:700; }}
.btab td {{ padding:1.2mm 3mm; text-align:center; border-bottom:.5pt solid {LINE}; color:{INK}; }}
.btab td.bn {{ text-align:left; font-weight:700; color:{NAVY}; }}
.ctab {{ width:100%; border-collapse:separate; border-spacing:2.5mm 0; table-layout:fixed; }}
.ctab .cc {{ background:#F6F8FB; border:0.8pt solid {LINE}; border-top:2.4pt solid {NAVY}; border-radius:2mm; padding:4.4mm 1.5mm; text-align:center; vertical-align:top; }}
.ctab .cl {{ font-size:9.6pt; color:{MUT}; font-weight:800; margin-bottom:2.4mm; }}
.ctab .cval {{ font-size:14pt; }}
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
.ft {{ position:absolute; bottom:0; left:0; width:100%; padding:3mm 11mm; background:{NAVY}; color:#9FB0C6; font-size:8pt; overflow:hidden; }}
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
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · {pgno} / 17</span></div>
</div>'''
    _n8='<div class="hcnote">★ <b>"특정Ⅰ·Ⅱ"는 회사마다 뜻이 다릅니다 — 라벨 말고 질병코드로 확인.</b> 흥국·롯데 특정Ⅰ=급성심근경색 / 한화·NH 특정Ⅰ=협심증·허혈·빈맥·부정맥·심부전 / DB 특정Ⅰ=협심증·허혈·염증 / KB 특정Ⅰ=협심증·허혈·빈맥·심부전 / 현대 특정Ⅰ=빈맥·심부전. 빈맥(I47·48)과 부정맥(I49)은 별개.</div>'
    _n9='<div class="hcnote">★ 삼성·메리츠는 허혈성심장질환을 6가지로 세분(급성기·후속·합병증·협심증·기타급성·만성). 롯데 특정심장Ⅰ=급성심근경색 / 흥국은 특정심혈관질환(기타부정맥제외)=협심·허혈·빈맥·심부전(급성심근 아님). 색: <b style="color:#1F5FA8">허혈·협심</b> / <b style="color:#B9540B">급성심근</b> / <b style="color:#5B7A2E">심근병</b> / <b style="color:#1E7A46">염증</b> / <b style="color:#9A7A12">부정맥·전도</b> / <b style="color:#6A4A9A">판막</b>.</div>'
    _n8b='<div class="hcnote">★ 색: <b style="color:#1F5FA8">허혈·협심</b> / <b style="color:#B9540B">급성심근</b> / <b style="color:#5B7A2E">심근병</b> / <b style="color:#1E7A46">염증</b> / <b style="color:#9A7A12">부정맥·전도</b> / <b style="color:#6A4A9A">판막</b>. 빈맥(I47·48)과 부정맥(I49)은 별개.</div>'
    heart_chart = _fullpage(16,'① 손해보험 (1/4)', ['한화손해보험','DB손해보험'], _n8) + '\n' + \
                  _fullpage(17,'② 손해보험 (2/4)', ['KB손해보험','현대해상'], _n8b) + '\n' + \
                  _fullpage(18,'③ 손해보험 (3/4)', ['NH농협손해보험','삼성화재 (허혈성심장질환)','메리츠화재 (허혈성심장질환)'], _n9) + '\n' + \
                  _fullpage(19,'④ 손해보험 (4/4)', ['흥국화재','롯데손해보험'], _n8b)
    # ★badge-5 담보별: 상단 박스 제거 → 뇌졸증·뇌출혈·급성심근경색 보유금액을 질병코드 표 행 안에 직접 기재(지점장 2026.07.07)
    p5box=''
    _amt_brain={}; _amt_heart={}
    _AMTKEY={'뇌출혈진단비':('b','hem'),'뇌졸증진단비':('b','infarct'),'뇌졸중진단비':('b','infarct'),'급성심근경색':('h','ami'),'뇌혈관진단비':('b','other'),'허혈성 진단비':('h','chronic')}
    for _it in rep.get('p5_own',[]):
        _v=_it.get('v')
        if not _v or _v=='미가입': continue
        _m=_AMTKEY.get(_it.get('t',''))
        if not _m: continue
        (_amt_brain if _m[0]=='b' else _amt_heart)[_m[1]]=_v
    # ★CI 계약(ci status=='ci')일 때만 CI 금액 맵 구성 → 없으면 6p에 CI칸 미표시(유동)
    _cimap = rep.get('ci_amounts') or {}
    if str(rep.get('ci',{}).get('status'))!='ci': _cimap={}
    scv_brain=_scv_build(_BRAIN_TBL,['뇌혈관<br>진단비','순환계','산정<br>특례'],rep.get('scope_brain'),_amt_brain,_cimap)
    scv_heart=_scv_build(_HEART_TBL,['허혈성<br>진단비','심장<br>(특정)','순환계','산정<br>특례'],rep.get('scope_heart'),_amt_heart,_cimap)
    # ★2026.07.11 실손 세대 자동판별(CI식) → 검출 세대 강조 표 + 세대별 맞춤 화법
    _sg=rep.get('silson_gen',{'status':'none'})
    # ★2026.07.11 지점장 확정 세분화: 2세대 3분할 / 1세대 생보·손보 구분(상해의료비)
    _GENROWS=[('1','1세대 (손보)','~ 2009.09','갱신 3·5년<br>재가입 없음','자기부담금 없음 · 상해의료비 별도 담보<br><span class="smn">※ 2009.07~09월 가입은 회사별로 1·2세대가 다를 수 있음</span>','손보'),
              ('1','1세대 (생보)','~ 2009.09','갱신 3·5년<br>재가입 없음','자기부담 20% · 상해의료비 포함형<br><span class="smn">※ 2009.07~09월 가입은 회사별로 1·2세대가 다를 수 있음</span>','생보'),
              ('2','2-1세대','2009.10 ~ 2012.12','갱신 3년<br>재가입 없음','자기부담 10% · 급여90/비급여80~90%','2-1'),
              ('2','2-2세대','2013.01 ~ 2015.12','갱신 1년<br><b>재가입 15년</b>','자기부담 10~20% 선택','2-2'),
              ('2','2-3세대','2016.01 ~ 2017.03','갱신 1년<br>재가입 15년','정신질환 급여 보상 · 응급실 비응급 면책','2-3'),
              ('3','3세대 (착한실손)','2017.04 ~ 2021.06','갱신 1년<br>재가입 15년','급여90/비급여80% · 3대 비급여 70%<br><span class="smn">도수 350만 · 주사 250만 · MRI 300만</span>','' ),
              ('4','4세대','2021.07 ~ 2026.05.05','갱신 1년<br><b>재가입 5년</b>','급여80/비급여70% · 비급여 할증<br><span class="smn">자기부담금 한도 연 200만(급여·입원만)</span>','' ),
              ('5','5세대','2026.05.06 ~','갱신 1년<br>재가입 5년','급여 통원 건보율 연동 · 비급여 중증/비중증 분리<br><span class="smn">도수·체외·주사제는 <b>비중증(특약2)만</b> 보장 제외</span>','' )]
    _hg=_sg.get('gen') if _sg.get('status')=='auto' else None
    _hs=_sg.get('sub','') if _sg.get('status')=='auto' else ''
    _rws=[]
    for gk,lbl,per,ren,feat,sub in _GENROWS:
        base='gen1' if gk=='1' else ('gen5' if gk=='5' else '')
        hit = (_hg is not None and str(_hg)==gk and (sub=='' or sub==_hs))
        cls=(base+' genhit').strip() if hit else base
        pcls=' class="bad"' if gk=='5' else ' class="g"'
        fcls=' class="bad"' if gk=='5' else ''
        star='<span class="hitmk">✓ 가입</span>' if hit else ''
        _rws.append(f'<tr class="{cls}"><td class="g">{lbl}{star}</td><td{pcls}>{per}</td><td class="rencol">{ren}</td><td{fcls}>{feat}</td></tr>')
    _gentab=('<table class="st gentab"><tr><th style="width:18%">세대</th><th style="width:18%">판매시기 (가입일)</th>'
             '<th style="width:14%">갱신 · 재가입</th><th style="width:52%">핵심 특징</th></tr>'+''.join(_rws)+'</table>')
    _GEN_TALK={
     '1손보':('1세대 손보 실손 — 가장 두터운 조건입니다. 자기부담금 없음.','★절대 해지 금지 — 재가입 불가. 상해의료비가 별도 담보로 붙어 있는지 증권 확인.','도수·체외충격파 본인부담 거의 없음'),
     '1생보':('1세대 생보 실손 — 자기부담 20%, 상해의료비 포함형입니다.','손보 1세대(0%)와 부담률이 다릅니다. 유지가 유리하나 자기부담 20% 안내 필요.','도수·체외충격파 본인부담 소액'),
     '2-1':('2-1세대(~2012.12) — 표준화 초기, 자기부담 10%.','재가입 주기가 없습니다. 유지가 유리합니다.','도수 약 4천원 · 체외충격파 약 2만원 — 보장 가능'),
     '2-2':('2-2세대(2013.01~2015.12) — 자기부담 10~20% 선택형.','★재가입 15년 주기가 시작된 구간입니다. 재가입 시점 확인 필요.','도수 약 4천원 · 체외충격파 약 2만원 — 보장 가능'),
     '2-3':('2-3세대(2016.01~2017.03) — 정신질환 급여 보상 개시.','응급실 비응급 내원은 면책입니다. 재가입 15년 유지.','도수 약 4천원 · 체외충격파 약 2만원 — 보장 가능'),
     '3':('3세대 착한실손 — 급여 10%/비급여 20%.','도수·비급여주사·MRI가 특약으로 분리됐습니다. 특약 가입 여부 확인.','도수 4~8천원 · 체외충격파 약 3만원 — 연 50회 한도'),
     '4':('4세대 — 급여 20%/비급여 30%, 비급여 할증 구조.','비급여를 많이 쓰면 보험료가 할증됩니다.','도수 약 8천원 · 체외충격파 약 3만원 — 보장 가능'),
     '5':('★5세대 — 비급여 대폭 축소.','기존 세대(특히 1·2세대) 보유 중이면 유지가 훨씬 유리합니다.','★도수 본인부담 약 4.2만원(95%) · 체외충격파 보장 제외')}
    def _talkkey(g,sub):
        if g==1: return '1생보' if sub=='생보' else '1손보'
        if g==2: return sub if sub in ('2-1','2-2','2-3') else '2-2'
        return str(g)
    # ★실손 계약 결론 박스 (어느 계약에 실손이 있는지)
    _sl=rep.get('silson_list',[])
    if _sl:
        _rows=''.join(
            f'<tr><td class="slc">{_html.escape(x["co"])}</td>'
            f'<td class="slp">{_html.escape(x["prod"])}</td>'
            f'<td class="sld">{_html.escape(x["join"])}</td>'
            f'<td class="sla">{x["amt"]:,}원</td></tr>' for x in _sl)
        _silbox=('<div class="silbox"><div class="sbt">■ 고객님의 <b>실손보험은 여기</b> 들어 있습니다</div>'
                 '<table class="sltab"><tr><th>보험사</th><th>상품명</th><th>가입일</th><th>월 보험료</th></tr>'
                 f'{_rows}</table>'
                 f'<div class="sbn">위 계약의 <b>가입일 기준</b>으로 세대를 판별했습니다. 실손이 여러 개면 <b>가장 오래된 계약</b>이 기준입니다.</div></div>')
    else:
        _silbox=('<div class="silbox none"><div class="sbt">■ 실손보험 <b>미보유</b></div>'
                 '<div class="sbn">보유 계약에서 실손 담보(입원·통원·약값)가 확인되지 않습니다. 실손 가입 검토가 필요합니다.</div></div>')
    if _sg.get('status')=='auto':
        _g=_sg['gen']; _sub=_sg.get('sub',''); _t=_GEN_TALK[_talkkey(_g,_sub)]
        _lbl=(f"{_g}세대 ({_sub})" if (_g==1 and _sub) else (_sub+'세대' if _g==2 and _sub else f'{_g}세대'))
        _silverd=(f'<div class="silverd g{_g}"><div class="svhead">● 고객님 실손 = <b>{_html.escape(_lbl)}</b>'
                  f'<span class="svmeta">{_html.escape(str(_sg.get("company","")))} · 가입 {_html.escape(str(_sg.get("date","")))} (자동 판별)</span></div>'
                  f'<div class="svbody"><b>{_t[0]}</b><br>{_t[1]}<br><span class="svtx">■ {_t[2]}</span></div></div>')
    elif _sg.get('status')=='check':
        _silverd=('<div class="silverd chk"><div class="svhead">■ 실손 세대 — 확인 필요'
                  f'<span class="svmeta">{_html.escape(str(_sg.get("company","")))} 실손 보유 · 가입일 미확정</span></div>'
                  '<div class="svbody">증권에서 <b>실손 가입일</b>을 확인하면 세대가 자동 결정됩니다. 아래 칸에 가입일을 기입하세요.</div></div>')
    else:
        _silverd=('<div class="silverd none"><div class="svhead">ℹ 실손보험 미보유</div>'
                  '<div class="svbody">현재 실손 담보가 확인되지 않습니다. 실손 가입 검토를 권유드립니다.</div></div>')
    _md = '.' if _PPT_MODE else ''
    memo_lines = ''.join('<div class="ml">%s</div>' % _md for _ in range(10))
    doc=f'''<!DOCTYPE html><html><head><meta charset="utf-8"><style>{css}</style></head><body>
<!-- P0: 표지 -->
<div class="pg cvpg">
 <div class="cvbar"></div>
 <div class="cvmark">MAKEONE</div>
 <div class="cvbody">
  <div class="cvbrand">MAKEONE<div class="ln"></div></div>
  <div class="cveyebrow">COVERAGE DIAGNOSIS REPORT</div>
  <div class="cvtitle">보장 진단서</div>
  <div class="cvrule"></div>
  <div class="cvsub">{rep.get('meta','')}</div>
  <div class="cvnamebox"><span class="g">{cust}</span><span class="s">고객님</span></div>
  <div class="cvstats">
   <div class="cvst"><div class="k">보유 계약</div><div class="v">{rep.get('n_contract',0)} <small>건</small></div></div>
   <div class="cvst"><div class="k">월 납입보험료</div><div class="v">{rep.get('premium',0):,} <small>원</small></div></div>
   <div class="cvst"><div class="k">갱신 / 비갱신</div><div class="v">{rep.get('renew',0)} <small>/ {rep.get('nonrenew',0)}</small></div></div>
  </div>
  <div class="cvspacer"></div>
  <div class="cvhr2"></div>
  <div class="cvfoot">MAKEONE · 보장분석 자동화 리포트</div>
 </div>
 <div class="cvfootbar"></div>
</div>
<!-- P1 -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>1</b>보장 현황</div><div class="bar"></div></div>
 <div class="body">
  <div class="gdcrit">
   <div class="gdhead">
    <div class="gdtt">좋은 보험 <b>가입 기준</b></div>
    <div class="gdeng">GOOD INSURANCE CRITERIA</div>
   </div>
   <div class="gdwrap"><table class="gdtab"><tr>
    <td><div class="gdcard"><div class="gdn">1</div><div class="gdk">보험료 대비</div><div class="gdv">담보의 다양성</div></div></td>
    <td><div class="gdcard"><div class="gdn">2</div><div class="gdk">보상이 잘 되는 담보</div><div class="gdv">제대로 가입됐는지</div></div></td>
    <td><div class="gdcard"><div class="gdn">3</div><div class="gdk">은퇴 · 노후</div><div class="gdv">준비되어 있는지</div></div></td>
   </tr></table></div>
   <div class="gdnt2">※ 월 보험료 <b>10만원 이상</b> = 아래 표에서 <b>빨간색 ▲</b>로 표시됩니다.</div>
  </div>

  <table class="meta"><tr>
   <td><div class="k">보유 계약</div><div class="v">{n_contract}<small> 건</small></div></td>
   <td><div class="k">월 납입보험료</div><div class="v">{premium:,}<small> 원</small></div></td>
   <td><div class="k">갱신 / 비갱신</div><div class="v">{renew}<small> / {nonrenew}</small></div></td>
   <td><div class="k">보장 공백</div><div class="v" style="color:{GAP}">{gap_cnt}<small> 영역</small></div></td>
  </tr></table>
  <div class="sect">보장 현황 <span>CATEGORY COVERAGE</span></div>
  <table class="cov">{rows}</table>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 1 / 17</span></div>
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
  <div class="sect" style="margin-top:4.5mm">갱신 / 비갱신 구조 <span>RENEWAL</span></div>
  <table class="ren"><tr>
   <td><div class="rbox b"><div class="rh">갱신형 {renew}건<small>보험료 인상 가능</small></div>{prem_rows(rep["renew_list"],True)}</div></td>
   <td><div class="rbox k"><div class="rh">비갱신형 {nonrenew}건<small>만기까지 고정</small></div>{prem_rows(rep["nonrenew_list"],False)}</div></td>
  </tr></table>
  <div class="sect" style="margin-top:4.5mm">월 보험료 구성 <span>PREMIUM</span></div>
  <table class="pbar">{bars}</table>
  <div class="flexnote">※ 이 페이지는 <b>계약 건수에 따라 표 길이·막대 개수가 유동적으로 바뀌는 페이지</b>입니다. 계약이 많아지면 갱신/비갱신 목록과 보험료 막대가 늘어나 배치가 자동 조정됩니다.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 2 / 17</span></div>
</div>
<!-- P3: 핵심 보장 분석 (CI 선지급 + 주요 치료비) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>3</b>핵심 보장 분석</div><div class="bar"></div></div>
 <div class="body">
  {ci_html}{noci_html}
  <div class="sect"{' style="margin-top:5mm"' if (ci_html or noci_html) else ''}>중대질병(CI) 진단 담보 <span>CRITICAL ILLNESS BENEFITS</span></div>
  <table class="ctab">{crows}</table>
  {comment_html}
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 3 / 17</span></div>
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
    <span><i style="background:#1F7A4D"></i>충실 70%↑</span>
    <span><i style="background:#D08B1F"></i>보강권장 40–69%</span>
    <span><i style="background:#C0242E"></i>취약 40%↓</span>
  </div>
  <div class="sect2">충족률 산정 근거 <span>보유 ÷ {band} 권장</span></div>
  <table class="btab">
    <tr><th>영역</th><th>보유</th><th>권장({band})</th><th>충족률</th></tr>
    {brows}
  </table>
  <div class="note">※ <b>충족률 = 보유 ÷ 연령밴드 권장액 × 100</b> (상한 100%). 권장액은 업계 적정 가입금액 가이드(암 진단비 5천만~1억·뇌혈관 3천만~5천만·허혈성 심장 3천만 등) 기준이며 {band} 표준밴드를 적용했습니다. 운전자·실손·일당·응급실은 핵심담보 보유개수 기준입니다. 개인 소득·가족력에 따라 권장액은 상담을 통해 조정됩니다.{age_warn}</div>
  {advice_html}
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 4 / 17</span></div>
</div>
<!-- P12b: 실손 세대 구분 (단독) -->
<div class="pg silgenpg">
 <div class="top"><div class="eb">BARUM 보장분석 · 실손 세대 확인</div>
  <div class="nm">실손보험 <b>세대 구분</b> — 고객님은 몇 세대이신가요?</div>
  <div class="pgn"><b>5</b>실손 세대 구분</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="wscap n2 wsectcap gap2">■ 실손 세대 구분 — 판매시기(가입일) 기준</div>
  {_gentab}
  <div class="note">※ 세대 판별은 <b>가입일</b> 기준입니다. 표준화 실손은 2009.10월부터 판매되어, 본 진단서는 <b>2009.09월까지 1세대 / 2009.10월부터 2세대</b>로 적용합니다.</div>
  {_silverd}
  {_silbox}
  <div class="renote">
   <div class="rentt">▶ 재가입 주기 — 약관이 바뀌는 시점</div>
   <div class="renrow"><b>2013.01 ~</b> &nbsp;재가입 <b>15년</b></div>
   <div class="renrow"><b>2021.07 ~</b> &nbsp;재가입 <b>5년</b> <span class="smn">(4세대)</span></div>
   <div class="renrow"><b>2026.05.06 ~</b> &nbsp;재가입 <b>5년</b> <span class="smn">(5세대)</span></div>
   <div class="rennt">재가입 시점이 오면 <b>그때 팔고 있는 상품 약관으로 바뀐다.</b> 2013.01 이전 가입은 재가입 조건이 없어 기존 약관 그대로 유지된다.</div>
  </div>

 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 5 / 17</span></div>
</div>
<!-- P5: 담보별 보장범위 (최종본 260707 스펙) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>6</b>담보별 보장범위</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">담보별 보장범위 — 질병코드 커버 <span>DISEASE-CODE COVERAGE · 각 축=개별 담보·각각 보상 · 산정특례·순환계·외상성뇌출혈=단독</span></div>
  {p5box}
  <div class="scv2">
   <div class="scvcol">
    <div class="scvhd brain">뇌 — 질병코드별 커버</div>
    {scv_brain}
    <div class="scvleg"><span class="on">●</span> 보장 &nbsp;<span class="off">○</span> 미보장 &nbsp;<span class="hold">[확인]</span> 산정특례 HOLD &nbsp;<span class="own2">노란행=보유</span></div>
    <div class="scvnote">· {cust} 보유 = <b>뇌출혈진단비(I60~62)</b> → <b class="r">뇌경색(I63)부터 공백</b>, 순환계·2대주요치료비로 확장 필요<br>· <b>외상성 뇌출혈(S06)</b> = 뇌혈관진단비 <b class="r">미보장</b>(산정특례 축only)<br>· <b>산정특례</b> = 진단 기반 별개 담보축. 대상 코드범위 = 뇌혈관질환 전체 I60~69 + Q28 선천 + S06(개별 담보·각각 보상). 지급조건은 회사·약관별 [확인]</div>
    <div class="rngbox">
     <div class="rngt">■ 뇌혈관 보장범위 <span class="smn">(발병률 · 심평원)</span></div>
     <table class="rngtab"><tr>
      <td class="rgcol">
       <div class="rngc">
        <div class="rgo o1"><span class="rgp">100%</span></div>
        <div class="rgo o2"><span class="rgp">60.8%</span></div>
        <div class="rgo o3"><span class="rgp">9.3%</span></div>
       </div>
      </td>
      <td class="rgtxt">
       <div class="rgl"><span class="d d3"></span><b>뇌출혈</b> <span class="cdx">I60·61·62</span> <b class="pc">9.3%</b></div>
       <div class="rgl"><span class="d d2"></span><b>뇌졸중</b> <span class="cdx">+뇌경색 I63·65·66</span> <b class="pc">60.8%</b></div>
       <div class="rgl"><span class="d d1"></span><b>뇌혈관질환</b> <span class="cdx">전체 I60~69</span> <b class="pc">100%</b></div>
      </td>
     </tr></table>
    </div>
   </div>
   <div class="scvcol">
    <div class="scvhd heart">심장 — 질병코드별 커버</div>
    {scv_heart}
    <div class="scvleg"><span class="on">●</span> 보장 &nbsp;<span class="off">○</span> 미보장 &nbsp;<span class="hold">[확인]</span> 산정특례 HOLD &nbsp;<span class="own2">노란행=보유</span></div>
    <div class="scvnote">· {cust} 보유 = <b>허혈성진단비 + 부정맥(I49)</b> → 허혈성(I20~25)·부정맥 커버, <b class="r">판막·염증·심부전은 심장특정/순환계 영역</b><br>· <b>빈맥(I47·48)</b> = 마스터 무행·전 묶음 <b class="r">제외</b> → 본 표 미기재<br>· <b>산정특례</b> = 진단 기반 별개 담보축. 대상 코드범위 = 심혈관질환 전체 I20~50 + 판막(개별 담보·각각 보상). 지급조건 [확인]<br>· <b>순환계</b> = 가장 넓은 범위(2대 + 동맥류·정맥류 등 순환기 전반). 세부 대상코드 회사·약관별 상이 [확인]</div>
    <div class="rngbox">
     <div class="rngt">■ 심장 보장범위 <span class="smn">(발병률 · 심평원)</span></div>
     <table class="rngtab"><tr>
      <td class="rgcol">
       <div class="rngc">
        <div class="rgo o1 h"><span class="rgp">100%</span></div>
        <div class="rgo o2 h"><span class="rgp">50%</span></div>
        <div class="rgo o3 h"><span class="rgp">9%</span></div>
       </div>
      </td>
      <td class="rgtxt">
       <div class="rgl"><span class="d d3h"></span><b>급성심근경색</b> <span class="cdx">I21·22·23</span> <b class="pc r">9%</b></div>
       <div class="rgl"><span class="d d2h"></span><b>허혈성심장</b> <span class="cdx">+협심증 I20·24·25</span> <b class="pc r">50%</b></div>
       <div class="rgl"><span class="d d1h"></span><b>심장특정질환</b> <span class="cdx">판막·염증·부정맥·심부전</span> <b class="pc r">100%</b></div>
      </td>
     </tr></table>
    </div>

   </div>
  </div>
  <div class="scvbot"><div class="h">산정특례 기준 (진단 기반 · 별개 담보축)</div>· 산정특례 = 위 4범위(허혈성·2대·순환계)와 <b>축이 다른 별개 담보</b> — 마스터 '산정특례심장'·'산정특례뇌혈관' 전용행에서 진단코드 기반으로 지급. &nbsp;· 외상성 뇌출혈(S06) = 뇌혈관진단비 미보장 → <b>산정특례 축only</b>로만 커버(고정사실). &nbsp;· 대상 코드범위(확정): <b>[뇌]</b> I60~69 전체 + I67.0·1·5·6 + Q28 선천 + S06 / <b>[심]</b> I20~25·I30~41·I42~45·I46·I47~50 + 판막. <b>각각 개별 담보로 각각 보상.</b> 지급조건·기간(30일·5% 등)만 회사·약관별 [확인].</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 6 / 17</span></div>
</div>

<!-- P7: 상담 워크시트 (FINAL版·3단 중앙이름+산정특례) -->
<div class="pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 상담 워크시트</div>
  <div class="nm">지금 고객의 <b>3대 주요치료비</b>는?</div>
  <div class="pgn"><b>7</b>상담 워크시트</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="ws3">
   <div class="wscol wsmain">
    <div class="wscap c">● 암 주요치료비</div>
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
    <div class="wscap h">● 뇌·심장 주요치료비</div>
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
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 7 / 17</span></div>
</div>

<!-- P8: 바뀌지 않는 담보 — 비갱신 추천 -->
<div class="pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 평생 지키는 준비</div>
  <div class="nm">바뀌지 않는 담보 <b>— 은퇴 후에도 평생</b></div>
  <div class="pgn"><b>8</b>비갱신 추천</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="ws3">
   <div class="wscol wsmain">
    <div class="wscap n2">■ 사망 · 진단비 · 수술</div>
    {_wcard_fix_list('사망','평생 유지 · 종신형',['종신 사망','질병 사망','상해 사망'])}
    {_wcard_fix_group('진단비','암·뇌·심 진단 일시금',[('■ 암',['고액암','일반암','유사암']),('■ 뇌',['뇌혈관','뇌졸증','뇌출혈']),('■ 심장',['허혈성','협심증','급성심근','심부전','염증','부정맥'])])}
    {_wcard_fix_group('수술','질병 · 상해 구분',[('■ 질병',['질병수술','1~5종','N대수술','뇌혈관','허혈성','심장','암수술']),('■ 상해',['상해수술','1~5종','중대상해','창상봉합','골절수술','5대골절','화상수술'])])}
   </div>
   <div class="wsmid">
    <div class="lb">OUR CLIENT</div>
    <div class="nm">{"<br>".join(list(cust)[:4])}</div>
    <div class="nmsub">고객님</div>
    <div class="cnt">평생<small>비갱신 고정</small></div>
   </div>
   <div class="wscol wsmain">
    <div class="wscap n2">■ 후유장해 · 상해 · 일당</div>
    {_wcard_fix_list('후유장해','상해·질병 각 3% · 80%',['상해 후유 3%','상해 후유 80%','질병 후유 3%','질병 후유 80%'])}
    {_wcard_fix_list('상해','골절·화상·깁스·응급실',['골절','5대 골절','화상진단비','중대화상진단비','반깁스','깁스','응급실'])}
    {_wcard_fix_list('일당','입원·수술·중환자실',['질병 입원일당','상해 입원일당','질병 수술일당','상해 수술일당','질병 중환자실','상해 중환자실','1인실(종합/상급)'])}
   </div>
  </div>
  <div class="wscap n2 wsectcap gap2">■ 연금 · 종신 · 저축 — 은퇴 후 평생 소득</div>
  <div class="ws2">
   {_wcard_fix_list('연금','달러 · 노후 소득',['달러연금','달러종신'])}
   {_wcard_fix_list('종신 · 저축','평생 보장 · 목돈 마련',['종신','저축보장'])}
  </div>
  <div class="wssj fixnote">
   <div class="cap">비갱신 3대 이유</div>
   <div class="fxrow"><b>① 평생 고정</b> 납입 끝나면 보험료 0원 · <b>② 납입면제</b> 중대질환 진단 시 이후 면제 · <b>③ 은퇴 전 확정</b> 나이 들수록 발생률↑, 미리 잠근다.</div>
  </div>
  <div class="note24">● 바뀌지 않는 담보는 <b>비갱신</b>으로 가입하셔서 <b>은퇴 후</b>를 준비해야 합니다.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 8 / 17</span></div>
</div>
<!-- P10: 운전자 · 간병 (병합) -->
<div class="pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 평생 지키는 준비</div>
  <div class="nm">운전자 · 간병 <b>— 일상 리스크 대비</b></div>
  <div class="pgn"><b>9</b>운전자·간병</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="wscap n2 wsectcap">■ 운전자보험 담보</div>
  <div class="ws2">
   {_wcard_fix_list('벌금 · 합의금','교통사고 형사·행정',['대인 벌금','대물 벌금','합의금','6주미만 합의금'])}
   {_wcard_fix_list('변호사 · 위로금','기타 지원',['변호사비','자동차부상위로금','영업용 운전면허취소지원'])}
  </div>
  <table class="st cmp mini"><tr><th style="width:18%">구분</th><th>자동차보험 <span style="font-weight:400">(의무)</span></th><th>운전자보험 <span style="font-weight:400">(선택)</span></th></tr>
   <tr><td class="g">보장 대상</td><td>타인의 피해</td><td class="bad">운전자 본인</td></tr>
   <tr><td class="g">책임 종류</td><td>민사 배상</td><td class="bad">형사·행정</td></tr>
   <tr><td class="g">주요 보장</td><td>대인 · 대물</td><td class="bad">벌금 · 형사합의금 · 변호사선임비</td></tr></table>
  <div class="wscap n2 wsectcap gap2">■ 간병비 담보</div>
  <div class="ws2">
   {_wcard_fix_list('간병인 지원','보험사 파견',['간병인지원일당','간호통합병동'])}
   {_wcard_fix_list('간병인 사용','직접 고용',['간병인사용일당','간병인사용일당(체증형)','요양병원 간호통합병동'])}
  </div>
  <table class="st cmp"><tr><th style="width:20%">구분</th><th>간병인지원일당</th><th>간병인사용일당</th></tr>
   <tr><td class="g">방식</td><td>보험사가 간병인 <b>직접 배정</b></td><td>내가 <b>직접 고용</b> 후 정액 지급</td></tr>
   <tr><td class="g">간병인 선택</td><td>불가 (배정 · 교체 가능)</td><td class="ok"><b>간병인 or 지인</b> 가능</td></tr>
   <tr><td class="g">갱신</td><td>5 · 10 · 15 · 20년 갱신</td><td class="ok">비갱신 선택 가능</td></tr>
   <tr><td class="g">납입면제</td><td>갱신형이라 <b>갱신 끝나면 다시 납부</b></td><td class="ok"><b>납입면제 가능</b></td></tr>
   <tr><td class="g">간호통합병동</td><td>가입금액 보장 (정액)</td><td>가입금액 보장 (정액)</td></tr>
   <tr><td class="g">인건비 상승</td><td>간병인 배정 (교체 가능)</td><td>체증형 (5년 10%↑)</td></tr></table>
  <div class="wscap n2 wsectcap gap2">■ 간병인사용일당 — 이렇게 하면 됩니다 (쉬운 6단계)</div>
  <div class="steprow big">
   <div class="stepc"><div class="stepn">1</div><b>간병인 부르기</b>입원하면 <b>바로</b> 간병협회·등록업체·앱(예: 케어네이션)에 연락해 간병인을 부른다. <b>늦으면 소급이 안 되니 입원 첫날에!</b></div>
   <div class="stepc"><div class="stepn">2</div><b>간병 받기</b>입원 동안 간병인이 식사·세면·체위변경·이동을 돕는다. 가족·지인도 가능(단 <b>24시간 같이 생활</b>).</div>
   <div class="stepc"><div class="stepn">3</div><b>서류 챙기기</b>퇴원 전에 병원서 <b>입퇴원확인서</b>, 업체서 <b>간병사용확인서</b>, 그리고 <b>간병비 영수증</b>을 받는다.</div>
   <div class="stepc"><div class="stepn">4</div><b>간병비 내기</b>간병비를 <b>실제로 지급</b>하고 증빙을 남긴다. 계좌이체 메모에 "○○○ 간병비"라 적으면 깔끔.</div>
   <div class="stepc"><div class="stepn">5</div><b>청구서류 모으기</b>입원확인서 + 간병사용확인서 + 지급내역(이체·카드) + <b>업체 사업자등록증</b>을 한 번에.</div>
   <div class="stepc"><div class="stepn">6</div><b>보험사에 청구</b>모은 서류를 보험사에 내면 <b>심사 후 간병비가 지급</b>된다. 끝!</div>
  </div>
  <div class="note24">● <b>간병인 사용일당</b> 사용 시 환자와 <b>24시간 동행</b>해야 인정됩니다.</div>
  </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 9 / 17</span></div>
</div>
<!-- P12: 재가보험 -->
<div class="pg jgbig">
 <div class="top"><div class="eb">BARUM 보장분석 · 평생 지키는 준비</div>
  <div class="nm">재가보험 <b>— 장기요양 · 노후 돌봄</b></div>
  <div class="pgn"><b>10</b>재가보험</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="wscap n2 wsectcap">■ 재가보험 · 장기요양</div>
  <div class="ws2">
   {_wcard_fix_list('장기요양 등급','등급 판정 · 인지지원',['1등급','2등급','3등급','4등급','5등급','인지지원등급'])}
   {_wcard_fix_list('복합재가 (통합재가)','한 기관·한 계약 · 5종 복합',['주야간보호 (노치원)','방문요양','방문목욕','방문간호','단기보호'])}
  </div>
  <div class="wscap n2 wsectcap gap2">■ 복지용구 18품목 — 품목별 한도가 다릅니다 (연 160만원 · 본인부담 0~15%)</div>
  <div class="ws2">
   <div class="wscol">
    <table class="st wg"><tr><th style="width:44%">구입품목 (10)</th><th>내구연한</th><th>개수 한도</th></tr>
     <tr><td class="n">이동변기</td><td>5년</td><td>1개</td></tr>
     <tr><td class="n">목욕의자</td><td>5년</td><td>1개</td></tr>
     <tr><td class="n">성인용보행기</td><td>5년</td><td class="ok">2개</td></tr>
     <tr><td class="n">욕창예방방석</td><td>3년</td><td>1개</td></tr>
     <tr><td class="n">지팡이</td><td>2년</td><td>1개</td></tr>
     <tr><td class="n">자세변환용구</td><td>-</td><td>연 5개</td></tr>
     <tr><td class="n">간이변기</td><td>-</td><td>연 2개</td></tr>
     <tr><td class="n">안전손잡이</td><td>-</td><td>연 10개</td></tr>
     <tr><td class="n">미끄럼방지(양말)</td><td>-</td><td>연 6켤레</td></tr>
     <tr><td class="n">요실금팬티</td><td>-</td><td>연 4개</td></tr></table>
   </div>
   <div class="wscol">
    <table class="st wg"><tr><th style="width:52%">대여품목 (6)</th><th>내구연한</th></tr>
     <tr><td class="n">수동휠체어</td><td>5년</td></tr>
     <tr><td class="n">전동침대</td><td>10년</td></tr>
     <tr><td class="n">수동침대</td><td>10년</td></tr>
     <tr><td class="n">이동욕조</td><td>5년</td></tr>
     <tr><td class="n">목욕리프트</td><td>3년</td></tr>
     <tr><td class="n">배회감지기</td><td>5년</td></tr></table>
    <table class="st wg" style="margin-top:2mm"><tr><th style="width:52%">구입 또는 대여 (2)</th><th>내구연한</th></tr>
     <tr><td class="n">욕창예방매트리스</td><td>3년</td></tr>
     <tr><td class="n">경사로</td><td>실내 2년(6개)/실외 8년</td></tr></table>
   </div>
  </div>
  <div class="wssj fixnote">
   <div class="cap">신인 설계사 · 이것만 기억하세요</div>
   <div class="fxrow"><b>한도</b> 등급 무관 <b>연 160만원</b>(구입+대여 합산) · 본인부담 0~15%(기초수급 0%) · 초과분은 전액 본인부담 · 미사용 잔액 이월 안 됨(유효기간 1년).</div>
   <div class="fxrow"><b>품목별 규칙</b> 내구연한 있는 품목은 그 기간 내 1개(성인용보행기 2·경사로 실내 6 예외) · 내구연한 없는 소모품은 연 개수 한도 · <b>시설 입소 시 전면 불가</b>, 의료기관 15일↑ 입원 중 침대·이동욕조·목욕리프트 대여 불가.</div>
   <div class="fxrow"><b>상담 포인트</b> 재가등급(1~5·인지지원)만 받으면 누구나 이용 · "급여확인서"에 사용가능 품목이 찍혀 나옴 · 온·오프라인 복지용구 사업소에서 구입/대여. <b>시설급여</b>(요양원 입소)·<b>특별현금급여</b>(가족요양비)는 재가·복지용구와 별개 급여.</div>
  </div>
  <div class="wstalk2"><b>■ 노치원(주야간보호)</b> — 어르신을 <b>낮 동안 시설에서 돌봄</b>(식사·목욕·재활·프로그램)하고 저녁엔 집으로 귀가. 가족은 낮에 생업·휴식이 가능하다. 복합재가 5종 중 하나로, 등급만 있으면 이용.</div>
  <div class="wstalk2"><b>■ 장기요양등급, 언제 신청?</b> — <b>만 65세 이상</b>은 소득 무관 누구나 신청. <b>65세 미만이라도</b> 치매·뇌혈관질환·파킨슨병 등 <b>노인성 질병</b>이 있으면 <b>의사소견서(진단서)</b> 첨부해 <b>미리 신청 가능</b>. 건강보험공단(■1577-1000)·앱·홈페이지 신청 → 방문조사 → 약 30일 내 등급 판정. 65세 전에 준비해 두면 은퇴 후 돌봄 공백을 막는다.</div>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 10 / 17</span></div>
<!-- P-INFO: 보험 인포메이션 간지 -->
<div class="pg infopg">
 <div class="ibar"></div>
 <div class="ibody">
  <div class="ieyebrow">INSURANCE INFORMATION</div>
  <div class="ititle2">보험 <b>인포메이션</b></div>
  <div class="irule"></div>
  <div class="isub">고객 계약과 무관한 <b>공통 참고 자료</b>입니다.</div>

  <div class="irow">
   <div class="inum">01</div>
   <div class="itx"><div class="ih">5세대 실손보험</div><div class="ip">급여 · 비급여 · 도수 · 체외충격파</div></div>
   <div class="ipg">11 – 14</div>
  </div>
  <div class="irow">
   <div class="inum">02</div>
   <div class="itx"><div class="ih">3대 주요치료비 변천사</div><div class="ip">비례형 → 정액형 → 비급여</div></div>
   <div class="ipg">15</div>
  </div>
  <div class="irow">
   <div class="inum">03</div>
   <div class="itx"><div class="ih">심혈관 담보 분류</div><div class="ip">보험사별 · 질병코드 기준</div></div>
   <div class="ipg">16 – 19</div>
  </div>
  <div class="irow">
   <div class="inum">04</div>
   <div class="itx"><div class="ih">회사별 비교표</div><div class="ip">11개 비교표 · 설계사 참고용</div></div>
   <div class="ipg">부1 – 부3</div>
  </div>

  <div class="ifoot">MAKEONE · 보장분석 자동화 리포트 &nbsp;|&nbsp; 자료기준일 2026.06.30 · 정리 BARUM지점 최은혜</div>
 </div>
</div>
<!-- P5d: 5세대 쉽게 이해하기 -->
<div class="pg ez5pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 5세대 실손</div>
  <div class="nm">5세대 실손 <b>— 쉽게 딱 3가지만</b></div>
  <div class="pgn"><b>11</b>5세대 쉽게</div><div class="bar"></div></div>
 <div class="body sbody">

  <div class="ezbox ez1">
   <div class="eztt"><span class="eznum">1</span> 병원비는 <b>두 종류</b>다</div>
   <table class="eztab">
    <tr>
     <td class="ezc ezblue">
      <div class="ezh">급여</div>
      <div class="ezd">건강보험이 <b>도와주는</b> 치료<br><span class="smn">감기 · 수술 · 입원 등 대부분</span></div>
     </td>
     <td class="ezc ezred">
      <div class="ezh">비급여</div>
      <div class="ezd">건강보험이 <b>안 도와주는</b> 치료<br><span class="smn">도수치료 · MRI · 영양주사 등</span></div>
     </td>
    </tr>
   </table>
   <div class="ezn">5세대는 이 둘을 <b>따로따로</b> 계산한다. 예전엔 뭉뚱그렸다.</div>
  </div>

  <div class="ezbox ez2">
   <div class="eztt"><span class="eznum">2</span> 급여 통원 — <b>큰 병원 갈수록 내 돈이 커진다</b></div>
   <div class="ezhosp">
    <table class="ezhtab">
     <tr>
      <td class="ezhc"><div class="ezhb h1">30%</div><div class="ezhl">의원<br><span class="smn">동네병원</span></div></td>
      <td class="ezarrow">→</td>
      <td class="ezhc"><div class="ezhb h2">40%</div><div class="ezhl">병원</div></td>
      <td class="ezarrow">→</td>
      <td class="ezhc"><div class="ezhb h3">50%</div><div class="ezhl">종합병원</div></td>
      <td class="ezarrow">→</td>
      <td class="ezhc"><div class="ezhb h4">60%</div><div class="ezhl">대학병원</div></td>
      <td class="ezarrow">→</td>
      <td class="ezhc"><div class="ezhb h5">90%</div><div class="ezhl">응급실<br><span class="smn">경증</span></div></td>
     </tr>
    </table>
   </div>
   <div class="ezn"><b>통원(외래)만</b> 해당. <b>입원은 어느 병원이든 20%</b>로 똑같다.<br>※ 4세대는 병원 상관없이 20%였다 → 5세대는 <b class="bad">대학병원 가면 60%를 내가 낸다.</b></div>

   <div class="gvcap">■ 내가 낸 <b>급여 본인부담금 10만원</b> 기준 — 직접 계산해보면</div>
   <table class="gvtab">
    <tr>
     <td class="gvc c1"><div class="gvh">동네 의원</div><div class="gvb">
       <div class="gvk">건보 부담률</div><div class="gvp">30%</div><div class="gvhr"></div>
       <div class="gvk">내 부담</div><div class="gvm">3만원</div>
       <div class="gvk">돌려받는 돈</div><div class="gvr">7만원</div>
       <div class="gvtag ok">✓ 가장 유리</div></div></td>
     <td class="gvc c2"><div class="gvh">병원</div><div class="gvb">
       <div class="gvk">건보 부담률</div><div class="gvp">40%</div><div class="gvhr"></div>
       <div class="gvk">내 부담</div><div class="gvm">4만원</div>
       <div class="gvk">돌려받는 돈</div><div class="gvr">6만원</div>
       <div class="gvtag ok">✓ 무난</div></div></td>
     <td class="gvc c3"><div class="gvh">종합병원</div><div class="gvb">
       <div class="gvk">건보 부담률</div><div class="gvp">50%</div><div class="gvhr"></div>
       <div class="gvk">내 부담</div><div class="gvm">5만원</div>
       <div class="gvk">돌려받는 돈</div><div class="gvr">5만원</div>
       <div class="gvtag wn">※ 절반만</div></div></td>
     <td class="gvc c4"><div class="gvh">대학병원</div><div class="gvb">
       <div class="gvk">건보 부담률</div><div class="gvp">60%</div><div class="gvhr"></div>
       <div class="gvk">내 부담</div><div class="gvm">6만원</div>
       <div class="gvk">돌려받는 돈</div><div class="gvr">4만원</div>
       <div class="gvtag wn">※ 주의</div></div></td>
     <td class="gvc c5"><div class="gvh">권역응급 <span class="gvsm">(경증)</span></div><div class="gvb">
       <div class="gvk">건보 부담률</div><div class="gvp">90%</div><div class="gvhr"></div>
       <div class="gvk">내 부담</div><div class="gvm">9만원</div>
       <div class="gvk">돌려받는 돈</div><div class="gvr">1만원</div>
       <div class="gvtag no">× 거의 없음</div></div></td>
    </tr>
   </table>
   <div class="gvfoot"><b>4세대였다면?</b> 어느 병원이든 내 부담 <b>2만원</b> → 돌려받는 돈 <b>8만원</b> &nbsp;|&nbsp; <span class="bad">× 권역응급 경증: 청구해도 1만원만!</span></div>
  </div>

  <div class="ezsum">
   <div class="ezst">★ 한 줄로 외우세요</div>
   <div class="ezsl"><b>"급여는 큰 병원 갈수록 내 돈이 커진다. 비급여는 큰 병이냐 아니냐로 갈린다."</b></div>
   <div class="ezsn">급여(건보 적용) = <b>병원 규모로 갈린다</b> — 동네의원 30%, 대학병원 60%, 응급실 경증 90%.<br>비급여 = <b>중증(30%) vs 비중증(50%)</b>으로 갈린다 → <b>다음 장에서 상세히 본다.</b></div>
  </div>

 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 11 / 17</span></div>
</div>
<!-- P5f: 중증 vs 비중증 (그림형) -->
<div class="pg ez5pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 5세대 실손</div>
  <div class="nm">중증 vs 비중증 <b>— 기준은 '산정특례'다</b></div>
  <div class="pgn"><b>12</b>중증·비중증</div><div class="bar"></div></div>
 <div class="body sbody">

  <div class="svdef">
   <div class="svdt">■ 중증의 기준 = <b>건강보험 산정특례 대상자</b></div>
   <div class="svdd">보험사가 정하는 게 아니다. <b>국민건강보험공단에 산정특례로 등록된 질환이면 '중증(특약1)'</b>, 아니면 전부 '비중증(특약2)'이다.<br>
   산정특례 = 암·뇌혈관·심장·희귀난치질환 등 큰 병에 대해 <b>건보 본인부담을 5~10%로 낮춰주는 제도</b>. 진단 후 공단에 등록하면 산정특례 등록증이 나온다.<br>
   <span class="bad">※ 같은 도수치료라도 — 암 치료 목적이면 중증(보장), 허리 통증이면 비중증(0원). 병명이 아니라 <b>산정특례 등록 여부</b>가 갈림길이다.</span></div>
  </div>

  <div class="svtab">
   <div class="svc g">
    <div class="svhd">
     <div class="svh1">특약1 · 중증 비급여</div>
     <div class="svh2">산정특례 등록 질환 + 인과관계 분명한 합병증</div>
    </div>
    <div class="svbd">
     <div class="svpct">내 돈 <b>30%</b> &nbsp;<span class="smn">입원 기준</span></div>
     <div class="svchips">
      <span class="svchip">암</span><span class="svchip">뇌혈관</span><span class="svchip">심장</span><span class="svchip">희귀·중증난치</span><span class="svchip">중증화상</span><span class="svchip">중증외상</span><span class="svchip">결핵</span><span class="svchip">중증치매</span>
     </div>
     <div class="svrow"><div class="svk">자기부담</div><div class="svv">입원 30% / 통원 Max[30%, 3만]</div></div>
     <div class="svrow"><div class="svk">한도</div><div class="svv">연 <b>5,000만</b> · 통원 회당 20만<br><span class="smn">입원 한도 없음</span></div></div>
     <div class="svrow"><div class="svk">상한</div><div class="svv"><span class="ok">★ 상급종합·종합 입원 시<br>연 자기부담 500만 초과분 <b>전액 보장</b></span></div></div>
     <div class="svrow"><div class="svk">도수·체외<br>비급여주사</div><div class="svv"><span class="ok">모두 보장 ✓</span> <span class="smn">면책 없음</span></div></div>
     <div class="svrow"><div class="svk">보험료 할증</div><div class="svv"><span class="ok">없음</span> <span class="smn">아무리 써도 그대로</span></div></div>
    </div>
   </div>
   <div class="svc r">
    <div class="svhd">
     <div class="svh1">특약2 · 비중증 비급여</div>
     <div class="svh2">그 외 전부 — 허리·어깨·감기·미용 아닌 일반 비급여</div>
    </div>
    <div class="svbd">
     <div class="svpct">내 돈 <b>50%</b> &nbsp;<span class="smn">입원 기준</span></div>
     <div class="svchips">
      <span class="svchip">디스크·요통</span><span class="svchip">어깨·관절</span><span class="svchip">일반 수술</span><span class="svchip">비중증 MRI</span><span class="svchip">영양·주사</span><span class="svchip">도수·체외</span>
     </div>
     <div class="svrow"><div class="svk">자기부담</div><div class="svv">입원 50% / 통원 Max[50%, 5만]</div></div>
     <div class="svrow"><div class="svk">한도</div><div class="svv">연 <span class="no">1,000만</span> · 입원 회당 300만<br><span class="smn">통원 1일 20만</span></div></div>
     <div class="svrow"><div class="svk">상한</div><div class="svv"><span class="no">없음</span> <span class="smn">초과분은 전부 본인 부담</span></div></div>
     <div class="svrow"><div class="svk">도수·체외<br>비급여주사</div><div class="svv"><span class="no">보장 제외 ×</span> <span class="smn">+ 미등재 신의료기술</span></div></div>
     <div class="svrow"><div class="svk">보험료 할증</div><div class="svv"><span class="no">최대 +300%</span> <span class="smn">비급여 보험료 4배</span></div></div>
    </div>
   </div>
  </div>

  <div class="svex">
   <div class="svec">
    <div class="svet">■ 비급여 <b>100만원</b> 냈다면</div>
    <div class="sveg">중증 → <b>70만원</b> 환급</div>
    <div class="sver">비중증 → <b>50만원</b> 환급</div>
    <div class="svesm">같은 100만원인데 20만원 차이. 비중증은 소액일수록 더 불리하다(4만원 청구 = 0원).</div>
   </div>
   <div class="svec">
    <div class="svet">■ <b>도수치료</b> 100만원</div>
    <div class="sveg">중증(암 치료) → <b>70만원</b></div>
    <div class="sver">비중증(요통) → <b>0원</b></div>
    <div class="svesm">2026.07.01부터 도수는 관리급여(1회 43,850원·본인 95%). 체외충격파·비급여주사는 비중증이면 그대로 0원.</div>
   </div>
   <div class="svec">
    <div class="svet">■ <b>MRI</b> 50만원</div>
    <div class="sveg">입원 → 중증 <b>35만</b> / 비중증 <b>25만</b></div>
    <div class="sver">통원 → <b>둘 다 20만</b> (한도컷)</div>
    <div class="svesm">자기부담 = 중증 Max[30%,3만] / 비중증 Max[50%,5만]. <b>통원은 회당 20만 한도</b>에 잘려 중증·비중증이 같아진다.</div>
   </div>
  </div>

  <div class="ezsum">
   <div class="ezst">★ 상담 한 줄</div>
   <div class="ezsl"><b>"산정특례 등록증이 나오는 병이면 5세대도 괜찮습니다. 아니면 절반만 나옵니다."</b></div>
   <div class="ezsn">암·뇌·심장은 오히려 <b>500만원 상한</b>이 생겨 강화됐다. 문제는 <b class="bad">도수·체외충격파·영양주사</b> — 비중증이면 한 푼도 안 나온다. 이 치료를 정기적으로 받는 분은 5세대 전환 금지.</div>
  </div>

 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 12 / 17</span></div>
</div>
<!-- P5e: 5세대 도표 + 예시 -->
<div class="pg ez5pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 5세대 실손</div>
  <div class="nm">5세대는 <b>얼마나 불리해지나</b> — 한눈에</div>
  <div class="pgn"><b>13</b>5세대 도표</div><div class="bar"></div></div>
 <div class="body sbody">

  <table class="g4wrap">
   <tr>
    <td class="g4cell">
     <div class="g4tt">① 급여 <span class="smn">통원 자기부담률</span></div>
     <div class="barwrap">
      <div class="barcol"><div class="hbar q1"><span>0</span></div><div class="barlb vt">1세대</div></div>
      <div class="barcol"><div class="hbar q2"><span>10</span></div><div class="barlb vt">2세대</div></div>
      <div class="barcol"><div class="hbar q2"><span>10</span></div><div class="barlb vt">3세대</div></div>
      <div class="barcol"><div class="hbar q3"><span>20</span></div><div class="barlb vt">4세대</div></div>
      <div class="barcol"><div class="hbar q5"><span>90</span></div><div class="barlb vt bad">5세대</div></div>
     </div>
     <div class="g4nt">4세대까지 <b>20%</b> → 5세대 <b class="bad">최대 90%</b><span class="smn">(응급실)</span></div>
    </td>
    <td class="g4cell">
     <div class="g4tt">② 비급여 <span class="smn">자기부담률</span></div>
     <div class="barwrap">
      <div class="barcol"><div class="hbar q1"><span>0</span></div><div class="barlb vt">1세대</div></div>
      <div class="barcol"><div class="hbar q3"><span>20</span></div><div class="barlb vt">2세대</div></div>
      <div class="barcol"><div class="hbar q4"><span>30</span></div><div class="barlb vt">3세대</div></div>
      <div class="barcol"><div class="hbar q4"><span>30</span></div><div class="barlb vt">4세대</div></div>
      <div class="barcol"><div class="hbar q5b"><span>50</span></div><div class="barlb vt bad">5세대</div></div>
     </div>
     <div class="g4nt">비중증 <b class="bad">50%</b> · 도수 <b class="bad">0원</b><span class="smn">(면책)</span></div>
    </td>
   </tr>
   <tr>
    <td class="g4cell">
     <div class="g4tt">③ 할증 제도 <span class="smn">비급여 보험료</span></div>
     <div class="barwrap">
      <div class="barcol"><div class="hbar q1"><span>無</span></div><div class="barlb vt">1세대</div></div>
      <div class="barcol"><div class="hbar q1"><span>無</span></div><div class="barlb vt">2세대</div></div>
      <div class="barcol"><div class="hbar q1"><span>無</span></div><div class="barlb vt">3세대</div></div>
      <div class="barcol"><div class="hbar q5"><span>300</span></div><div class="barlb vt">4세대</div></div>
      <div class="barcol"><div class="hbar q5"><span>300</span></div><div class="barlb vt bad">5세대</div></div>
     </div>
     <div class="g4nt">1~3세대 <b>할증 없음</b> → 4·5세대 <b class="bad">최대 +300%</b></div>
    </td>
    <td class="g4cell">
     <div class="g4tt">④ 보장 한도 <span class="smn">비급여 연간</span></div>
     <div class="barwrap">
      <div class="barcol"><div class="hbar w1"><span>1억</span></div><div class="barlb vt">1세대</div></div>
      <div class="barcol"><div class="hbar w2"><span>5천</span></div><div class="barlb vt">2세대</div></div>
      <div class="barcol"><div class="hbar w2"><span>5천</span></div><div class="barlb vt">3세대</div></div>
      <div class="barcol"><div class="hbar w2"><span>5천</span></div><div class="barlb vt">4세대</div></div>
      <div class="barcol"><div class="hbar w5"><span>1천</span></div><div class="barlb vt bad">5세대</div></div>
     </div>
     <div class="g4nt">비중증 한도 <b class="bad">5천만 → 1천만</b> (1/5)</div>
    </td>
   </tr>
  </table>

  <div class="wscap n2 wsectcap gap2">■ 비급여 보험료 할증 — <b>직전 1년 청구액에 따라 5단계</b></div>
  <table class="hjtab">
   <tr><th>1등급</th><th>2등급</th><th>3등급</th><th>4등급</th><th>5등급</th></tr>
   <tr><td class="k">보험금 <b>0원</b></td><td class="k">100만 미만</td><td class="k">150만 미만</td><td class="k">300만 미만</td><td class="k">300만 이상</td></tr>
   <tr><td class="v ok">-5% 할인</td><td class="v n">유지</td><td class="v w">+100%</td><td class="v w2">+200%</td><td class="v b">+300% (4배)</td></tr>
  </table>
  <div class="exn">※ 할증은 <b>비중증(특약2) 비급여</b> 보험료에만 붙는다. <b class="good">급여·중증(특약1)은 아무리 써도 할증 없음.</b> 직전 2년 특약2 보험금 미수령 시 전체보험료 <b>10% 추가 할인</b>.</div>
  <div class="wscap n2 wsectcap gap2">■ 실제로 얼마나 받나 (5세대) — <b>MRI 50만원 · 입원 vs 통원</b></div>
  <table class="mrig">
   <tr>
    <td class="mrc"><div class="mrbw"><div class="mrbar g" style="width:57%">20만</div></div><div class="mrcap">중증 · 통원 <span class="mrcut">회당 20만 한도컷</span></div></td>
    <td class="mrc"><div class="mrbw"><div class="mrbar g" style="width:100%">35만</div></div><div class="mrcap">중증 · 입원 <span class="mrsub">한도 없음</span></div></td>
   </tr>
   <tr>
    <td class="mrc"><div class="mrbw"><div class="mrbar r" style="width:57%">20만</div></div><div class="mrcap">비중증 · 통원 <span class="mrcut">일당 20만 한도컷</span></div></td>
    <td class="mrc"><div class="mrbw"><div class="mrbar r" style="width:71%">25만</div></div><div class="mrcap">비중증 · 입원 <span class="mrsub">회당 300만 한도</span></div></td>
   </tr>
  </table>
 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 13 / 17</span></div>
</div>
<!-- P5g: 도수·체외충격파 (세대별 부담) -->
<div class="pg ez5pg">
 <div class="top"><div class="eb">BARUM 보장분석 · 5세대 실손</div>
  <div class="nm">도수치료 · 체외충격파 <b>— 2026년 7월부터 달라집니다</b></div>
  <div class="pgn"><b>14</b>도수·체외</div><div class="bar"></div></div>
 <div class="body sbody">

  <table class="dstab">
   <tr>
    <td class="dscell">
     <div class="dsh">■ 도수치료 <span class="smn">1회 43,850원 기준(예시)</span></div>
     <table class="dsg">
      <tr><th>실손 세대</th><th>예상 본인부담</th><th>보험 적용</th></tr>
      <tr><td class="g1">1세대</td><td class="a good">0원</td><td>보장 가능</td></tr>
      <tr><td class="g2">2세대</td><td class="a">약 4천원</td><td>보장 가능</td></tr>
      <tr><td class="g3">3세대</td><td class="a">약 4천~8천원</td><td>보장 가능</td></tr>
      <tr><td class="g4">4세대</td><td class="a">약 8천원</td><td>보장 가능</td></tr>
      <tr class="hl"><td class="g5">5세대</td><td class="a bad">약 4만 2천원</td><td class="bad">본인부담 95%</td></tr>
     </table>
    </td>
    <td class="dscell">
     <div class="dsh">■ 체외충격파 <span class="smn">1회 10만원 기준(예시)</span></div>
     <table class="dsg">
      <tr><th>실손 세대</th><th>예상 본인부담</th><th>보험 적용</th></tr>
      <tr><td class="g1">1세대</td><td class="a good">0원</td><td>보장 가능</td></tr>
      <tr><td class="g2">2세대</td><td class="a">약 2만원</td><td>보장 가능</td></tr>
      <tr><td class="g3">3세대</td><td class="a">약 3만원</td><td>보장 가능</td></tr>
      <tr><td class="g4">4세대</td><td class="a">약 3만원</td><td>보장 가능</td></tr>
      <tr class="hl"><td class="g5">5세대</td><td class="a bad">10만원 전액</td><td class="bad">보장 제외</td></tr>
     </table>
    </td>
   </tr>
  </table>
  <div class="dsn">※ 위 금액은 예시이며, 실제 지급액은 보험사·약관·질환·치료기록 등에 따라 달라질 수 있습니다. ※ 5세대는 <b>중증(특약1)이면 보장</b> — 위 표는 비중증(특약2) 기준.</div>

  <table class="dstab">
   <tr>
    <td class="dscell">
     <div class="dsh2">■ 도수치료 — 이용 제한</div>
     <div class="dsl">✓ 정부 <b>관리급여 전환</b> (2026.07 시행)</div>
     <div class="dsl">✓ <b>연 15회 제한</b> (주 2회 이내)</div>
     <div class="dsl">✓ 전국 동일 수가 적용 (1회 43,850원)</div>
     <div class="dsl">✓ 수술·골절 후 재활 시 최대 <b>연 24회</b> 인정</div>
     <div class="dsl">✓ 선행 물리치료 2주·4회 후 호전 없어야 인정</div>
    </td>
    <td class="dscell">
     <div class="dsh2">■ 체외충격파 — 이용 제한</div>
     <div class="dsl">✓ 비급여 유지 (자율 가이드라인)</div>
     <div class="dsl">✓ <b>연 12회 제한</b> · 부위당 최대 6회</div>
     <div class="dsl">✓ 주 1회 권장, 1회 2,000타 이상 권장</div>
     <div class="dsl">✓ 동일 회차 다부위 치료 제한</div>
     <div class="dsl r">✓ 5세대 실손(비중증) = <b>보장 대상 제외</b></div>
    </td>
    <td class="dscell">
     <div class="dsh2">■ 체외충격파 인정 부위 (7개)</div>
     <div class="dsl">● <b>어깨</b> <span class="smn">석회성건염·회전근개</span></div>
     <div class="dsl">● <b>팔꿈치</b> <span class="smn">테니스엘보·골프엘보</span></div>
     <div class="dsl">● <b>고관절</b> <span class="smn">대전자통증증후군</span></div>
     <div class="dsl">● <b>무릎</b> <span class="smn">슬개건염</span></div>
     <div class="dsl">● <b>발목</b> <span class="smn">아킬레스건염</span></div>
     <div class="dsl">● <b>발바닥</b> <span class="smn">족저근막염</span></div>
     <div class="dsl">● <b>척추</b> <span class="smn">경추·요추 근막통증증후군</span></div>
    </td>
   </tr>
  </table>

  <table class="dx3">
   <tr>
    <td class="dxc"><div class="dxh">도수치료 <span class="smn">1회 43,850원</span></div>
      <div class="dxg">중증 → <b>보장</b> (30%·최소3만)</div>
      <div class="dxr">비중증 → <b>관리급여 95%</b> · 내 돈 약 4만 2천원</div>
      <div class="dxs">2026.07.01 시행 · 연 15회 · 주 2회</div></td>
    <td class="dxc"><div class="dxh">체외충격파 <span class="smn">1회 10만원</span></div>
      <div class="dxg">중증 → <b>보장</b> (30%·최소3만)</div>
      <div class="dxr">비중증 → <b>보장 제외 · 0원</b></div>
      <div class="dxs">전액 본인부담 (10만원 그대로)</div></td>
    <td class="dxc"><div class="dxh">비급여 주사 <span class="smn">영양주사 등</span></div>
      <div class="dxg">중증 → <b>보장</b> (30%·최소3만)</div>
      <div class="dxr">비중증 → <b>보장 제외 · 0원</b></div>
      <div class="dxs">입원해도 안 나온다</div></td>
   </tr>
  </table>
  <div class="exn">★ <b>같은 MRI 50만원인데 입원이냐 통원이냐로 돌려받는 돈이 달라진다.</b> 통원은 <b class="bad">회당 20만 한도</b>에 잘려 중증·비중증이 똑같이 20만. 입원이면 중증 35만 / 비중증 25만.<br>※ <b>중증</b> = 암·뇌혈관·심장·희귀난치(산정특례) 질환 <b>치료 목적</b>일 때만. 허리·어깨 통증으로 받는 도수치료는 <b class="bad">비중증 → 0원</b>이다.<br>※ MRI 자기부담 = 중증 Max[30%, 3만] / 비중증 Max[50%, 5만]. <b class="bad">통원은 20만 한도에 걸려 잘린다.</b></div>


  <div class="ezsum">
   <div class="ezst">★ 상담 첫 질문</div>
   <div class="ezsl"><b>"고객님, 실손 몇 세대세요?"</b></div>
   <div class="ezsn">실손 세대에 따라 보장 결과가 완전히 달라진다. ① 실손보험 몇 세대인가? ② 올해 치료를 몇 번 받았는가? ③ 인정 질환·인정 부위인가?</div>
  </div>

 </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 14 / 17</span></div>
</div>
<!-- P6: 주요치료비 변천사 (juyo_a4_v7 상세본) -->
<div class="pg">
 <div class="top"><div class="eb">MAKEONE · 보장분석 리포트</div>
  <div class="nm">{cust} <b>고객님</b> 보장 진단서</div>
  <div class="pgn"><b>15</b>주요치료비 변천사</div><div class="bar"></div></div>
 <div class="body sbody">
  <div class="sect">3대 주요치료비 담보의 변천사 <span>①비례형(구간) → ②정액형 → ③비급여 → 생활비 · 정액형 가입금액 100만원부터</span></div>
  <div class="jc">
   <div class="jcol">
    <div class="jhd">● 암 주요치료비 <span class="sub">4세대 진화</span></div>
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
    <div class="jwarn"><div class="h">※ 실손으론 부족 — 5세대 기준</div>· 비중증 비급여: 자기부담 50%·통원 회당 5만원·한도 <b>1,000만</b>(4세대 5천만→축소)<br>· 중증(암·산정특례): 자기부담 30%·연 5천만이나 표적·면역 고가 약값엔 부족<br>· 외래(통원) 주사 반복 → <b>주요치료비·하이클래스가 메움</b></div>
    <div class="jtalk"><span class="h">■ 상담 화법</span> &nbsp;"표적항암 한 달 수백만 원이 전액 본인부담이에요. 5세대 실손은 비급여가 1천만으로 줄었습니다. 구간형은 이제 못 드는 담보라 유지가 답이고, 없으면 정액형·하이클래스로 채워야 합니다."</div>
   </div>
   <div class="jcol">
    <div class="jhd">● 뇌·심장 주요치료비 <span class="sub">암과 동일 진화</span></div>
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
    <div class="jwarn"><div class="h">▶ 한 번으로 안 끝난다</div>뇌경색 재발 1년 10%·<b>5년 20~30%</b> / 급성심근경색 <b>1년 30%·3년 50%</b> / 스텐트 5년 재협착 15%</div>
    <table class="jt"><tr><th>수술</th><th>건당 총진료비(평균)</th></tr>
     <tr><td>심장수술(개흉)</td><td class="p">2,832만원</td></tr>
     <tr><td>관상동맥우회술(CABG)</td><td class="p">2,690만원</td></tr>
     <tr><td>뇌동맥류 코일색전술</td><td class="p">1,100~1,600만원</td></tr>
     <tr><td>심장 스텐트(다발혈관)</td><td class="p">1,000~1,400만원</td></tr>
     <tr><td>뇌기저부수술</td><td class="p">1,475만원</td></tr></table>
    <div class="jnote2">총진료비 기준(심평원·건보공단 통계) · 재발 시 이 비용이 반복 — 주요치료비가 채운다</div>
   </div>
  </div>
 <div class="ft"><b>MAKEONE</b> 보장분석 자동화<span class="r">{cust} 고객님 · 15 / 17</span></div>
</div>
</div>
</div>
{heart_chart}
{_ga_html}
</body></html>'''
    if _PPT_MODE:
        import re as _re
        doc = _re.sub(r'<span class="mb([^"]*)"></span>', r'<span class="mb\1">.</span>', doc)
        doc = _re.sub(r'<span class="wbox([^"]*)"></span>', r'<span class="wbox\1">.</span>', doc)
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
