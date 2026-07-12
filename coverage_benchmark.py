# ===== BARUM coverage_benchmark.py v33-ci-rate-20260708 (선지급률=CI계약열 직접산출·50/80 고정) =====
# -*- coding: utf-8 -*-
"""
BARUM 충족률 엔진 + map_excel_to_report
- 입력: 완성된 보장진단 엑셀(.xlsx) 1개 (등식2: 리포트는 완성 엑셀만 읽음)
- 출력: report_weasy.build_report_pdf 가 먹는 rep dict
- 충족률 % = 보유합 / 연령밴드별 권장액 * 100 (상한 없음·실제치, 2026.07.12 지점장 확정)
나이/성별 미추출 시 기본밴드 '40s' 적용 + gap_count/배지에 [확인] 노출.
"""
import re, openpyxl

# ── 충족률 권장 벤치마크 (단위: 만원). 출처 주석.
#   진단계열만 금액기반, 비금액계열(운전자/실손/일당/응급)은 presence 가중.
# SOURCES:
#   암  : 일반암 적정 5,000 + 고액암(중위소득×3, 최대 1억~1.5억) → 권장 합 밴드별
#         (banksalad 암보험 진단금 가이드)
#   뇌혈관: 최소 3,000, 40대 이후 4,000~5,000 (cancerok/뇌혈관 적정수준)
#   심장  : 허혈성 평균 설계 3,000 (banksalad 심장질환 11개사 비교)
#   수술비: 중수술 500~1,000 + 종수술 (signalplanner 리모델링 표)
#   사망·후유: 가장세대 보장공백 대비, 밴드별 상향
BENCHMARK = {
    #          사망후유  암    뇌혈관  심장  수술비  사망후유는 '사망후유' 키
    '20s': {'사망·후유':5000, '암':8000,  '뇌혈관':3000, '심장':3000},
    '30s': {'사망·후유':10000,'암':10000, '뇌혈관':4000, '심장':3000},
    '40s': {'사망·후유':15000,'암':10000, '뇌혈관':5000, '심장':3000},
    '50s': {'사망·후유':10000,'암':10000, '뇌혈관':5000, '심장':4000},
    '60s': {'사망·후유':5000, '암':8000,  '뇌혈관':5000, '심장':4000},
}
# 비금액(presence) 카테고리: 핵심담보 보유개수 / 기준개수
PRESENCE = {
    '운전자':    {'keys':['합의금','변호사','대인','대물','자부상','6주미만'], 'need':4},
    '입원·일당': {'keys':['간병인','상해일당','질병일당','간호통합','중환자'], 'need':3},
    '실손·일배책':{'keys':['입원','통원','약값','일상배상','일배책','MRI'],      'need':3},   # ★v30 행명 '입원' 매칭
    '골절·화상': {'keys':['골절','화상','깁스','5대골절','중증화상'],          'need':3},
    # ★v30i 수술비 = 종·핵심수술담보 '개수' 기반(금액합·종슬래시 최댓값 폐기). 몇 종·몇 개 가입했는지가 기준.
    '수술비':    {'keys':['수술','창상'],                                    'need':4},
    '응급실·독감':{'keys':['응급실','독감','식중독','骨'],                      'need':2},
}
# 리포트 10카테고리 ← 엑셀 A열 15그룹
CATEGORY_GROUPS = {
    '사망·후유':  ['사망','후유장애'],
    '암':        ['암'],
    '뇌혈관':     ['뇌혈관'],
    '심장':       ['심장'],
    '수술비':     ['수술비','수술'],
    '운전자':     ['운전자','운전'],
    '입원·일당':  ['일당'],
    '실손·일배책':['실손','일배책'],
    '골절·화상':  ['골 절','골절','화상','깁스'],
    '응급실·독감':['응급실','독감'],
}
DONUT_ORDER = ['암','운전자','실손·일배책','수술비','뇌혈관','사망·후유','골절·화상','심장','입원·일당','응급실·독감']

def _man(v):
    try: return float(v)
    except:
        # ★v30 슬래시 행(종수술 1-5/1-8종·MRI트리오): 대표값 = 최대 칸
        try:
            if isinstance(v,str) and '/' in v:
                return max(float(p) for p in v.split('/') if p.strip().replace('.','').isdigit())
        except: pass
        return 0.0

def _fmt(man):
    """만원 정수 → '1억7,100만' 한글단위"""
    man=int(round(man))
    if man<=0: return ''
    eok, rest = divmod(man,10000)
    s=''
    if eok: s+=f'{eok}억'
    if rest: s+=f'{rest:,}만'
    elif not eok: s='0'
    return s

def _bundle_adjust(path):
    """★v30e: [근거] 심장 섹션에서 묶음(여러 행 전개) 담보의 (행수-1)×금액을 카테고리별로 산출 → 보유합 중복 차감."""
    adj={'심장':0.0,'뇌혈관':0.0}
    try:
        wb=openpyxl.load_workbook(path,data_only=True)
        if '📋확인사항' not in wb.sheetnames: return adj
        qs=wb['📋확인사항']; on=False
        for r in range(1,qs.max_row+1):
            a=str(qs.cell(r,1).value or '')
            if '[근거] 심장' in a: on=True; continue
            if on:
                rows=str(qs.cell(r,3).value or ''); amt=qs.cell(r,4).value
                if not rows or rows=='기재 행': continue
                if '·' in rows and isinstance(amt,(int,float)):
                    parts=[p.strip() for p in rows.split('·')]
                    n_h=sum(1 for p in parts if p not in ('뇌졸증진단비','뇌혈관진단비'))
                    n_b=len(parts)-n_h
                    if n_h>1: adj['심장']+=amt*(n_h-1)
                    if n_b>1: adj['뇌혈관']+=amt*(n_b-1)
    except Exception: pass
    return adj


def _ci_meta(path):
    """★v33 선지급률 정본 계산 — CI 계약 '열'에서 직접 읽는다.
    끝열 '중대한CI적용' 은 비CI 계약 일반사망이 합산돼 오염되어 있어 사용 금지.
    지침: 선지급률은 50% 또는 80% 두 가지뿐."""
    wb=openpyxl.load_workbook(path,data_only=True); ws=wb.active
    last=ws.max_column
    def _isci(t):
        t=re.sub(r'[\s\u3000]','',str(t or ''))
        return any(k in t for k in ('CI보험','리빙케어','GI보험'))
    cols=[c for c in range(3,last) if _isci(ws.cell(1,c).value)]
    if not cols: return None
    rows={}
    for r in range(6,ws.max_row+1):
        b=str(ws.cell(r,2).value or '').strip()
        if b in ('중대한 암','중대한 뇌졸증','중대한 급성심근','중대한CI적용'): rows[b]=r
    c=cols[0]
    # 본체 = 중대한 암·뇌졸증 (급성심근은 CI추가보장특약이 가산돼 오염)
    pure=[_man(ws.cell(rows[k],c).value) for k in ('중대한 암','중대한 뇌졸증') if k in rows]
    pure=[v for v in pure if v>0]
    if not pure: return None
    bonche=max(pure)
    resid=_man(ws.cell(rows['중대한CI적용'],c).value) if '중대한CI적용' in rows else 0
    samang=bonche+resid
    if not samang: return None
    raw=bonche/samang*100
    pct=80 if abs(raw-80)<=abs(raw-50) else 50     # ★50/80 두 가지뿐
    return {'bonche':bonche,'samang':samang,'resid':resid,'pct':pct,'raw':round(raw)}


def load_excel(path):
    """완성 엑셀 → (groups_rows, headers). groups_rows[cat]=[(담보명,끝열값),..]"""
    wb=openpyxl.load_workbook(path,data_only=True); ws=wb.active
    last=ws.max_column
    # 그룹 경계: A열에 구분명 등장하는 행
    bounds=[]
    for r in range(6,ws.max_row+1):
        a=ws.cell(r,1).value
        if a and str(a).strip(): bounds.append((r,str(a).strip()))
    bounds.append((ws.max_row+1,'__END__'))
    grp_rows={}  # 엑셀그룹명 → [(담보,값)]
    disp={}      # ★v30h 담보명 → 슬래시 원문 표시(진단서에서 합산·최댓값 대신 가로 그대로)
    for i in range(len(bounds)-1):
        r0,name=bounds[i]; r1=bounds[i+1][0]
        rows=[]
        for r in range(r0,r1):
            b=ws.cell(r,2).value; _raw=ws.cell(r,last).value; v=_man(_raw)
            if b:
                _b=str(b).strip(); rows.append((_b,v))
                if isinstance(_raw,str) and '/' in _raw and any(c.isdigit() for c in _raw):
                    disp[_b]='('+_raw.strip()+')'   # 종수술비(1-5종)·n대·MRI = 슬래시 원문
        grp_rows[name]=rows
    # ★v30e 등식3: 혈전용해치료비(심장, 일당 그룹 첫 행 구조) → 심장으로 재배치 / 뇌쪽 혈전용해는 뇌혈관 그룹에 이미 소속
    if '일당' in grp_rows:
        _mv=[(b,v) for b,v in grp_rows['일당'] if '혈전용해' in b]
        if _mv:
            grp_rows['일당']=[(b,v) for b,v in grp_rows['일당'] if '혈전용해' not in b]
            grp_rows.setdefault('심장',[]).extend(_mv)
    load_excel._disp=disp   # ★v30h caller에 표시맵 전달(반환 시그니처 불변 → 회귀 0)
    # ★2026.07.11 실손 세대 자동판별용: '실손' 구분 그룹의 코어 담보(입원·통원·약값·의료비) 행
    _sil_rows=[]
    for _i in range(len(bounds)-1):
        _r0,_nm=bounds[_i]; _r1=bounds[_i+1][0]
        if '실손' in str(_nm):
            _sil_rows=[r for r in range(_r0,_r1)
                       if ws.cell(r,2).value
                       and any(k in str(ws.cell(r,2).value) for k in ('입원','통원','약값','의료비'))
                       and not any(x in str(ws.cell(r,2).value) for x in ('MRI','도수','비급여주사','일당'))]
            break
    # 헤더(계약 메타)
    headers=[]
    for c in range(2,last):  # 마지막=합계 제외
        nm=ws.cell(1,c).value; pr=_man(ws.cell(2,c).value)
        if nm is None and pr==0: continue
        _raw=str(nm or '')                                       # ★원본(회사\n상품\n[갱신])
        _rl=[x.strip() for x in _raw.split('\n') if x.strip()]
        _co_=(_rl[0] if _rl else '')
        _pr_=(_rl[1] if len(_rl)>1 else '')
        _pr_=re.sub(r'\s*\[[^\]]*\]','',_pr_).strip()
        nm=str(nm or '').replace('\n',' ').strip()
        renew = '[갱신]' in nm   # 헤더 형식 [갱신]/[비갱신(종신)] — 대괄호 정확매칭
        nm=re.sub(r'\s*\[[^\]]*\]','',nm)                       # [갱신]·[비갱신(종신)] 제거
        nm=re.sub(r'\((무|표준형|종신|표준형-종신|갱신|비갱신)\)','',nm)  # 괄호 수식 제거
        nm=re.sub(r'\s+',' ',nm).strip()
        _join=str(ws.cell(3,c).value or '').strip()             # 3행=가입년일
        _hassil=any(_man(ws.cell(r,c).value)>0 for r in _sil_rows)  # 실손 담보 보유 계약?
        headers.append({'nm':nm or '계약','amt':int(pr),'renew':renew,'join':_join,'sil':_hassil,'co':_co_,'prod':_pr_})
    total_prem=int(_man(ws.cell(2,last).value))
    return grp_rows, headers, total_prem

def cat_total(grp_rows, cat):
    """리포트 카테고리 보유합(만원) + 대표담보 top3"""
    rows=[]
    for g in CATEGORY_GROUPS[cat]:
        rows+=grp_rows.get(g,[])
    total=sum(v for _,v in rows)
    top=sorted([(b,v) for b,v in rows if v>0], key=lambda x:-x[1])[:3]
    return total, top, rows

def pct_for(cat, grp_rows, band):
    total,top,rows=cat_total(grp_rows,cat)
    if cat in BENCHMARK[band]:
        rec=BENCHMARK[band][cat]
        return (round(total/rec*100) if rec else 0), total, top
    # presence 계열
    spec=PRESENCE.get(cat)
    if not spec: return 0,total,top
    have=sum(1 for b,v in rows if v>0 and any(k in b for k in spec['keys']))
    return round(have/spec['need']*100), total, top

def map_excel_to_report(xlsx_path, settings=None, age_band='40s', age_known=False):
    """완성 엑셀 → rep dict (report_weasy.build_report_pdf 입력)"""
    settings=settings or {}
    grp_rows, headers, total_prem = load_excel(xlsx_path)
    _badj=_bundle_adjust(xlsx_path)   # ★v30e 묶음 전개 중복 차감(심장·뇌 보유합)
    client=settings.get('client','고객')

    coverage=[]; donut_map={}; detail_map={}
    for cat in CATEGORY_GROUPS:
        p,total,top=pct_for(cat,grp_rows,age_band)
        if cat in ('심장','뇌혈관') and _badj.get(cat):
            total=max(0,total-_badj[cat])
            rec=BENCHMARK[age_band].get(cat)
            if rec: p=round(total/rec*100)
        donut_map[cat]=p
        if cat in BENCHMARK[age_band]:
            detail_map[cat]={'have':_fmt(total) or '0','rec':_fmt(BENCHMARK[age_band][cat]),'unit':'만'}
        else:
            spec=PRESENCE.get(cat,{'keys':[],'need':1})
            have=sum(1 for b,v in cat_total(grp_rows,cat)[2] if v>0 and any(k in b for k in spec['keys']))
            detail_map[cat]={'have':f'{have}개','rec':f"{spec['need']}개",'unit':'개'}
        status='full' if p>=70 else ('part' if p>=40 else 'gap')
        blue = cat in ('실손·일배책','입원·일당')  # 실손·간병인·일배책 항상 파랑
        _disp=getattr(load_excel,'_disp',{})   # ★v30h 슬래시 원문 우선
        items=[{'t':b,'v':(_disp.get(b) or _fmt(v)),**({'blue':True} if blue else {})} for b,v in top]
        if not items or all(not it['v'] for it in items):
            items=[{'t':f'{cat} 없음','none':True}]
        coverage.append({'name':cat if cat!='심장' else '심장 (＋빈맥)','status':status,'items':items})

    # 강점/공백 = pct 임계
    ranked=sorted(donut_map.items(), key=lambda x:-x[1])
    strength=[{'h':c,'d':f'충족률 {p}% — 핵심담보 보유'} for c,p in ranked if p>=70][:4]
    weak=[{'h':c,'d':f'충족률 {p}% — 보강 필요'} for c,p in ranked if p<40][:4]
    gap_count=sum(1 for _,p in ranked if p<40)

    renew_list=[{'nm':h['nm'][:18],'v':f"{h['amt']:,}원"} for h in headers if h['renew']]
    nonren_list=[{'nm':h['nm'][:18],'v':f"{h['amt']:,}원"} for h in headers if not h['renew']]
    _co=lambda nm:(nm.split(' ')[0] if ' ' in nm else nm[:6])
    bars=sorted([{'nm':_co(h['nm']),'amt':h['amt'],'renew':h['renew']} for h in headers],
                key=lambda x:-x['amt'])
    donuts=[{'name':('심장' if c=='심장' else c.split('·')[0] if c in('실손·일배책','입원·일당','응급실·독감','골절·화상','사망·후유') else c),
             'pct':min(100,donut_map[c])} for c in DONUT_ORDER]
    # 도넛 라벨 보정
    label={'실손·일배책':'실손·배상','입원·일당':'입원·일당','응급실·독감':'응급실·독감','골절·화상':'골절·화상','사망·후유':'사망·후유'}
    # ★버킷 방식(지점장 2026.07.12): 도넛은 100%까지만, 초과분은 '충분+' 배지
    _raw={}
    for c in DONUT_ORDER:
        t,_tp,_rw=cat_total(grp_rows,c)
        if c in ('심장','뇌혈관') and _badj.get(c): t=max(0,t-_badj[c])
        if c in BENCHMARK[age_band] and BENCHMARK[age_band][c]:
            _raw[c]=round(t/BENCHMARK[age_band][c]*100)
        else:
            _sp=PRESENCE.get(c,{'keys':[],'need':1})
            _hv=sum(1 for b,v in _rw if v>0 and any(k in b for k in _sp['keys']))
            _raw[c]=round(_hv/_sp['need']*100) if _sp['need'] else 0
    donuts=[{'name':label.get(c,c),'pct':min(100,donut_map[c]),'raw':_raw.get(c,0),
             'over':_raw.get(c,0)>100} for c in DONUT_ORDER]

    # ── 치료비 정리 5항목 (라벨 → 엑셀 정확 담보명) ──
    CHIRYO=[('암주요치료비','암주요치료비'),('비급여주요치료비','하이클래스(암)'),
            ('순환계주요치료비','2대 주요치료비'),('산정특례(뇌혈관)','산정특례뇌혈관'),
            ('산정특례(심장)','산정특례심장')]
    allrows={}
    for rows in grp_rows.values():
        for b,v in rows:
            if v>0: allrows[b]=max(allrows.get(b,0),v)
    # ★v39 원본담보명 로드 (_dambo_raw 숨김시트) → 워크시트 흰칸에 '담보명 금액' 표기
    _raw_map={}
    try:
        import openpyxl as _ox
        _wbr=_ox.load_workbook(xlsx_path, data_only=True)
        if '_dambo_raw' in _wbr.sheetnames:
            _rs=_wbr['_dambo_raw']
            for _r in range(2,_rs.max_row+1):
                _s=_rs.cell(_r,1).value; _rw=_rs.cell(_r,2).value
                if _s and _rw: _raw_map[str(_s).strip()]=str(_rw).strip()
    except Exception:
        pass
    def _clean_dnm(nm):   # 원본담보명 정리: 괄호주석·병원접두·회차 제거해 짧게
        import re as _re
        s=_re.sub(r'[（(].*?[）)]','',nm)                       # 괄호내용 제거
        s=_re.sub(r'상급종합병원[ᅵI|\s]*(II|Ⅱ|III|Ⅲ)?\s*','',s) # 병원접두 제거
        s=_re.sub(r'\s+',' ',s).strip(' ·-|Ⅲ III II')
        return s[:22]
    chiryo=[{'name':lab,'value':(_fmt(allrows.get(key,0)) or '미가입'),
             'raw':_clean_dnm(_raw_map[key]) if _raw_map.get(key) else ''} for lab,key in CHIRYO]

    # ── CI/생명보험 선지급 분석 (사망값 의존X: 선지급률=본체/(본체+잔여)) ──
    def _gv(nm):
        for rows in grp_rows.values():
            for b,v in rows:
                if str(b).strip()==nm: return v
        return 0
    # ★P5 상단 핵심 진단비 3종(지점장 지시): 뇌출혈·뇌졸증·급성심근경색
    def _fv(v): return _fmt(v) if v and v>0 else '미가입'
    _p5=[('뇌출혈진단비',_gv('뇌출혈진단비') or _gv('중대한 뇌졸증')),
         ('뇌졸증진단비',_gv('뇌졸증진단비') or _gv('뇌혈관진단비')),
         ('급성심근경색',_gv('급성심근경색') or _gv('중대한 급성심근') or _gv('허혈성 진단비')),
         ('뇌혈관진단비',_gv('뇌혈관진단비')),
         ('허혈성 진단비',_gv('허혈성 진단비') or _gv('허혈성심장질환'))]
    p5_own=[{'t':n,'v':_fv(v)} for n,v in _p5]
    _ci_pairs=[(n,_gv(n)) for n in ('중대한 암','중대한 뇌졸증','중대한 급성심근')]
    _ci_pairs=[(n,v) for n,v in _ci_pairs if v>0]
    _ci_apply=_gv('중대한CI적용')
    _ci_bonche=max((v for _,v in _ci_pairs), default=0)
    # ★v33 선지급률: 끝열(_ci_apply)은 비CI 일반사망 오염 → CI 계약 열에서 직접 산출
    _cm=_ci_meta(xlsx_path)
    if _cm:
        _ci_bonche=_cm['bonche']; _ci_samang=_cm['samang']; _ci_rate=_cm['pct']; _ci_apply=_cm['resid']
    else:
        _ci_samang=_ci_bonche+_ci_apply
        _ci_rate=round(_ci_bonche/_ci_samang*100) if _ci_samang else 0
        _ci_rate=(80 if abs(_ci_rate-80)<=abs(_ci_rate-50) else 50) if _ci_rate else 0
    ci={'present':bool(_ci_pairs or _ci_apply>0),'samang':_fmt(_ci_samang),
        'rate':_ci_rate,'residual':_fmt(_ci_apply),
        'items':[{'t':{'중대한 암':'ci암진단비','중대한 뇌졸증':'ci뇌졸증','중대한 급성심근':'ci급성심근경색'}.get(n,n),'v':_fmt(v)} for n,v in _ci_pairs]}
    # ★CI 3상태 판정(2026.07.07 지점장): 상품명 CI/GI/리빙케어 + 중대한OO담보 값
    _ci_prod=any(('CI' in str(h.get('nm','')) or '리빙케어' in str(h.get('nm','')) or 'GI보험' in str(h.get('nm',''))) for h in headers)
    # ★2026.07.12 지점장 확정: 상품명에 CI/GI/리빙케어가 없으면 '중대한OO' 담보가 있어도 진짜 CI가 아니다(가짜).
    #   → 상품명이 1순위. 상품명에 표기 없으면 무조건 none(Plan B).
    if not _ci_prod:
        ci['status']='none'        # 상품명에 CI/GI/리빙케어 없음 = 확실 비CI (중대한OO는 가짜)
        ci['present']=False
        ci['items']=[]
    elif _ci_pairs or _ci_apply>0:
        ci['status']='ci'          # 상품명 CI + 담보값 있음 = 확실 CI
    else:
        ci['status']='check'       # 상품명 CI인데 담보값 없음/애매 = 회색지대 [확인]

    # ── Plan B: 비CI 진단비 정액 지급 구조 (CI 미보유 시 P3 상단 CI블록 대체) ──
    def _sumnm(*names):
        s=0
        for rows in grp_rows.values():
            for b,v in rows:
                if str(b).strip() in names and v>0: s+=v
        return s
    _amt_cancer=max(_gv('일반암'),_gv('고액암'))
    _amt_brain=_sumnm('뇌혈관진단비','뇌졸증진단비')
    _amt_heart=_sumnm('급성심근경색','허혈성 진단비')
    noci_items=[]
    if _amt_cancer>0: noci_items.append({'t':'암 진단비','v':_fmt(_amt_cancer)})
    if _amt_brain>0:  noci_items.append({'t':'뇌혈관·뇌졸증','v':_fmt(_amt_brain)})
    if _amt_heart>0:  noci_items.append({'t':'급성심근·허혈성','v':_fmt(_amt_heart)})
    noci={'present':(not ci['present']) and bool(noci_items),'items':noci_items}

    rep={
        'client':client,
        'branch':settings.get('branch',''),'manager':settings.get('manager',''),
        'title':settings.get('title',''),'phone':settings.get('phone',''),
        'n_contract':len(headers),'premium':total_prem,
        'renew':len(renew_list),'nonrenew':len(nonren_list),'gap_count':gap_count,
        'coverage':coverage,'strength':strength,'weak':weak,
        'renew_list':renew_list,'nonrenew_list':nonren_list,
        'premium_bars':bars,'donuts':donuts,
        'donut_detail':[{'name':label.get(c,c),'have':detail_map[c]['have'],
                         'rec':detail_map[c]['rec'],'pct':min(100,donut_map[c]),
                         'raw':_raw.get(c,0),'over':_raw.get(c,0)>100} for c in DONUT_ORDER],
        'band_label':{'20s':'20대','30s':'30대','40s':'40대','50s':'50대','60s':'60대'}.get(age_band,age_band),
        'chiryo':chiryo,
        'ci':ci,
        'noci':noci,
        'p5_own':p5_own,
        'age_band':age_band,'age_known':age_known,
    }
    # ── 리모델링 제안: 1-5종 권유 · 운전자 재가입 권유 (지침 §7·§8.6) ──
    advice=[]
    try:
        _wb2=openpyxl.load_workbook(xlsx_path, data_only=True); _ws2=_wb2.active
        _rowmap={}
        for _r in range(6,_ws2.max_row+1):
            _b=_ws2.cell(_r,2).value
            if _b: _rowmap[str(_b).strip()]=_ws2.cell(_r,_ws2.max_column).value
        def _has(nm):
            v=_rowmap.get(nm); return (v not in (None,'',0)) and (str(v).strip() not in ('','0'))
        def _num(nm):
            v=_rowmap.get(nm)
            try: return int(float(str(v).replace(',','').split('/')[0]))
            except: return 0
        # 수술: 1-8종/N대만 있고 1-5종 없으면 → 1-5종 권유
        if (_has('상해 종수술비(1-8종)') or _has('질병 종수술비(1-8종)') or _has('120대수술비')) \
           and not (_has('상해 종수술비(1-5종)') or _has('질병 종수술비(1-5종)')):
            advice.append({'t':'수술비 리모델링 권유',
                'd':'현재 수술비가 1-5종이 아닙니다(1-7·8·9종 또는 N대). 이 유형은 비급여 항목 미보장·청구 절차 복잡으로 실질 보장 범위가 좁습니다. 관혈/비관혈을 구분하고 비급여까지 보장하는 1-5종 수술비 가입(전환)을 권유드립니다.'})
        # 운전자: 최신 기준 미달이면 → 재가입 권유
        _std={'합의금':20000,'6주미만':1000,'대인':3000,'대물':500,'변호사':5000}
        _lab={'합의금':'합의금(기준 2억)','6주미만':'6주미만 합의금(기준 1천만)','대인':'벌금 대인(기준 3천만)','대물':'벌금 대물(기준 500만)','변호사':'변호사선임(기준 5천만)'}
        if any(_num(k)>0 for k in _std):
            _short=[f"{_lab[k]}: {('보유 '+format(_num(k),',')+'만') if _num(k)>0 else '미보유'}" for k in _std if _num(k)<_std[k]]
            if _short:
                advice.append({'t':'운전자 리모델링 권유',
                    'd':'2022년 보행자보호의무·처벌 강화로 옛 운전자보험은 담보가 부족합니다. 미달 항목 — '+' / '.join(_short)+'. 최신 기준으로 재가입(리모델링)을 권유드립니다.'})
    except Exception:
        pass
    # ── ★P5 질병코드 커버표: 고객 실제 보유 감지 (하드코딩 제거) ──
    _allnm=set()
    for _rows in grp_rows.values():
        for _b,_v in _rows:
            if _v and _v>0: _allnm.add(str(_b).strip())
    def _any(*subs):
        return any(any(s in nm for s in subs) for nm in _allnm)
    scope_brain=[]   # 보유 행 key
    if _any('뇌출혈','뇌혈관진단','중대한 뇌졸증','뇌졸증진단'): scope_brain.append('hem')
    if _any('뇌경색','뇌졸증진단','뇌혈관진단'): scope_brain.append('infarct')
    if _any('뇌혈관진단'): scope_brain.append('other')
    if _any('외상성뇌출혈','외상성 뇌출혈'): scope_brain.append('trauma')
    if _any('산정특례뇌','산정특례(뇌'): scope_brain.append('brain_snjt')
    scope_heart=[]
    if _any('급성심근','중대한 급성심근'): scope_heart.append('ami')
    if _any('허혈성 진단비','허혈심장','허혈성진단'): scope_heart+=['angina','chronic']
    if _any('부정맥'): scope_heart.append('arrhy')
    if _any('심부전'): scope_heart.append('hf')
    if _any('심장판막','판막'): scope_heart.append('valve')
    if _any('심장염증','심근염','심내막','심장막'): scope_heart.append('inflam')
    if _any('심근병증'): scope_heart.append('cardiomyo')
    if _any('산정특례심장','산정특례(심'): scope_heart.append('heart_snjt')
    rep['scope_brain']=scope_brain
    rep['scope_heart']=scope_heart
    rep['advice']=advice
    # ★ 전체 담보→표시값 맵 (워크시트 자동주입용; coverage 상위3개 한계 보완, 2026.07.11)
    import re as _re2
    _disp2=getattr(load_excel,'_disp',{})
    def _nrm(s): return _re2.sub(r'[\s·()\[\]/]','',str(s))
    _dm={}
    for _rows in grp_rows.values():
        for _b,_v in _rows:
            if not _b: continue
            _dv=_disp2.get(_b) or _fmt(_v)
            if _dv and _dv!='0':
                _dm.setdefault(_nrm(_b), _dv)
    # 골절 합산(치아파절 포함+제외, 5대·수술 제외) = PPT 골절 한 값
    _gj=0
    for _rows in grp_rows.values():
        for _b,_v in _rows:
            if _b and '골절' in _b and '5대' not in _b and '수술' not in _b and isinstance(_v,(int,float)):
                _gj+=int(_v)
    if _gj>0: _dm['골절합산']=_fmt(_gj)
    rep['dambo']=_dm
    # ★2026.07.11 실손 세대 자동판별(CI식 3상태): 실손 계약 가입일 → 세대
    def _gen_of(js, comp=''):
        import re as _r
        m=_r.search(r'(\d{4})\D+(\d{1,2})(?:\D+(\d{1,2}))?', str(js))
        if not m: return None
        from datetime import date
        y=int(m.group(1)); mo=int(m.group(2)); d=int(m.group(3) or 1)
        try: dt=date(y,mo,d)
        except Exception: return None
        # ★2026.07.11 지점장 확정: 1세대=~2009.09 / 2세대=2009.10~ (2009.07~09은 회사별 상이 → 주석)
        #   2세대 3분할(2-1 ~2012.12 / 2-2 2013.01~2015.12 / 2-3 2016.01~2017.03)
        #   1세대는 생보·손보 구분(자기부담 손보0%·생보20%, 상해의료비 별도)
        _LIFE=('생명','AIA','ABL','푸본','교보','동양','미래에셋','신한','KDB','메트라이프','처브','라이나','KB라이프','라이프플래닛','하나생명','IBK연금')
        _NONLIFE=('손보','손해','화재','해상','손해보험')
        sub=''
        if dt<=date(2009,9,30):
            g=1
            _c1=str(comp).split('\n')[0].strip()   # ★회사명 줄만(상품명의 '라이프' 오매칭 차단)
            if any(k in _c1 for k in _NONLIFE): sub='손보'
            elif any(k in _c1 for k in _LIFE): sub='생보'
            else: sub='손보'
        elif dt<=date(2017,3,31):
            g=2
            if dt<=date(2012,12,31): sub='2-1'
            elif dt<=date(2015,12,31): sub='2-2'
            else: sub='2-3'
        elif dt<=date(2021,6,30): g=3
        elif dt<=date(2026,5,5): g=4
        else: g=5
        dstr=f'{y}.{mo:02d}.{d:02d}' if m.group(3) else f'{y}.{mo:02d}'
        return {'gen':g,'sub':sub,'date':dstr}
    _sil=[h for h in headers if h.get('sil')]
    if not _sil:
        rep['silson_gen']={'status':'none'}
    else:
        _sh=min(_sil, key=lambda h:str(h.get('join','')))  # 가장 오래된 실손 계약(세대 판별 기준)
        _cnm=_sh.get('nm','')
        _g=_gen_of(_sh.get('join'), _cnm)
        # 실손 계약 전체 목록(회사·상품명·가입일·보험료)
        _sillist=[]
        for _h in sorted(_sil, key=lambda x:str(x.get('join',''))):
            _sillist.append({'co':str(_h.get('co',''))[:14],
                             'prod':str(_h.get('prod',''))[:34],
                             'join':str(_h.get('join','')),
                             'amt':_h.get('amt',0),
                             'renew':_h.get('renew',False)})
        rep['silson_list']=_sillist
        if _g:
            rep['silson_gen']={'status':'auto','gen':_g['gen'],'sub':_g['sub'],'date':_g['date'],
                               'company':str(_sh.get('co',''))[:14],
                               'product':str(_sh.get('prod',''))[:34]}
        else:
            rep['silson_gen']={'status':'check','company':_cnm[:12]}
    return rep

if __name__=='__main__':
    import sys
    rep=map_excel_to_report(sys.argv[1] if len(sys.argv)>1 else '보장진단_정기철.xlsx',
        settings={'client':'정기철','branch':'온빛센터 바름지점','manager':'최은혜','title':'지점장','phone':'010-XXXX-XXXX'})
    for d in rep['donuts']: print(f"  {d['name']:<8} {d['pct']}%")
    print('계약',rep['n_contract'],'/ 보험료',f"{rep['premium']:,}",'/ 공백',rep['gap_count'])
