# -*- coding: utf-8 -*-
"""
BARUM 충족률 엔진 + map_excel_to_report
- 입력: 완성된 보장진단 엑셀(.xlsx) 1개 (등식2: 리포트는 완성 엑셀만 읽음)
- 출력: report_weasy.build_report_pdf 가 먹는 rep dict
- 충족률 % = min(100, 보유합 / 연령밴드별 권장액 * 100)  (추측 아님: 아래 SOURCES 기반)
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
    '20s': {'사망·후유':5000, '암':8000,  '뇌혈관':3000, '심장':3000, '수술비':1500},
    '30s': {'사망·후유':10000,'암':10000, '뇌혈관':4000, '심장':3000, '수술비':2000},
    '40s': {'사망·후유':15000,'암':10000, '뇌혈관':5000, '심장':3000, '수술비':2000},
    '50s': {'사망·후유':10000,'암':10000, '뇌혈관':5000, '심장':4000, '수술비':2500},
    '60s': {'사망·후유':5000, '암':8000,  '뇌혈관':5000, '심장':4000, '수술비':2500},
}
# 비금액(presence) 카테고리: 핵심담보 보유개수 / 기준개수
PRESENCE = {
    '운전자':    {'keys':['합의금','변호사','대인','대물','자부상','6주미만'], 'need':4},
    '입원·일당': {'keys':['간병인','상해일당','질병일당','간호통합','중환자'], 'need':3},
    '실손·일배책':{'keys':['실손입원','통원','약값','일상배상','일배책'],      'need':3},
    '골절·화상': {'keys':['골절','화상','깁스','5대골절','중증화상'],          'need':3},
    '응급실·독감':{'keys':['응급실','독감','식중독','骨'],                      'need':2},
}
# 리포트 10카테고리 ← 엑셀 A열 15그룹
CATEGORY_GROUPS = {
    '사망·후유':  ['사망','후유장애'],
    '암':        ['암'],
    '뇌혈관':     ['뇌혈관'],
    '심장':       ['심장'],
    '수술비':     ['수술비'],
    '운전자':     ['운전자'],
    '입원·일당':  ['일당'],
    '실손·일배책':['실손','일배책'],
    '골절·화상':  ['골 절','골절','화상','깁스'],
    '응급실·독감':['응급실','독감'],
}
DONUT_ORDER = ['암','운전자','실손·일배책','수술비','뇌혈관','사망·후유','골절·화상','심장','입원·일당','응급실·독감']

def _man(v):
    try: return float(v)
    except: return 0.0

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
    for i in range(len(bounds)-1):
        r0,name=bounds[i]; r1=bounds[i+1][0]
        rows=[]
        for r in range(r0,r1):
            b=ws.cell(r,2).value; v=_man(ws.cell(r,last).value)
            if b: rows.append((str(b).strip(),v))
        grp_rows[name]=rows
    # 헤더(계약 메타)
    headers=[]
    for c in range(2,last):  # 마지막=합계 제외
        nm=ws.cell(1,c).value; pr=_man(ws.cell(2,c).value)
        if nm is None and pr==0: continue
        nm=str(nm or '').replace('\n',' ').strip()
        renew='(갱신)' in nm
        nm=re.sub(r'\s*\((갱신|비갱신)\)','',nm).strip()
        headers.append({'nm':nm or '계약','amt':int(pr),'renew':renew})
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
        return min(100, round(total/rec*100)) if rec else 0, total, top
    # presence 계열
    spec=PRESENCE.get(cat)
    if not spec: return 0,total,top
    have=sum(1 for b,v in rows if v>0 and any(k in b for k in spec['keys']))
    return min(100, round(have/spec['need']*100)), total, top

def map_excel_to_report(xlsx_path, settings=None, age_band='40s', age_known=False):
    """완성 엑셀 → rep dict (report_weasy.build_report_pdf 입력)"""
    settings=settings or {}
    grp_rows, headers, total_prem = load_excel(xlsx_path)
    client=settings.get('client','고객')

    coverage=[]; donut_map={}; detail_map={}
    for cat in CATEGORY_GROUPS:
        p,total,top=pct_for(cat,grp_rows,age_band)
        donut_map[cat]=p
        if cat in BENCHMARK[age_band]:
            detail_map[cat]={'have':_fmt(total) or '0','rec':_fmt(BENCHMARK[age_band][cat]),'unit':'만'}
        else:
            spec=PRESENCE.get(cat,{'keys':[],'need':1})
            have=sum(1 for b,v in cat_total(grp_rows,cat)[2] if v>0 and any(k in b for k in spec['keys']))
            detail_map[cat]={'have':f'{have}개','rec':f"{spec['need']}개",'unit':'개'}
        status='full' if p>=70 else ('part' if p>=40 else 'gap')
        blue = cat in ('실손·일배책','입원·일당')  # 실손·간병인·일배책 항상 파랑
        items=[{'t':b,'v':_fmt(v),**({'blue':True} if blue else {})} for b,v in top]
        if not items or all(not it['v'] for it in items):
            items=[{'t':f'{cat} 없음','none':True}]
        coverage.append({'name':cat if cat!='심장' else '심장 (＋빈맥)','status':status,'items':items})

    # 강점/공백 = pct 임계
    ranked=sorted(donut_map.items(), key=lambda x:-x[1])
    strength=[{'h':c,'d':f'충족률 {p}% — 핵심담보 보유'} for c,p in ranked if p>=70][:4]
    weak=[{'h':c,'d':f'충족률 {p}% — 보강 필요'} for c,p in ranked if p<40][:4]
    gap_count=sum(1 for _,p in ranked if p<40)

    renew_list=[{'nm':h['nm'],'v':f"{h['amt']:,}원"} for h in headers if h['renew']]
    nonren_list=[{'nm':h['nm'],'v':f"{h['amt']:,}원"} for h in headers if not h['renew']]
    bars=sorted([{'nm':h['nm'],'amt':h['amt'],'renew':h['renew']} for h in headers],
                key=lambda x:-x['amt'])
    donuts=[{'name':('심장' if c=='심장' else c.split('·')[0] if c in('실손·일배책','입원·일당','응급실·독감','골절·화상','사망·후유') else c),
             'pct':donut_map[c]} for c in DONUT_ORDER]
    # 도넛 라벨 보정
    label={'실손·일배책':'실손·배상','입원·일당':'입원·일당','응급실·독감':'응급실·독감','골절·화상':'골절·화상','사망·후유':'사망·후유'}
    donuts=[{'name':label.get(c,c),'pct':donut_map[c]} for c in DONUT_ORDER]

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
                         'rec':detail_map[c]['rec'],'pct':donut_map[c]} for c in DONUT_ORDER],
        'band_label':{'20s':'20대','30s':'30대','40s':'40대','50s':'50대','60s':'60대'}.get(age_band,age_band),
        'age_band':age_band,'age_known':age_known,
    }
    return rep

if __name__=='__main__':
    import sys
    rep=map_excel_to_report(sys.argv[1] if len(sys.argv)>1 else '보장진단_정기철.xlsx',
        settings={'client':'정기철','branch':'온빛센터 바름지점','manager':'최은혜','title':'지점장','phone':'010-XXXX-XXXX'})
    for d in rep['donuts']: print(f"  {d['name']:<8} {d['pct']}%")
    print('계약',rep['n_contract'],'/ 보험료',f"{rep['premium']:,}",'/ 공백',rep['gap_count'])
