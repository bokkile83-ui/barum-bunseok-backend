# -*- coding: utf-8 -*-
"""
let_engine — 롯데/let GA 양식 보장분석지 → 범용 추출 엔진
검증완료(5파일, 계약수 3·4·8·10·11):
  · 한장보장표 53담보 가입금액 = 전수 일치
  · 계약별 그리드 세로합 == 한장보장표 합계 = 불일치 0 (보장지수 밴드 앵커)
원리: 담보명·페이지헤더가 이미지라도 (1)한장보장표 토큰순서 (2)보장지수(부족/충분/미가입)
      가 모든 행에 찍혀 미가입행도 안 사라지는 점을 앵커로 위치 직독.
"""
import fitz, re
from collections import defaultdict

_NUM = re.compile(r'-?[\d,]+')
def _i(s):
    s=str(s).replace(',','').replace('원','').strip()
    if s in ('-',''): return 0
    m=re.search(r'-?\d+', s)
    return int(m.group()) if m else 0

# 한장보장표 53담보 순서 (5블록: 14·7·9·12·11). 표준금액 시그니처로 페이지 검출.
_STD0 = ['10,000','20,000','3,000','3,000','10,000','5,000','5,000','10,000','3,000','2,000','300','7,000','2,000','5,000']
CATS = (['질병사망','상해사망','질병후유3%','질병후유50%','질병후유80%','상해후유3%','상해후유50%','상해후유80%',
 '일반암','고액암','암수술','고액항암','재진단암','전이암']
 +['뇌혈관진단','뇌졸중진단','뇌출혈진단','뇌혈관수술','허혈성심장','급성심근경색','심장수술']
 +['중증뇌혈관','중증심장','중증암','중증치매','특정중증','장기요양1급','장기요양1_2급','장기요양1_2_3_4급','경증이상치매']
 +['상해입원의료','상해통원의료','질병입원의료','질병통원의료','상해수술','질병수술','질병입원','질병중환자실','간병질병입원','상해입원','상해중환자실','간병상해입원']
 +['교통사고처리','자동차벌금대인','자동차벌금대물','변호사선임','부상위로금','응급실내원','화상진단','골절진단','깁스치료','화재벌금','일상생활배상'])
_SIZES = [14,7,9,12,11]
GROUP = {  # 출력 구분 그룹
 '사망':['질병사망','상해사망'],
 '후유장해':['질병후유3%','질병후유50%','질병후유80%','상해후유3%','상해후유50%','상해후유80%'],
 '암':['일반암','고액암','암수술','고액항암','재진단암','전이암'],
 '뇌혈관':['뇌혈관진단','뇌졸중진단','뇌출혈진단','뇌혈관수술'],
 '심장':['허혈성심장','급성심근경색','심장수술'],
 '산정특례':['중증뇌혈관','중증심장','중증암','중증치매','특정중증'],
 '요양/치매':['장기요양1급','장기요양1_2급','장기요양1_2_3_4급','경증이상치매'],
 '실손의료':['상해입원의료','상해통원의료','질병입원의료','질병통원의료'],
 '수술비':['상해수술','질병수술'],
 '입원비':['질병입원','질병중환자실','간병질병입원','상해입원','상해중환자실','간병상해입원'],
 '운전자':['교통사고처리','자동차벌금대인','자동차벌금대물','변호사선임','부상위로금'],
 '기타':['응급실내원','화상진단','골절진단','깁스치료','화재벌금','일상생활배상'],
}

def 한장보장표(d):
    """53담보 가입금액 dict + page index. 결정론적."""
    for pi in range(len(d)):
        L = [x.strip() for x in d[pi].get_text('text').split('\n') if x.strip()]
        vals = [x for x in L if re.fullmatch(r'-|-?[\d,]+', x)]
        if vals[:14] == _STD0 and len(vals) >= 106:
            out, i = [], 0
            for bi, sz in enumerate(_SIZES):
                i += sz                              # 표준 스킵
                out += [_i(x) for x in vals[i:i+sz]] # 가입
                i += sz
                if bi == 3: i += 3                   # 블록4 뒤 카운트(충분/부족/미가입) 스킵
            return dict(zip(CATS, out)), pi
    return None, None

def 보유계약(d):
    """보유계약현황에서 계약 메타 파싱: 회사/상품/계약일/만기일/납입개월/횟수/보험료/갱신."""
    txt = ''
    for pi in range(min(3, len(d))):
        t = d[pi].get_text('text')
        if '보험서비스' in t or '납입횟수' in t: txt += '\n' + t
    L = [x.strip() for x in txt.split('\n') if x.strip()]
    rows, i = [], 0
    date = re.compile(r'\d{4}\.\d{2}\.\d{2}')
    cnt = re.compile(r'^\d+/\d+$')
    while i < len(L):
        if L[i].isdigit() and i+8 < len(L) and date.match(L[i+3] if i+3<len(L) else ''):
            no=L[i]; co=L[i+1]; pr=L[i+2]; gd=L[i+3]; md=L[i+4]
            # 납입주기/횟수/보험료 뒤쪽 탐색
            j=i+5; cyc=''; cc=''; fee=''
            while j < min(i+12,len(L)):
                if cnt.match(L[j]): cc=L[j]
                if L[j].endswith('원') and not fee: fee=L[j]
                if L[j] in('월납','연납','일시납'): cyc=L[j]
                j+=1
            tot = int(cc.split('/')[1]) if cc else 0
            mat_y = int(gd.split('.')[0]) if date.match(gd) else 0
            con_y = int(md.split('.')[0]) if date.match(md) else 0
            # 갱신판정(5규칙): 9999=비갱신, 납입기간==보장기간→갱신, else 비갱신
            if mat_y>=9999: gen='비갱신(종신)'
            else:
                pay_y = round(tot/12) if tot else 0
                cov_y = mat_y-con_y
                gen='갱신' if pay_y and pay_y==cov_y else '비갱신'
            rows.append(dict(no=no,회사=co,상품=pr,계약일=md,만기일=gd,납입주기=cyc,납입개월=tot,횟수=cc,보험료=_i(fee),갱신=gen))
            i=j
        else: i+=1
    return rows

def 그리드밴드합(d):
    """보장지수 밴드 앵커로 계약별 그리드 세로합 산출 + 한장보장표 검산."""
    gp=[pi for pi in range(len(d)) if sum(1 for w in d[pi].get_text('words') if w[4] in('부족','충분','미가입'))>=15]
    if not gp: return None,None,gp
    w0=d[gp[0]].get_text('words')
    bands=sorted({round(y,0) for x0,y,*_ in [(w[0],w[1]) for w in w0] if False}) # placeholder
    bands=sorted({round(w[1],0) for w in w0 if w[4] in('부족','충분','미가입')})
    near=lambda y: min(bands,key=lambda b:abs(b-y))
    left=defaultdict(int); grid=defaultdict(int)
    for x0,y0,x1,y1,tt,*_ in w0:
        if re.fullmatch(r'-?[\d,]+',tt) and 175<x0<375 and y0>180: left[near(y0)]+=_i(tt)
    for pi in gp:
        for x0,y0,x1,y1,tt,*_ in d[pi].get_text('words'):
            if re.fullmatch(r'-?[\d,]+',tt) and x0>=440 and 180<y0<525: grid[near(y0)]+=_i(tt)
    return left,grid,gp

if __name__=='__main__':
    import sys
    d=fitz.open(sys.argv[1])
    t,pi=한장보장표(d); meta=보유계약(d); left,grid,gp=그리드밴드합(d)
    print(f'한장보장표 idx{pi}, 가입담보 {sum(1 for v in t.values() if v)}개')
    print(f'계약 {len(meta)}건:', [(m["회사"],m["갱신"]) for m in meta])
    bad=[b for b in left if left[b]!=grid[b]]
    print(f'그리드 검산: 밴드 {len(left)}개, 불일치 {len(bad)}, 세로합총={sum(grid.values())} 합계총={sum(t.values())}')
