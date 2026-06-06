# -*- coding: utf-8 -*-
"""
바름 보장분석 추출엔진 v1  (영구법칙: PyMuPDF 확정)
poppler 금지(한글 0개). PyMuPDF(fitz)로만 추출.
계약별 세로 추출(법칙4) · 갱신/비갱신 계약레벨(법칙12)+담보칸별 override(법칙13 칸>행).
사용: python barum_extract.py 보장분석.pdf  →  contracts.json
"""
import fitz, re, json, sys, datetime

AMOUNT = re.compile(r'^[\d,]+$')
FOOTER = re.compile(r'(지점|LP|☎|\d+/\d+|^\d{4}\.\d{2}\.\d{2}$|상기 내용|본 리포트)')
GAENG  = re.compile(r'\(갱신')          # 담보명 칸별 갱신 표식
YEARS  = re.compile(r'(\d{4})\.\d{2}\.\d{2}')

def parse_page(text):
    L=[l.strip() for l in text.split('\n') if l.strip()]
    if not L or L[0]!='(정상계약 리스트)':
        return None
    회사, 상품 = L[1], L[2]
    납입=만기=''
    for i,l in enumerate(L):
        if l.startswith('납입주기') and i+1<len(L): 납입=L[i+1]
        if '~' in l and YEARS.search(l): 만기=l
    start=0
    for i,l in enumerate(L):
        if l=='보장기간': start=i+2; break
    covers=[]; buf=[]
    for l in L[start:]:
        if FOOTER.search(l): break
        if AMOUNT.match(l) and buf:
            nm=' '.join(buf).replace(' -','-')
            covers.append([nm, int(l.replace(',','')), bool(GAENG.search(nm))])
            buf=[]
        elif not AMOUNT.match(l):
            buf.append(l)
    return {'회사':회사,'상품':상품,'납입':납입,'만기':만기,'담보':covers}

def contract_renewal(c):
    """법칙12 정순"""
    상품, 납입, 만기 = c['상품'], c['납입'], c['만기']
    if '갱신형' in 상품: return '갱신'        # ①
    if '운전자' in 상품: return '비갱신'       # ②
    if '9999' in 만기:  return '비갱신'       # ③
    m=re.search(r'/(\d+)년', 납입); 납입년=int(m.group(1)) if m else 0
    if 납입년*12 > 240: return '비갱신'        # ④ 총회차>240
    yrs=YEARS.findall(만기)
    if 납입년 and len(yrs)==2:                 # ⑤ 납입≠보장→비갱신
        return '갱신' if 납입년==(int(yrs[1])-int(yrs[0])) else '비갱신'
    return '비갱신'

def is_expired(c):
    yrs=YEARS.findall(c['만기'])
    return bool(yrs) and int(yrs[-1]) < datetime.date.today().year

def extract(pdf):
    d=fitz.open(pdf)
    raw=[p for p in (parse_page(pg.get_text()) for pg in d) if p]
    merged=[]
    for p in raw:                              # 동일상품 연속페이지 병합(법칙4)
        if merged and merged[-1]['상품']==p['상품'] and merged[-1]['만기']==p['만기']:
            merged[-1]['담보']+=p['담보']
        else:
            merged.append(p)
    merged=[c for c in merged if not is_expired(c)]   # 만료계약 제외
    for i,c in enumerate(merged,1):
        c['계약']=i
        c['갱신구분']=contract_renewal(c)
        c['담보수']=len(c['담보'])
        c['갱신담보']=[nm for nm,_,g in c['담보'] if g]
    return merged

if __name__=='__main__':
    cs=extract(sys.argv[1] if len(sys.argv)>1 else 'src.pdf')
    tot=0
    for c in cs:
        prem=re.search(r'[\d,]+', c.get('납입','')); 
        gd=f"  ·칸별갱신 {len(c['갱신담보'])}건" if c['갱신담보'] else ""
        print(f"계약{c['계약']:2} [{c['갱신구분']:3}] {c['회사']:13}{c['상품'][:26]:26} 담보{c['담보수']:3}{gd}")
    json.dump(cs, open('contracts.json','w'), ensure_ascii=False, indent=1)
    print(f"\n총 {len(cs)}계약(만료제외) 추출 → contracts.json")
