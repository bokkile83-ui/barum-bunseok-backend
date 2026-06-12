# -*- coding: utf-8 -*-
import re
def 보유계약(d):
    """보유계약 리스트(2p) 파싱: No/회사/상품/계약일/만기일/회차/보험료/갱신. 최대 50건."""
    full='\n'.join(d[pi].get_text('text') for pi in range(min(3,len(d))))
    h=full.find('잔여보험료')
    if h>=0: full=full[h+len('잔여보험료'):]   # 헤더 뒤부터(상단 미흡카운트 노이즈 제거)
    pat=re.compile(r'\n(\d{1,2})\n([^\n]+)\n(.*?)(\d{4}\.\d{2}\.\d{2})\n?(\d{4}\.\d{2}\.\d{2})\n(월납|연납|일시납)\n(\d+)/(\d+)\n([\d,]+)원', re.S)
    out=[]
    for m in pat.finditer(full):
        no,co,prod,con,mat,cyc,paid,tot,fee=m.groups()
        if co.strip().isdigit(): continue           # 회사가 숫자면 노이즈
        if int(no)>50: continue
        prod=re.sub(r'\s+',' ',prod).strip(); tot=int(tot); pay_y=round(tot/12)
        is_drv=bool(re.search(r'운전자(상해|보험)', prod+co))
        if mat[:4]=='9999': gen='비갱신'
        elif is_drv: gen='비갱신'
        else:
            cov_y=int(mat[:4])-int(con[:4])
            gen='갱신' if tot<=240 and pay_y==cov_y else '비갱신'
        out.append(dict(no=int(no),회사=co.strip(),상품=prod,계약일=con,만기일=mat,
                        주기=cyc,회차=f'{paid}/{tot}',납입년=f'{pay_y}년',보험료=int(fee.replace(',','')),갱신=gen))
    return out[:50]
