# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

CATS=[('사망',['일반사망(종신)','질병사망(80세)','질병사망(__세)','상해사망','교통상해사망']),
('후유장애',['상해후유3%','상해후유80%','질병후유3%','질병후유80%','교통후유장해']),
('암',['암진단비','암진단비(갱신)','유사암(갑.기.경.제)','3대암진단비','전이암진단비','항암치료비','표적항암치료비','다빈치로봇수술비','양성자','세기조절','암수술','암주요치료비','하이클래스(비급여주요치료비)','암통원비(상급병원)','암일당','CAR-T세포치료비','13대항암치료비','면역항암치료비','호르몬항암치료비']),
('뇌혈관',['뇌혈관진단비','뇌졸증진단비','뇌졸증진단비(갱신)','뇌출혈진단비','산정특례(뇌혈관)','혈전용해치료비(뇌졸중)','2대주요치료비(뇌)','뇌혈관수술비']),
('심장',['부정맥질환(I49)','심부전','염증','협심증','빈맥','심근병증','심장판막','급성심근경색','허혈성','산정특례(심장)','2대주요치료비(심)','심장수술비']),
('일당',['질병일당','상해일당','질병중환자실','상해중환자실','1인실 상급병원일당','1인실 종합병원일당','간병인사용일당','간병인지원일당(질병)','간병인지원일당(상해)','요양병원일당','간호통합병동일당','암일당(입원)']),
('입원비',['상해입원','질병입원']),
('수술비',['상해수술비','질병수술비','7대질병 수술비','상해종수술비(1-5종)','질병종수술비(1-5종)','n대수술비(대표)','119대 수술비','골절수술비','5대골절수술비','중대상해수술비','화상수술비','창상봉합수술비']),
('운전자',['교통사고처리지원금','6주미만처리지원금','교통사고벌금(대인)','교통사고벌금(대물)','변호사선임비용','자동차부상위로금']),
('배상',['일상생활배상책임']),('골절',['골절진단비','5대골절진단비']),
('화상',['화상진단비','중대화상진단비']),('응급실',['응급실내원비(응급)']),('깁스',['깁스치료비']),
('화재/치아',['화재벌금/화재보험','크라운','임플란트','치아 기타']),
('요양치매',['장기요양자금(1-4급)','경증이상치매진단비','재가급여','시설급여']),
('실손',['실손입원','실손통원','실손 약(처방조제)','MRI','도수치료','비급여주사'])]

BLUE='FF0000FF';RED='FFFF0000';WHITE='FFFFFFFF';BLACK='FF000000';GREEN='FF006400';GRAY='FFEAEAEA'
thin=Side('thin',color='FF999999');med=Side('medium',color='FF333333')
GRID=Border(top=med,bottom=med,left=thin,right=thin)

def build_excel(data, path):
    contracts={c['col']:c for c in data['contracts']}
    cols=[c['col'] for c in data['contracts']]
    colnum={c:i+3 for i,c in enumerate(cols)}
    matrix={}
    for r in data.get('rows',[]):
        for col,amt in r['values'].items():
            matrix.setdefault((r['category'],r['name']),{})[col]=(amt, col in (r.get('renew') or []))
    # 심장 허혈성→급성심근경색·협심증 중복채움
    for col in cols:
        h=matrix.get(('심장','허혈성'),{}).get(col)
        if h:
            for tgt in ['급성심근경색','협심증']:
                k=('심장',tgt); cur=matrix.setdefault(k,{}).get(col)
                if (not cur) or cur[0]<h[0]: matrix[k][col]=(h[0],h[1])
    surg=data.get('surg15',{}); ndae=data.get('ndae',{})

    wb=Workbook(); ws=wb.active; ws.title='고객 보장진단'
    ws['A1']=f"{data.get('customer','고객')} 고객님\n보장진단"; ws.merge_cells('A1:B7')
    ws['A1'].alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); ws['A1'].font=Font(bold=True,size=12)
    for c in cols:
        cn=colnum[c]; info=contracts[c]; gb=info.get('renewal','비갱신')
        completed = False
        try:
            cnt=info.get('count','/'); a,b=cnt.split('/'); completed = a.strip()==b.strip() and a.strip()!=''
        except: pass
        fill = GREEN if completed else (BLUE if gb=='갱신' else RED)
        h=ws.cell(row=1,column=cn,value=info.get('company','')); h.fill=PatternFill('solid',fgColor=fill)
        h.font=Font(color=WHITE,bold=True,size=9); h.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); h.border=GRID
        meta=[gb,info.get('premium',''),info.get('start',''),info.get('end',''),info.get('pay',''),info.get('count','')]
        for i,v in enumerate(meta):
            cc=ws.cell(row=2+i,column=cn,value=v); cc.alignment=Alignment(horizontal='center'); cc.border=GRID; cc.font=Font(size=8)
            if i==1 and isinstance(v,(int,float)): cc.number_format='#,##0'
    for i,lb in enumerate(['갱신구분','보험료','가입일','만기일','납입','회차']): ws.cell(row=2+i,column=17,value=lb).font=Font(size=8,italic=True)
    ws.cell(row=1,column=17,value='합계').font=Font(bold=True); ws.cell(row=1,column=18,value='메모').font=Font(bold=True)
    for rr in range(1,8): ws.cell(row=rr,column=17).border=GRID; ws.cell(row=rr,column=18).border=GRID

    def renew_of(c): return contracts[c].get('renewal')=='갱신'
    r=8
    for cat,rows in CATS:
        extra=sorted({k[1] for k in matrix if k[0]==cat and k[1] not in rows})
        allrows=rows+extra; start=r
        for rw in allrows:
            ws.cell(row=r,column=2,value=rw).font=Font(size=9); ws.cell(row=r,column=2).border=GRID; ws.cell(row=r,column=1).border=GRID
            for c in cols: ws.cell(row=r,column=colnum[c]).border=GRID
            ws.cell(row=r,column=18).border=GRID
            q=ws.cell(row=r,column=17); q.border=GRID; q.font=Font(bold=True,size=8); q.number_format='#,##0'
            if rw in ('상해종수술비(1-5종)','질병종수술비(1-5종)'):
                kind='상해' if '상해' in rw else '질병'; tot=[0]*5
                kk=surg.get(kind,{})
                for c in cols:
                    seq=kk.get(c)
                    if seq:
                        for j in range(5): tot[j]+=seq[j]
                        cell=ws.cell(row=r,column=colnum[c],value='/'.join(f'{x:,}' for x in seq))
                        cell.font=Font(color=(BLUE if renew_of(c) else BLACK),size=8); cell.alignment=Alignment(horizontal='center',shrink_to_fit=True)
                if any(tot): q.value='/'.join(f'{x:,}' for x in tot); q.number_format='General'
            elif rw=='n대수술비(대표)':
                for c,a in ndae.items():
                    if c in colnum:
                        cell=ws.cell(row=r,column=colnum[c],value=a); cell.number_format='#,##0'
                        cell.font=Font(color=(BLUE if renew_of(c) else BLACK),size=8); cell.alignment=Alignment(horizontal='center')
                if ndae: q.value=f'=SUM(C{r}:P{r})'
            else:
                has=False
                for c in cols:
                    v=matrix.get((cat,rw),{}).get(c)
                    if v:
                        has=True; amt,rn=v
                        cell=ws.cell(row=r,column=colnum[c],value=amt); cell.number_format='#,##0'
                        cell.font=Font(color=(BLUE if (renew_of(c) or rn) else BLACK),size=8); cell.alignment=Alignment(horizontal='center')
                if has: q.value=f'=SUM(C{r}:P{r})'
            r+=1
        ws.merge_cells(start_row=start,start_column=1,end_row=r-1,end_column=1)
        a=ws.cell(row=start,column=1,value=cat); a.font=Font(bold=True,size=9)
        a.fill=PatternFill('solid',fgColor=GRAY); a.alignment=Alignment(horizontal='center',vertical='center'); a.border=GRID
    ws.column_dimensions['A'].width=6; ws.column_dimensions['B'].width=20
    for c in cols: ws.column_dimensions[ws.cell(row=1,column=colnum[c]).column_letter].width=7.2
    ws.column_dimensions['Q'].width=12; ws.column_dimensions['R'].width=8
    # 메모시트
    ms=wb.create_sheet('📋확인사항'); ms['A1']=f"{data.get('customer','고객')} — 확인사항"; ms['A1'].font=Font(bold=True,size=11)
    ms.append([]); ms.append(['구분','대상','내용'])
    for cc in ms[3]: cc.font=Font(bold=True); cc.fill=PatternFill('solid',fgColor=GRAY)
    for m in data.get('memo',[]): ms.append([m.get('type',''),m.get('target',''),m.get('note','')])
    ms.column_dimensions['A'].width=12; ms.column_dimensions['B'].width=40; ms.column_dimensions['C'].width=50
    wb.save(path)
    return path
