"""fill_excel.py — 보장분석 JSON → 색칠된 보장진단 엑셀.
법칙: 헤더 갱신=파랑배경/비갱신=빨강배경+흰글자, 값셀 갱신담보=파랑글자/비갱신=검정,
담보명 매칭(템플릿 B열), 미매칭·결합담보는 메모(K열)."""
import re, openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Alignment

BLUE="FF0000FF"; RED="FFFF0000"; BLACK="FF000000"; WHITE="FFFFFFFF"
F_BLUE=PatternFill("solid", fgColor=BLUE)
F_RED =PatternFill("solid", fgColor=RED)

def _norm(s):
    s=re.sub(r"\s+","",str(s or ""))
    return re.sub(r"[()（）·.\-_/]","",s)

def fill_excel(data, template_path, out_path):
    wb=openpyxl.load_workbook(template_path); ws=wb.active
    client=data.get("client","고객")
    ws["A1"]=f"{client}\n보장진단"

    contracts=data.get("contracts",[])[:7]   # C..I = 7열
    # 1) 계약 헤더/메타
    for i,c in enumerate(contracts):
        col=3+i  # C=3
        ren = "비" not in (c.get("renewal","")) and "갱신" in (c.get("renewal",""))
        head=ws.cell(1,col)
        circ="①②③④⑤⑥⑦"[i] if i<7 else ""
        head.value=f"{circ}{c.get('company','')}\n{c.get('product','')}"
        head.fill=F_BLUE if ren else F_RED
        head.font=Font(bold=True, color=WHITE)
        head.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(2,col).value=c.get("renewal","")
        ws.cell(3,col).value=c.get("premium")
        ws.cell(4,col).value=c.get("start") or _split(c.get("period"),0)
        ws.cell(5,col).value=c.get("end")   or _split(c.get("period"),1)
        ws.cell(6,col).value=c.get("payTerm")
        ws.cell(7,col).value=c.get("payCount")

    # 2) 담보명 → 행 인덱스
    name2row={}
    for r in range(8,ws.max_row+1):
        b=ws.cell(r,2).value
        if b: name2row.setdefault(_norm(b), r)

    unmatched=[]
    for rw in data.get("coverage",[]):
        row=name2row.get(_norm(rw.get("name")))
        if not row:   # 완화 매칭(포함)
            key=_norm(rw.get("name"))
            row=next((v for k,v in name2row.items() if key and (key in k or k in key)), None)
        if not row:
            vals=", ".join(f"{a.get('value')}" for a in rw.get("amounts",[]))
            unmatched.append(f"{rw.get('group','')}>{rw.get('name')}={vals}")
            continue
        for a in rw.get("amounts",[]):
            ci=a.get("ci"); 
            if ci is None or ci>=len(contracts): continue
            cell=ws.cell(row,3+ci)
            if isinstance(cell,MergedCell): continue
            cell.value=a.get("value")
            cell.font=Font(color=BLUE if a.get("renew") else BLACK)

    # 3) 메모(K열) — 셀프팩폭/증권확인 + 미매칭
    memos=list(data.get("memo",[]))
    if unmatched: memos.append("미매칭(폼 외 담보): "+" / ".join(unmatched))
    if memos:
        anchor=ws.cell(8,11)  # K8 (K8:K12 병합 앵커)
        if not isinstance(anchor,MergedCell):
            anchor.value="\n".join(f"• {m}" for m in memos)
            anchor.alignment=Alignment(wrap_text=True, vertical="top")
    wb.save(out_path)
    return out_path

def _split(period,idx):
    if not period: return None
    m=re.split(r"[~\-]", str(period).split("/")[0])
    return m[idx].strip() if idx<len(m) else None
