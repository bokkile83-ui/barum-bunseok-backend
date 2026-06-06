# -*- coding: utf-8 -*-
"""BARUM 보장분석 백엔드 v2 — PDF→PyMuPDF추출→캐논분류(색상·CI)→엑셀+PPT(zip)."""
import os, tempfile, datetime, zipfile
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fill_canon import fill
from ppt_gen import gen_ppt
app=FastAPI(title="BARUM 보장분석 v2")
PW=os.environ.get("BARUM_PW","1009"); HERE=os.path.dirname(os.path.abspath(__file__))
TPL=os.path.join(HERE,"template_canon.xlsx")
PAGE="""<!doctype html><html lang=ko><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>BARUM 보장분석</title><style>body{background:#070d18;color:#eaf0f8;font-family:system-ui,'Apple SD Gothic Neo';display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0}
.c{background:linear-gradient(135deg,#0b1f3a,#12294b);border:1px solid #1e3358;border-top:3px solid #c9a14a;border-radius:16px;padding:30px 28px;max-width:390px;width:90%}
h1{font-size:19px;margin:0 0 4px;color:#e6c879}p{color:#9fb0c8;font-size:13px;margin:0 0 18px}
input,button{width:100%;box-sizing:border-box;padding:11px;border-radius:9px;border:1px solid #2a4570;background:#0b1f3a;color:#eaf0f8;margin-top:9px;font-size:14px}
button{background:#c9a14a;color:#0b1f3a;font-weight:800;border:0;cursor:pointer;margin-top:16px}.m{margin-top:12px;font-size:13px;color:#8fe0b0}.e{color:#ff9aa6}</style>
<div class=c><h1>BARUM 보장분석</h1><p>PDF 업로드 → 보장진단 엑셀 + 보장요약 PPT (zip)</p>
<input type=password id=pw placeholder=비밀번호><input type=file id=f accept=.pdf>
<button onclick=go()>분석 → 엑셀+PPT 다운로드</button><div class=m id=msg></div></div>
<script>async function go(){const f=document.getElementById('f').files[0],pw=document.getElementById('pw').value,m=document.getElementById('msg');
if(!f){m.textContent='PDF를 선택하세요';m.className='m e';return}m.textContent='분석 중…(추출+캐논분류+PPT)';m.className='m';
const fd=new FormData();fd.append('file',f);fd.append('pw',pw);const r=await fetch('/analyze',{method:'POST',body:fd});
if(!r.ok){m.textContent='실패: '+(await r.text());m.className='m e';return}
const b=await r.blob(),u=URL.createObjectURL(b),a=document.createElement('a');a.href=u;a.download='보장분석.zip';a.click();m.textContent='완료 — zip 다운로드됨'}</script></html>"""
@app.get("/",response_class=HTMLResponse)
def home(): return PAGE
@app.get("/health")
def health(): return {"ok":True}
@app.post("/analyze")
async def analyze(file:UploadFile=File(...),pw:str=Form("")):
    if pw!=PW: raise HTTPException(401,"비밀번호 오류")
    if not file.filename.lower().endswith(".pdf"): raise HTTPException(400,"PDF만 가능")
    with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as t: t.write(await file.read()); pdf=t.name
    d=tempfile.mkdtemp(); xl=os.path.join(d,"보장진단.xlsx"); pp=os.path.join(d,"보장요약.pptx")
    try:
        info=fill(pdf,TPL,xl); gen_ppt(xl,pp)
    except Exception as e: raise HTTPException(500,f"분석오류: {e}")
    zp=os.path.join(d,f"보장분석_{datetime.datetime.now():%H%M%S}.zip")
    with zipfile.ZipFile(zp,"w") as z: z.write(xl,"보장진단.xlsx"); z.write(pp,"보장요약.pptx")
    return FileResponse(zp,filename="보장분석.zip",media_type="application/zip")
