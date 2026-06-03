# -*- coding: utf-8 -*-
import os, io, json, base64, zipfile, re, tempfile
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
import anthropic
import rules, fill_excel, fill_pptx

app=FastAPI()
ACCESS_PW=os.environ.get('ACCESS_PW','1009')
MODEL=os.environ.get('MODEL','claude-sonnet-4-6')
HERE=os.path.dirname(__file__)

@app.get('/', response_class=HTMLResponse)
def home():
    with open(os.path.join(HERE,'static','index.html'),encoding='utf-8') as f: return f.read()

@app.post('/analyze')
async def analyze(password: str = Form(...), file: UploadFile = File(...)):
    if password.strip()!=ACCESS_PW:
        raise HTTPException(401,'비밀번호 오류')
    pdf=await file.read()
    b64=base64.standard_b64encode(pdf).decode()
    client=anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])
    msg=client.messages.create(model=MODEL, max_tokens=8000, system=rules.SYSTEM_PROMPT,
        messages=[{'role':'user','content':[
            {'type':'document','source':{'type':'base64','media_type':'application/pdf','data':b64}},
            {'type':'text','text':'이 PDF를 법칙대로 계약별로 분석해 JSON만 출력.'}]}])
    raw=''.join(b.text for b in msg.content if b.type=='text')
    m=re.search(r'\{.*\}', raw, re.S)
    if not m: raise HTTPException(500,'분석 JSON 파싱 실패: '+raw[:300])
    data=json.loads(m.group(0))
    cust=re.sub(r'[^\w가-힣]','',data.get('customer','고객')) or '고객'
    xlsx=os.path.join(tempfile.gettempdir(),f'보장진단_{cust}.xlsx')
    pptx=os.path.join(tempfile.gettempdir(),f'보장분석지_{cust}.pptx')
    fill_excel.build_excel(data, xlsx)
    try: fill_pptx.fill_pptx(data, os.path.join(HERE,'form.pptx'), pptx)
    except Exception as e: pptx=None
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        z.write(xlsx, os.path.basename(xlsx))
        if pptx: z.write(pptx, os.path.basename(pptx))
        z.writestr('분석결과.json', json.dumps(data,ensure_ascii=False,indent=2))
    buf.seek(0)
    return StreamingResponse(buf, media_type='application/zip',
        headers={'Content-Disposition':f'attachment; filename=barum_{cust}.zip'})
