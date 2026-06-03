# -*- coding: utf-8 -*-
import os, io, json, base64, zipfile, re, tempfile, traceback
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, HTMLResponse
import anthropic
import rules, fill_excel, fill_pptx

app=FastAPI()
ACCESS_PW=os.environ.get('ACCESS_PW','1009')
MODEL=os.environ.get('MODEL','claude-sonnet-4-6')
HERE=os.path.dirname(os.path.abspath(__file__))

def client():
    key=os.environ.get('ANTHROPIC_API_KEY')
    if not key:
        raise RuntimeError('ANTHROPIC_API_KEY 환경변수가 없습니다. Railway → Variables에 추가하세요.')
    return anthropic.Anthropic(api_key=key)

@app.get('/health')
def health():
    return {'ok':True,'model':MODEL,'has_key':bool(os.environ.get('ANTHROPIC_API_KEY'))}

@app.get('/', response_class=HTMLResponse)
def home():
    return open(os.path.join(HERE,'static','index.html'),encoding='utf-8').read()

@app.post('/analyze')
async def analyze(password: str = Form(...), file: UploadFile = File(...)):
    try:
        if password.strip()!=ACCESS_PW:
            return JSONResponse({'ok':False,'error':'비밀번호 오류 (기본 1009)'})
        pdf=await file.read()
        if not pdf:
            return JSONResponse({'ok':False,'error':'PDF가 비어있습니다.'})
        b64=base64.standard_b64encode(pdf).decode()
        msg=client().messages.create(model=MODEL, max_tokens=16000, system=rules.SYSTEM_PROMPT,
            messages=[{'role':'user','content':[
                {'type':'document','source':{'type':'base64','media_type':'application/pdf','data':b64}},
                {'type':'text','text':'이 PDF를 법칙대로 계약별로 분석해 JSON만 출력.'}]}])
        raw=''.join(b.text for b in msg.content if getattr(b,'type',None)=='text')
        m=re.search(r'\{.*\}', raw, re.S)
        if not m:
            return JSONResponse({'ok':False,'error':'분석 JSON을 못 찾음','raw':raw[:1200]})
        data=json.loads(m.group(0))
        cust=re.sub(r'[^\w가-힣]','',str(data.get('customer','고객'))) or '고객'
        xlsx=os.path.join(tempfile.gettempdir(),f'보장진단_{cust}.xlsx')
        fill_excel.build_excel(data, xlsx)
        pptx=None
        try:
            pptx=os.path.join(tempfile.gettempdir(),f'보장분석지_{cust}.pptx')
            fill_pptx.fill_pptx(data, os.path.join(HERE,'form.pptx'), pptx)
        except Exception:
            pptx=None
        buf=io.BytesIO()
        with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
            z.write(xlsx, os.path.basename(xlsx))
            if pptx: z.write(pptx, os.path.basename(pptx))
            z.writestr('분석결과.json', json.dumps(data,ensure_ascii=False,indent=2))
        return JSONResponse({'ok':True,'data':data,
            'zip_b64':base64.b64encode(buf.getvalue()).decode(),
            'filename':f'barum_{cust}.zip',
            'pptx_filled':bool(pptx)})
    except anthropic.APIStatusError as e:
        return JSONResponse({'ok':False,'error':f'Anthropic API 오류 {e.status_code}: {e.message}. 모델명({MODEL})/키 확인.'})
    except Exception as e:
        return JSONResponse({'ok':False,'error':str(e),'trace':traceback.format_exc()[-1500:]})

@app.post('/chat')
async def chat(req: dict):
    try:
        data=req.get('data',{}); messages=req.get('messages',[])
        ctx=json.dumps(data,ensure_ascii=False)[:12000]
        sysp=rules.CHAT_PROMPT.replace('{{DATA}}',ctx)
        msg=client().messages.create(model=MODEL, max_tokens=1500, system=sysp, messages=messages)
        out=''.join(b.text for b in msg.content if getattr(b,'type',None)=='text')
        return {'ok':True,'reply':out}
    except Exception as e:
        return {'ok':False,'error':str(e)}
