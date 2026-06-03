"""main.py — 바름 AI 보장분석 백엔드 (단일 Render 서비스).
흐름: PDF 업로드 → Claude(29+4조)→JSON → 엑셀 색칠 채움 → [2단계: PPT 연계] → 파일 다운로드 반환."""
import os, json, base64, uuid, re, tempfile, httpx
from fastapi import FastAPI, UploadFile, File, Header, Request
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from rules import RULES
from fill_excel import fill_excel
try:
    from fill_pptx import fill_pptx
    PPTX_READY = True
except Exception:
    PPTX_READY = False

BASE = os.path.dirname(__file__)
OUT = os.path.join(tempfile.gettempdir(), "barum_out"); os.makedirs(OUT, exist_ok=True)
app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.get("/", response_class=HTMLResponse)
def home():
    with open(os.path.join(BASE, "static", "index.html"), encoding="utf-8") as f:
        return f.read()

ACCESS_PW = os.environ.get("ACCESS_PW")  # 설정 시 잠금 활성화

@app.post("/check")
async def check(request: Request):
    if not ACCESS_PW:
        return JSONResponse({"ok": True, "locked": False})  # 미설정=잠금없음
    try:
        body = await request.json()
    except Exception:
        body = {}
    return JSONResponse({"ok": str(body.get("pw", "")) == ACCESS_PW, "locked": True})

def _gate(x_access_pw):
    return (not ACCESS_PW) or (x_access_pw == ACCESS_PW)

@app.post("/analyze")
async def analyze(file: UploadFile = File(...), x_access_pw: str = Header(None)):
    if not _gate(x_access_pw):
        return JSONResponse({"ok": False, "error": "잠금: 비밀번호가 필요합니다"}, status_code=401)
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return JSONResponse({"ok": False, "error": "ANTHROPIC_API_KEY 미설정 (Render 환경변수 확인)"} )
    pdf = await file.read()
    b64 = base64.b64encode(pdf).decode()
    payload = {
        "model": "claude-sonnet-4-6", "max_tokens": 6000, "system": RULES,
        "tools": [{"type": "web_search_20250305", "name": "web_search", "max_uses": 5}],
        "messages": [{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
            {"type": "text", "text": "이 보장분석지를 법칙대로 분석해 JSON만 출력하세요."}]}],
    }
    try:
        async with httpx.AsyncClient(timeout=280) as cx:
            r = await cx.post("https://api.anthropic.com/v1/messages", json=payload,
                headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
        d = r.json()
        if d.get("error"):
            return JSONResponse({"ok": False, "error": "Claude: " + str(d["error"].get("message"))})
        text = "".join(b.get("text", "") for b in d.get("content", []))
    except Exception as e:
        return JSONResponse({"ok": False, "error": "API 호출 오류: " + str(e)})

    data = _extract_json(text)
    if not data:
        return JSONResponse({"ok": False, "error": "JSON 파싱 실패", "raw": text[:4000]})

    token = uuid.uuid4().hex[:10]
    client = data.get("client", "고객")
    files = []
    # 1) 엑셀 먼저
    xlsx = os.path.join(OUT, f"{token}_보장진단_{client}.xlsx")
    fill_excel(data, os.path.join(BASE, "template.xlsx"), xlsx)
    files.append({"name": f"보장진단_{client}.xlsx", "url": f"/download/{token}/excel"})
    # 2) PPT 연계 (엑셀 값 기반) — 준비되면
    if PPTX_READY and os.path.exists(os.path.join(BASE, "form.pptx")):
        try:
            pptx = os.path.join(OUT, f"{token}_보장분석지_{client}.pptx")
            fill_pptx(data, xlsx, os.path.join(BASE, "form.pptx"), pptx)
            files.append({"name": f"보장분석지_{client}.pptx", "url": f"/download/{token}/pptx"})
        except Exception as e:
            data.setdefault("memo", []).append("PPT 생성 실패(엑셀은 정상): " + str(e))
    return JSONResponse({"ok": True, "client": client,
                         "recommend": data.get("recommend", []), "memo": data.get("memo", []),
                         "files": files, "pptx_ready": PPTX_READY})

@app.get("/download/{token}/{kind}")
def download(token: str, kind: str):
    ext = "xlsx" if kind == "excel" else "pptx"
    for fn in os.listdir(OUT):
        if fn.startswith(token) and fn.endswith(ext):
            return FileResponse(os.path.join(OUT, fn), filename=fn.split("_", 1)[1])
    return JSONResponse({"error": "파일 없음(만료)"}, status_code=404)

def _extract_json(t):
    s = re.sub(r"```json|```", "", t).strip()
    a, b = s.find("{"), s.rfind("}")
    if a < 0 or b < 0: return None
    try: return json.loads(s[a:b+1])
    except Exception: return None
