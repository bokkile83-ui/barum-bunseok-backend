# -*- coding: utf-8 -*-
"""BARUM 보장진단서 PPT v39 — 뼈대 그대로 · 워크시트 흰칸 전부 편집(빈칸 상자 흰색 채우기로 클릭 가능)."""
import os, re, subprocess, tempfile
import xml.etree.ElementTree as ET
from collections import Counter

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

NS = {'x': 'http://www.w3.org/1999/xhtml'}
EMU_PER_PT = 12700
DPI = 200
FONT = "맑은 고딕"


def _setfont(run, name=FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    from pptx.oxml.ns import qn
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", name)


def _valueset(rep):
    V = set()

    def add(x):
        if x is None:
            return
        s = str(x).strip().replace(' ', '')
        if s and s not in ('-', '0'):
            V.add(s)

    add(f"{rep.get('n_contract', 0)}건")
    add(f"{rep.get('premium', 0):,}원")
    add(f"{rep.get('renew', 0)}/{rep.get('nonrenew', 0)}")
    add(f"{rep.get('gap_count', 0)}영역")
    for c in rep.get('coverage', []):
        for i in c.get('items', []):
            add(i.get('v'))
    ci = rep.get('ci', {}) or {}
    for i in ci.get('items', []):
        add(i.get('v'))
    add(ci.get('samang')); add(ci.get('residual'))
    if ci.get('rate'):
        add(f"{ci['rate']}%"); add(f"{100 - ci['rate']}%")
    for c in rep.get('chiryo', []):
        add(c.get('value'))
    for d in rep.get('donuts', []):
        add(f"{d.get('pct')}%")
    for d in rep.get('donut_detail', []):
        add(d.get('have')); add(d.get('rec')); add(f"{d.get('pct')}%")
    for c in rep.get('renew_list', []) + rep.get('nonrenew_list', []):
        add(c.get('v'))
    for i in rep.get('p5_own', []):
        add(i.get('v'))
    add(rep.get('client'))
    add(rep.get('band_label'))
    return V


_WS = set()          # 워크시트 흰칸(wbox) 안 문자열


def _patch_worksheet():
    """워크시트 카드의 흰칸 내용을 수집하고, 빈칸엔 '.' 을 넣는다. report_weasy 원본 무수정."""
    import report_weasy as _rw
    if getattr(_rw, '_barum_ws', False):
        return
    _E = '<span class="wbox"></span>'
    _D = '<span class="wbox">.</span>'
    _RX = re.compile(r'<span class="wbox">(.*?)</span>')
    for _fn in ('_wcard', '_wcard_sj', '_wcard_fix'):
        _o = getattr(_rw, _fn, None)
        if _o is None:
            continue

        def _mk(o):
            def _w(*a, **k):
                h = o(*a, **k).replace(_E, _D)
                for m in _RX.finditer(h):
                    t = m.group(1).strip().replace(' ', '')
                    if t:
                        _WS.add(t)
                return h
            return _w
        setattr(_rw, _fn, _mk(_o))

    # ★_wcard_fix_list(.mb 흰칸) 개방: 빈 .mb → '.' 주입, 채워진 값은 _WS 수집
    _EM = '<span class="mb"></span>'
    _DM = '<span class="mb">.</span>'
    _RXM = re.compile(r'<span class="mb">(.*?)</span>')
    _ol = getattr(_rw, '_wcard_fix_list', None)
    if _ol is not None:
        def _mkl(o):
            def _w(*a, **k):
                h = o(*a, **k).replace(_EM, _DM)
                for m in _RXM.finditer(h):
                    t = m.group(1).strip().replace(' ', '')
                    if t and t != '.':
                        _WS.add(t)
                return h
            return _w
        setattr(_rw, '_wcard_fix_list', _mkl(_ol))

    _rw._barum_ws = True


def _pageset(rep):
    P = set()
    for b in rep.get('premium_bars', []):
        nm = str(b.get('nm', '')).strip().replace(' ', '')
        if nm:
            P.add(nm)
        try:
            P.add(f"{int(b.get('amt', 0)):,}")
        except Exception:
            pass
    # 편집 대상 페이지(0-based): 2=보험료막대 / 7=워크시트 / 8=비갱신 / 9=운전자 / 10=간병 / 11=재가 / 12=실손
    D = {'.'}
    return {2: P, 7: set(_WS) | D, 8: set(_WS) | D, 9: set(_WS) | D,
            10: set(_WS) | D, 11: set(_WS) | D, 12: set(_WS) | D}


def _value_boxes(xml_path, V, PG=None):
    root = ET.parse(xml_path).getroot()
    PG = PG or {}
    out = []
    for pi, pg in enumerate(root.findall('.//x:page', NS)):
        VV = V | PG.get(pi, set()) | {'.'}
        pw, ph = float(pg.get('width')), float(pg.get('height'))
        found = []
        for ln in pg.findall('.//x:line', NS):
            ws = [(w.text or '', float(w.get('xMin')), float(w.get('yMin')),
                   float(w.get('xMax')), float(w.get('yMax')))
                  for w in ln.findall('x:word', NS)]
            n = len(ws); i = 0
            while i < n:
                for L in (8, 7, 6, 5, 4, 3, 2, 1):
                    if i + L > n:
                        continue
                    grp = ws[i:i + L]
                    if ''.join(g[0] for g in grp) in VV:
                        txt = ' '.join(g[0] for g in grp)
                        found.append((txt, grp[0][1], min(g[2] for g in grp),
                                      grp[-1][3], max(g[4] for g in grp)))
                        i += L
                        break
                else:
                    i += 1
        out.append((pw, ph, found))
    return out


def _fg_of(img, bx):
    x0, y0, x1, y1 = bx
    W, H = img.size
    x0, y0 = max(0, x0), max(0, y0); x1, y1 = min(W, x1), min(H, y1)
    if x1 <= x0 or y1 <= y0:
        return (0, 0, 0)
    ring = []
    for xx in range(x0, x1, max(1, (x1 - x0) // 10)):
        for yy in (max(0, y0 - 2), min(H - 1, y1 + 1)):
            ring.append(img.getpixel((xx, yy))[:3])
    bg = Counter(ring).most_common(1)[0][0] if ring else (255, 255, 255)
    px = list(img.crop((x0, y0, x1, y1)).convert('RGB').getdata())
    return max(px, key=lambda c: (c[0]-bg[0])**2 + (c[1]-bg[1])**2 + (c[2]-bg[2])**2)


def _ink(img, bx):
    x0, y0, x1, y1 = bx
    W, H = img.size
    x0, y0 = max(0, x0), max(0, y0); x1, y1 = min(W, x1), min(H, y1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    px = list(img.crop((x0, y0, x1, y1)).convert('L').getdata())
    if not px:
        return 0.0
    px.sort(); bg = px[len(px) // 2]
    return sum(1 for v in px if abs(v - bg) > 60) / len(px)


def _erase(img, bx):
    x0, y0, x1, y1 = bx
    W, H = img.size
    x0, y0 = max(0, x0), max(0, y0); x1, y1 = min(W, x1), min(H, y1)
    if x1 <= x0 or y1 <= y0:
        return
    lx, rx = max(0, x0 - 2), min(W - 1, x1 + 1)
    n = x1 - x0
    for yy in range(y0, y1):
        lc = img.getpixel((lx, yy))[:3]
        rc = img.getpixel((rx, yy))[:3]
        if lc == rc:
            for xx in range(x0, x1):
                img.putpixel((xx, yy), lc)
        else:
            for i, xx in enumerate(range(x0, x1)):
                t = i / max(1, n - 1)
                img.putpixel((xx, yy), tuple(int(lc[k] + (rc[k] - lc[k]) * t) for k in range(3)))


def _txt_w(txt, fs):
    """텍스트 렌더 폭(pt) 추정. 한글=fs*1.0, 숫자/영문/기호=fs*0.55, 공백=fs*0.3"""
    w = 0.0
    for ch in txt:
        o = ord(ch)
        if o > 0x1100 and o < 0xD7A4:      # 한글
            w += fs * 1.0
        elif ch == ' ':
            w += fs * 0.3
        elif ch in ',.·/':
            w += fs * 0.35
        else:                               # 숫자·영문·기타
            w += fs * 0.55
    return w


def build_report_pptx(rep, out, dpi=DPI):
    import report_weasy
    from report_weasy import build_report_pdf
    report_weasy._PPT_MODE = True
    from pdf2image import convert_from_path
    try:
        _WS.clear()
        _patch_worksheet()
    except Exception:
        pass

    tmp = tempfile.mkdtemp()
    pdf = os.path.join(tmp, 'rep.pdf')
    build_report_pdf(rep, pdf)

    try:
        xml = os.path.join(tmp, 'bb.xml')
        subprocess.run(['pdftotext', '-bbox-layout', pdf, xml], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pages = _value_boxes(xml, _valueset(rep), _pageset(rep))
    except Exception:
        pages = []

    imgs = convert_from_path(pdf, dpi=dpi)

    prs = Presentation()
    pw_pt, ph_pt = (pages[0][0], pages[0][1]) if pages else (595.276, 841.890)
    prs.slide_width = Emu(int(pw_pt * EMU_PER_PT))
    prs.slide_height = Emu(int(ph_pt * EMU_PER_PT))
    BL = prs.slide_layouts[6]

    for idx, im in enumerate(imgs):
        im = im.convert('RGB')
        vals = pages[idx][2] if idx < len(pages) else []
        sx = im.width / pw_pt
        sy = im.height / ph_pt

        meta = []
        for txt, x0, y0, x1, y1 in vals:
            bx = (int(x0 * sx) - 1, int(y0 * sy) - 1, int(x1 * sx) + 2, int(y1 * sy) + 1)
            meta.append((txt, x0, y0, x1, y1, _fg_of(im, bx), _ink(im, bx), bx))
        for *_, bx in meta:
            _erase(im, bx)

        ip = os.path.join(tmp, f'p{idx}.png')
        im.save(ip, 'PNG')

        s = prs.slides.add_slide(BL)
        s.shapes.add_picture(ip, 0, 0, prs.slide_width, prs.slide_height)

        for txt, x0, y0, x1, y1, fg, ink, _bx in meta:
            h_pt = y1 - y0
            fs = max(5.0, round(h_pt * 0.90, 1))
            pad = 2.0
            if txt == '.':
                x0 = x1 - 28.0
                fs = 8.5
            else:
                # ★폭 기반 자동 축소: 문자열이 상자폭 넘으면 폰트 줄임(오버 방지)
                box_w = (x1 - x0) + pad * 1.0   # 가용 폭(pt), 좌우 여백 약간 남김
                need = _txt_w(txt, fs)
                if need > box_w and need > 0:
                    fs = max(5.0, round(fs * box_w / need, 1))
            sh = s.shapes.add_textbox(Emu(int((x0 - pad) * EMU_PER_PT)),
                                      Emu(int((y0 - pad * 0.6) * EMU_PER_PT)),
                                      Emu(int((x1 - x0 + pad * 2.4) * EMU_PER_PT)),
                                      Emu(int((h_pt + pad * 1.2) * EMU_PER_PT)))
            tf = sh.text_frame
            tf.word_wrap = False
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if txt == '.' else PP_ALIGN.LEFT
            r = p.add_run(); r.text = txt
            r.font.size = Pt(fs)
            r.font.color.rgb = RGBColor(*fg)
            r.font.bold = (ink > 0.30)
            _setfont(r)
            if txt == '.':
                f = sh.fill
                f.solid()
                f.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                sh.line.fill.background()

    prs.save(out)
    return True
