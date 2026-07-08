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
    if
