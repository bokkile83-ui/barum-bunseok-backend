# -*- coding: utf-8 -*-
"""BARUM 보장진단서 PPT — 보장설명서 PDF와 100% 일치(페이지=슬라이드 이미지).
   지점장 지시(2026.07.07): 기존 run-injection PPT 폐기. PDF 디자인을 그대로 슬라이드화.
   각 PDF 페이지를 고해상 이미지로 렌더 → A4 세로 슬라이드에 꽉 채움 → 시각 100% 일치.
"""
import os, tempfile
from pptx import Presentation
from pptx.util import Emu

def _a4_emu():
    # A4 세로 210×297mm → EMU (1in=914400EMU, 1in=25.4mm)
    return Emu(int(210/25.4*914400)), Emu(int(297/25.4*914400))

def build_report_pptx(rep, out, dpi=150):
    """rep으로 보장설명서 PDF를 만든 뒤, 그 페이지들을 슬라이드 이미지로 넣어 PDF와 100% 일치시킨다."""
    from report_weasy import build_report_pdf
    from pdf2image import convert_from_path

    tmpdir = tempfile.mkdtemp()
    pdf_path = os.path.join(tmpdir, 'rep.pdf')
    build_report_pdf(rep, pdf_path)

    imgs = convert_from_path(pdf_path, dpi=dpi)

    prs = Presentation()
    w, h = _a4_emu()
    prs.slide_width = w
    prs.slide_height = h
    blank = prs.slide_layouts[6]  # 빈 레이아웃

    for idx, im in enumerate(imgs):
        slide = prs.slides.add_slide(blank)
        ip = os.path.join(tmpdir, f'p{idx}.png')
        im.save(ip, 'PNG')
        slide.shapes.add_picture(ip, 0, 0, width=w, height=h)

    prs.save(out)
    return True

if __name__ == '__main__':
    import sys
    from coverage_benchmark import map_excel_to_report
    if len(sys.argv) < 2:
        print('사용법: python report_pptx.py <엑셀> [고객명]'); sys.exit(1)
    _rep = map_excel_to_report(sys.argv[1], settings={'client': sys.argv[2] if len(sys.argv) > 2 else '고객',
        'branch': '온빛센터 바름지점', 'manager': '최은혜', 'title': '지점장', 'phone': ''})
    build_report_pptx(_rep, f'보장진단서_{sys.argv[2] if len(sys.argv)>2 else "고객"}.pptx')
    print('진단서 PPT 생성 완료')
