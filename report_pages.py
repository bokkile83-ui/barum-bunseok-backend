# -*- coding: utf-8 -*-
"""★★★★★ report_pages.py — 리모델링 리포트 7쪽 (지점장 시안 정본 2026.08.15)
# 각인: v441-brand-20260817

   지점장이 시안(HTML)을 7장 다 주셨다. <b>흉내내지 않고 그 시안을 렌더한다.</b>
   ㆍ값은 전부 엑셀에서 온다 — 하드코딩 0건(뮤테이션 테스트 실패 0건으로 확인)
   ㆍPPT 도형으로 그리던 구판(build_pptx)은 <b>폐기</b>했다. 뷰어가 다시 그려 서식을 버렸다.
   ㆍ진단서 HTML을 빌려 쓰던 경로도 폐기했다.
   build(cmp_, client, base_date) → [7쪽 HTML]
"""
import io


# ═══════════════════ 1 쪽 ═══════════════════


CSS1 = """
@page{size:A4 portrait;margin:0}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:#1c2430}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.tbar{position:absolute;left:0;top:0;width:210mm;height:14mm;background:#0b2340}
.tgold{position:absolute;left:0;top:14mm;width:210mm;height:2.4mm;background:#c5a052}
.fbar{position:absolute;left:0;bottom:0;width:210mm;height:6mm;background:#0b2340}
.body{position:absolute;left:16mm;right:16mm;top:30mm}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:13pt;font-weight:800;color:#0b2340;letter-spacing:.18em}
.bln{width:24mm;height:1.6mm;background:#c5a052;margin-top:3mm}
.mark{position:absolute;right:0;top:14mm;font-size:44pt;font-weight:900;color:#f2f4f7;letter-spacing:.06em}
.eyebrow{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:26mm;font-size:9.5pt;font-weight:800;color:#9c7c32;letter-spacing:.28em}
.title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:4mm;font-size:44pt;font-weight:900;color:#0b2340;line-height:1.14;letter-spacing:-.03em}
.rule{width:34mm;height:1.8mm;background:#c5a052;margin-top:7mm}
.sub{margin-top:5mm;font-size:11pt;font-weight:700;color:#0b2340}
.namebox{margin-top:11mm;background:#f4f7fb;border-left:1.4mm solid #0b2340;padding:7mm 8mm;
display:flex;align-items:baseline;gap:6mm}
.namebox .nm{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:40pt;font-weight:900;color:#0b2340}
.namebox .sfx{font-size:13pt;font-weight:800;color:#9c7c32}
.stats{display:flex;gap:5mm;margin-top:9mm}
.stat{flex:1;border-top:1.4mm solid #c5a052;border-left:.3mm solid #c9d2dc;
border-right:.3mm solid #c9d2dc;border-bottom:.3mm solid #c9d2dc;padding:4mm 5mm 5mm}
.stat .k{font-size:8.5pt;font-weight:800;color:#33404f}
.stat .v{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3mm;font-size:20pt;font-weight:900;color:#06203f}
.inbox{margin-top:9mm;border:.3mm solid #c9d2dc;padding:5mm 6mm}
.inbox .t{font-size:8.5pt;font-weight:900;color:#9c7c32}
.inbox .l{margin-top:3mm;font-size:10.5pt;font-weight:800;color:#0b2340;line-height:1.8}
.hr{margin-top:12mm;border-top:.2mm solid #d9dee6}
.foot{margin-top:4mm;text-align:center;font-size:9.5pt;font-weight:800;color:#0b2340}
"""


def p1(client='고객', base_date='', pg=1, totpg=7, cmp_=None):
    """★표지 — 지점장 선택본(2026.08.16). 곡선·방패 시안은 버리고 이 형태로 간다.
       제목은 「MAKEONE LIFE PLAN」(지점장 지시)."""
    c = cmp_ or {}
    old = int(c.get('prem_old', 0)); new = int(c.get('prem_new', 0)); sv = new - old
    up = len(c.get('up', [])); add = len(c.get('add', []))
    dn = len(c.get('down', [])); de = len(c.get('delete', []))
    n_old = len(c.get('old', {}).get('contracts', [])); n_new = len(c.get('new', {}).get('contracts', []))
    st = [('기존 월 보험료', '%s 원' % format(old, ',')),
          ('변경 후 월 보험료', '%s 원' % format(new, ',')),
          ('월 %s액' % ('증가' if sv > 0 else '절감'), '%s 원' % format(abs(sv), ','))]
    sh = ''.join('<div class="stat"><div class="k">%s</div><div class="v">%s</div></div>' % t for t in st)
    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<div class="page"><div class="tbar"></div><div class="tgold"></div>'
            '<div class="body">'
            '<div class="mark">MAKEONE</div>'
            '<div class="brand">M A K E O N E</div><div class="bln"></div>'
            '<div class="eyebrow">INSURANCE REMODELING</div>'
            '<div class="title">MAKEONE<br>LIFE PLAN</div>'
            '<div class="rule"></div>'
            '<div class="sub">보험료 변화와 보장 비교 · 제안 기준일 %s</div>'
            '<div class="namebox"><span class="nm">%s</span><span class="sfx">고객님</span></div>'
            '<div class="stats">%s</div>'
            '<div class="inbox"><div class="t">이 리포트에 담긴 것</div><div class="l">'
            '계약별 전 · 후 비교　%d건 → %d건<br>'
            '보장 증가 %d항목　·　보장 감소 %d항목<br>'
            '신규 특약 %d개　·　삭제 특약 %d개</div></div>'
            '<div class="hr"></div>'
            '<div class="foot">MAKEONE · 보장분석 자동화 리포트</div>'
            '</div><div class="fbar"></div></div></body></html>'
            % (CSS1, base_date, client, sh, n_old, n_new, up, dn, add, de))



CSS2 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--gold:#c88d20;--gold2:#e6b74d;--green:#14945e;
--softBlue:#f4f8fd;--softGreen:#f1faf5;--line:#dde4ec;--text:#132b48;--muted:#7a8796}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text)}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.top-curve{display:none}

.header{position:relative;z-index:2;padding:8.4mm 10.4mm 4.4mm;display:flex;justify-content:space-between}
.header>div:first-child{width:118mm;flex:none}   /* ★제목이 3줄로 접히던 원인 */
.eyebrow{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:800;letter-spacing:.04em}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.4mm 0 3.6mm;color:#06203f;font-weight:900;font-size:21pt;white-space:nowrap;line-height:1.16;letter-spacing:-.04em}
.title-line{width:100mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 24%,var(--navy) 24%)}
.client{display:none}
.client strong{font-size:12pt;display:block;margin-bottom:1.6mm}
.client span{display:block;font-size:8pt;margin-top:1mm;color:#7a8796}
.client .step{color:#ffcf5c;font-weight:800;margin-top:2mm}
.content{padding:0 10mm 17mm}
.summary-card{margin-top:1.6mm;border:.5pt solid #c9d2dc;border-radius:3.2mm;padding:5.2mm 6mm;
display:flex}
.summary-col{flex:1;min-height:34mm;padding:2.4mm 5mm}
.summary-col:first-child{border-right:.4pt solid #d7dde5}
.summary-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:#06203f;font-size:14pt;font-weight:900}
.summary-sub{margin-top:2.8mm;font-size:9.5pt;color:#33404f}
.big-gold{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3.6mm;color:#b07d0e;font-size:34pt;font-weight:900}
.big-navy{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3.6mm;color:#06203f;font-size:32pt;font-weight:900}
.summary-note{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3.8mm;font-size:12pt;font-weight:900;color:#06203f}
.summary-gold{margin-top:4.8mm;color:var(--gold);font-size:11.5pt;font-weight:900}
.section-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:6.8mm 4mm 3.4mm;color:#06203f;font-size:15pt;font-weight:900;border-left:1.6mm solid #c5a052;padding-left:3mm}
.compare{display:flex;gap:2.4mm;align-items:center;padding:0 4mm}
.premium-card{flex:1;min-height:30mm;padding:3.2mm 4.4mm;border-radius:2.6mm}
.premium-card.before{background:var(--softBlue);border:.4pt solid #b8cce3}
.premium-card.after{background:var(--softGreen);border:.4pt solid #acd4ba}
.premium-card b{display:block;font-size:10pt;margin-bottom:2.4mm}
.premium-card.after b{color:var(--green)}
.insurers{font-size:8pt;color:#33404f}
.premium-number{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:5.2mm;font-size:22pt;font-weight:900;color:var(--navy)}
.after .premium-number{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--green)}
.arrow{width:11mm;flex:none;text-align:center;color:var(--gold);font-size:22pt}
.contract-head{display:flex;justify-content:space-between;align-items:flex-end;margin:5.4mm 4mm 2mm}
.contract-head h3{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:0;font-size:15pt;color:#06203f;font-weight:900;border-left:1.6mm solid #c5a052;padding-left:3mm}
.contract-head span{font-size:8.5pt;color:#33404f;font-weight:700}
.report-table{width:calc(100% - 8mm);margin:0 4mm;border-collapse:collapse;font-size:8.5pt}
.report-table th{padding:2.2mm 1.8mm;text-align:left;color:#b57a12;border-bottom:.4pt solid #d8dde4}
.report-table td{padding:2.2mm 1.8mm;border-bottom:.4pt solid #e7ebef}
.report-table tbody tr:nth-child(odd) td{background:#eef3f9}
.report-table tbody tr.new td{background:#d7f0e2;font-weight:800}
.report-table tbody tr.del td{background:#fbdedb;font-weight:800}
.report-table th:nth-child(3),.report-table th:nth-child(4),
.report-table td:nth-child(3),.report-table td:nth-child(4){text-align:right}
.report-table td:nth-child(3){color:#4b5a6b}
.report-table td:nth-child(4){font-weight:800;color:var(--navy)}
.status-new{color:var(--green);font-weight:900}
.total-row td{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:10.5pt;font-weight:900;padding-top:3.4mm;color:var(--navy)!important}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;
background:linear-gradient(90deg,#082d59,#123f75);color:#fff;display:flex;align-items:center;
justify-content:space-between;padding:0 10mm;font-size:8.5pt}
.footer strong{color:#ffcb57;font-size:9.5pt}
"""


def _m2(v):
    return format(int(v or 0), ',') + '원'


def p2(cmp_, client='고객', base_date='', pg=2, totpg=7):
    sv = cmp_['save_m']
    up = sv < 0
    kl = {(c['company'], c['product']) for c in cmp_['kill']}
    rows = [(c, 'keep') for c in cmp_['old']['contracts']] + [(c, 'new') for c in cmp_['prop']]
    trs = ''
    for c, kind in rows:
        if kind == 'new':
            bf, af, tg, cls = 0, c['premium'], '신규', 'status-new'
        elif (c['company'], c['product']) in kl:
            bf, af, tg, cls = c['premium'], 0, '삭제', 'status-new'
        else:
            bf = af = c['premium']
            tg, cls = '유지', ''
        trs += ('<tr class="%s"><td>%s</td><td>%s</td><td>%s</td>'
                '<td>%s <span class="%s">%s</span></td></tr>'
                % ('new' if tg == '신규' else 'del' if tg == '삭제' else '',
                   c['company'], c['product'][:30],
                   _m2(bf) if bf else '-', _m2(af) if af else '-', cls, tg))

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<article class="page"><div class="top-curve"></div>'
            '<header class="header"><div>'
            '<div class="eyebrow">MAKEONE · PREMIUM CHANGE</div>'
            '<h1>보험료 변화와<br>최종 리포트 금액</h1>'
            '<div class="title-line"></div></div>'
            '<div class="client"><strong>%s 고객님</strong>'
            '<span>제안 기준일 %s</span>'
            '<span class="step">02 보험료 · 금액</span></div></header>'
            '<main class="content">'
            '<section class="summary-card">'
            '<div class="summary-col"><div class="summary-title">최종 리포트 금액</div>'
            '<div class="summary-sub">변경 후 월 보험료</div>'
            '<div class="big-gold">%s</div>'
            '<div class="summary-note">월 %s %s (%s%%)</div></div>'
            '<div class="summary-col"><div class="summary-title">1년 %s</div>'
            '<div class="summary-sub">12개월 기준</div>'
            '<div class="big-navy">%s</div>'
            '<div class="summary-gold">보유 %d건 → 최종 %d건</div></div>'
            '</section>'
            '<div class="section-title">월 보험료 변경</div>'
            '<section class="compare">'
            '<div class="premium-card before"><b>기존 보험료</b>'
            '<div class="insurers">%s</div><div class="premium-number">%s</div></div>'
            '<div class="arrow">&#8594;</div>'
            '<div class="premium-card after"><b>변경 후 총 보험료</b>'
            '<div class="insurers">%s</div><div class="premium-number">%s</div></div>'
            '</section>'
            '<div class="contract-head"><h3>계약별 전 · 후</h3>'
            '<span>제안 %s  +  삭제 후 %s</span></div>'
            '<table class="report-table"><thead><tr>'
            '<th style="width:18%%">보험사</th><th style="width:42%%">상품명</th>'
            '<th style="width:18%%">전 (기존)</th><th style="width:22%%">후 (변경 후) 상태</th>'
            '</tr></thead><tbody>%s</tbody>'
            '<tfoot><tr class="total-row"><td colspan="2">합계</td><td>%s</td><td>%s</td></tr></tfoot>'
            '</table></main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS2, client, base_date,
               _m2(cmp_['prem_new']), _m2(abs(sv)), '증가' if up else '절감', abs(cmp_['save_pct']),
               '추가 부담액' if up else '예상 절감액', _m2(abs(cmp_['save_y'])),
               len(cmp_['old']['contracts']), len(cmp_['new']['contracts']),
               ' + '.join(dict.fromkeys(c['company'] for c in cmp_['old']['contracts']))[:44],
               _m2(cmp_['prem_old']),
               ' + '.join(dict.fromkeys(c['company'] for c in cmp_['new']['contracts']))[:44],
               _m2(cmp_['prem_new']),
               _m2(cmp_['prem_prop']), _m2(cmp_['prem_keep']),
               trs, _m2(cmp_['prem_old']), _m2(cmp_['prem_new']),
               client, pg, totpg))

# ═══════════════════ 3 쪽 ═══════════════════


CSS3 = """
@page{size:A4;margin:0}
:root{--navy:#0b2d58;--navy2:#123f79;--gold:#c79532;--gold-light:#e7c274;--green:#159263;
--red:#cc5656;--text:#17273c;--gray:#33404f;--line:#e5e9ef;--soft:#f5f8fb}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text)}
.report{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.top-shape{display:none}

.header{position:relative;z-index:2;display:flex;justify-content:space-between;
align-items:flex-start;padding:6mm 13mm 2mm}
.brand-line{font-size:9.5pt;color:var(--gold);font-weight:800;letter-spacing:.08em}
h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:1.5mm 0 2mm;color:#06203f;font-size:23pt;font-weight:900;line-height:1.18;letter-spacing:-.04em}
.gold-line{display:flex;gap:1.5mm}
.gold-line span:first-child{width:23mm;height:1.8mm;background:var(--gold);border-radius:5mm}
.gold-line span:last-child{width:12mm;height:1.8mm;background:#edd9b2;border-radius:5mm}
.customer{display:none}
.customer strong{display:block;font-size:13.5pt;margin-bottom:2mm}
.customer span{display:block;margin-top:1.2mm;font-size:8.5pt}
.customer .step{color:#ffd478;font-weight:800;margin-top:2.5mm}
.content{padding:1mm 13mm 19mm}
.panel{border:.9pt solid #b9c6d6;background:#fff;border-radius:4mm;margin-bottom:2.8mm;overflow:hidden;box-shadow:0 .6mm 1.4mm rgba(11,45,89,.10)}
.panel-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;display:inline-block;padding:2.6mm 6mm;
border-radius:0 0 3mm 0;color:#fff;background:linear-gradient(90deg,var(--navy),var(--navy2));
font-size:13pt;font-weight:900}
.number{display:inline-block;width:5.4mm;height:5.4mm;line-height:5.4mm;text-align:center;
border-radius:50%;color:var(--navy);background:#fff;font-size:8.5pt;font-weight:900;margin-right:1.5mm}
.table-wrap{padding:.6mm 5mm 1.4mm}
table{width:100%;border-collapse:collapse;font-size:10pt}
th,td{padding:.5mm 3mm;border-bottom:.4pt solid #edf0f4;text-align:right}
th{color:#96671e;background:#fffaf2;font-weight:800;font-size:9pt}
th:first-child,td:first-child{text-align:left}
tbody tr:last-child td{border-bottom:0}
.old{color:#4b5a6b}
.new{color:var(--navy);font-weight:800}
.diff{color:var(--green);font-weight:800}
.circle-icon{display:inline-block;width:7mm;height:7mm;line-height:7mm;text-align:center;
margin-right:2mm;border-radius:50%;background:#fff6e7;color:var(--gold);font-size:9pt;font-weight:900}
.chart{padding:1.6mm 7mm 2mm}
.legend{text-align:right;margin-bottom:1mm;color:#738095;font-size:8.5pt}
.legend span{margin-left:4mm}
.legend i{display:inline-block;width:3mm;height:3mm;margin-right:1.2mm}
.legend .gray{background:#d8dee7}
.legend .navy{background:var(--navy)}
.bar-row{display:flex;align-items:center;gap:4mm;margin:0.5mm 0}
.bar-label{width:42mm;font-weight:700;font-size:9.5pt;flex:none}
.bar-group{flex:1}
.bar-line{display:flex;align-items:center;gap:2mm;height:3.4mm}
.bar{height:3.4mm;min-width:.5mm}
.bar-before{background:#d8dee7}
.bar-after{background:linear-gradient(90deg,#0b2d58,#154579)}
.value{font-size:8.5pt;color:#33404f;white-space:nowrap}
.value.after{color:var(--navy);font-weight:800}
.premium-wrap{display:flex;gap:6mm;align-items:center;padding:2mm 7mm 2mm}
.premium-left{flex:1}
.premium-row{display:flex;align-items:center;gap:3mm;margin:1.6mm 0}
.premium-row b{width:22mm;font-size:11pt;flex:none}
.premium-barwrap{flex:1}
.premium-bar{height:7mm;background:#d8dee7}
.premium-bar.after{background:linear-gradient(90deg,#0b2d58,#164679)}
.premium-value{width:26mm;text-align:right;font-weight:800;font-size:10.5pt;flex:none}
.increase-box{width:40mm;flex:none;text-align:center;border:.4pt solid #dde4ec;border-radius:4mm;
padding:4mm 2mm;background:linear-gradient(180deg,#f6fbff,#fff)}
.arrow{display:none}
.increase-box strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;display:block;margin-top:2mm;color:var(--gold);font-size:19pt}
.increase-box small{display:block;margin-top:1.2mm;color:var(--navy);font-weight:700;font-size:8.5pt}
.change-wrap{display:flex;gap:6mm;align-items:center;padding:1.5mm 7mm 2mm}
.change-left{flex:1}
.change-row{display:flex;align-items:center;gap:2.5mm;margin:1mm 0}
.change-row b{width:26mm;font-size:9.5pt;flex:none}
.change-barwrap{flex:1}
.change-bar{height:4mm;background:var(--navy);min-width:.4mm}
.change-bar.green{background:var(--green)}
.change-bar.gray{background:#d8dee5}
.change-num{width:12mm;font-size:9pt;font-weight:700;flex:none}
.summary{width:56mm;flex:none;text-align:center;padding:2mm 4mm;border-radius:7mm;
border:.4pt solid #d9dfe7;background:radial-gradient(circle at 50% 10%,#fff,#f3f7fb)}
.shield{width:7mm;height:8.4mm;margin:0 auto .6mm;position:relative;
background:linear-gradient(180deg,#deb354,#96670d);
clip-path:polygon(50% 0,92% 18%,82% 70%,50% 100%,18% 70%,8% 18%)}
.shield span{position:absolute;left:0;right:0;top:1.4mm;color:#fff;font-size:10pt;font-weight:900}
.summary-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--navy);font-size:11.5pt;font-weight:800}
.summary-number{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:24pt;line-height:1;color:var(--navy);font-weight:900;margin:1mm 0 1.5mm}
.summary-text{font-size:8.5pt;color:#33404f}
.footer{position:absolute;left:0;right:0;bottom:0;height:14mm;padding:0 13mm;display:flex;
align-items:center;justify-content:space-between;
background:linear-gradient(90deg,#092b55,#154477);color:#fff;font-size:9pt}
.footer strong{color:#ffd26f}
"""

ICONS3 = ['♥', '◈', '✚', '▦', '✥', '▣', '◆', '●']


def _mw(v):
    v = float(v or 0)
    return ('%s만원' % format(int(v), ',')) if v else '0만원'


def p3(cmp_, client='고객', base_date='', pg=3, totpg=7):
    rw = (cmp_['up'] + cmp_['add'])[:6]
    vmax = max([max(o, n) for _n, o, n, _d in rw] + [1])
    sv = cmp_['save_m']
    up = sv < 0
    pmax = max(cmp_['prem_old'], cmp_['prem_new'], 1)
    cnt = [('보장 증가', len(cmp_['up']), ''), ('신규 추가', len(cmp_['add']), 'green'),
           ('보장 감소', len(cmp_['down']), 'gray'), ('삭제', len(cmp_['delete']), 'gray')]
    cmax = max([c[1] for c in cnt] + [1])

    trs = ''
    for i, (nm, o, n, d) in enumerate(rw):
        trs += ('<tr><td><span class="circle-icon">%s</span>%s</td>'
                '<td class="old">%s</td><td class="new">%s</td>'
                '<td class="diff">%s%s</td></tr>'
                % (ICONS3[i % len(ICONS3)], nm, _mw(o), _mw(n),
                   '+' if d > 0 else '', _mw(abs(d))))

    bars = ''
    for nm, o, n, d in rw:
        bars += ('<div class="bar-row"><div class="bar-label">%s</div><div class="bar-group">'
                 '<div class="bar-line"><div class="bar bar-before" style="width:%.1f%%"></div>'
                 '<span class="value">%s</span></div>'
                 '<div class="bar-line"><div class="bar bar-after" style="width:%.1f%%"></div>'
                 '<span class="value after">%s</span></div></div></div>'
                 % (nm, max(o / vmax * 88, 0.2), format(int(o), ','),
                    max(n / vmax * 88, 0.2), format(int(n), ',')))

    chg = ''
    for lb, v, cls in cnt:
        chg += ('<div class="change-row"><b>%s</b><div class="change-barwrap">'
                '<div class="change-bar %s" style="width:%.1f%%"></div></div>'
                '<span class="change-num">%d개</span></div>'
                % (lb, cls, max(v / cmax * 92, 0.4), v))

    strong = len(cmp_['up']) + len(cmp_['add'])
    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<div class="report"><div class="top-shape"></div>'
            '<header class="header"><div>'
            '<div class="brand-line">MAKEONE · COVERAGE COMPARISON</div>'
            '<h1>핵심 보장 변화와<br>전 · 후 비교</h1>'
            '<div class="gold-line"><span></span><span></span></div></div>'
            '<div class="customer"><strong>%s 고객님</strong>'
            '<span>제안 기준일 %s</span><span class="step">03 보장 비교</span></div></header>'
            '<main class="content">'
            '<section class="panel"><div class="panel-title">고객님이 꼭 확인할 핵심 보장</div>'
            '<div class="table-wrap"><table><thead><tr><th>보장 항목</th><th>전 (기존)</th>'
            '<th>후 (변경 후)</th><th>증감</th></tr></thead><tbody>%s</tbody></table></div></section>'
            '<section class="panel"><div class="panel-title"><span class="number">1</span>'
            '담보별 전 · 후 보장금액 (만원)</div><div class="chart">'
            '<div class="legend"><span><i class="gray"></i>전 (기존)</span>'
            '<span><i class="navy"></i>후 (변경 후)</span></div>%s</div></section>'
            '<section class="panel"><div class="panel-title"><span class="number">2</span>'
            '월 보험료 전 · 후 (원)</div><div class="premium-wrap"><div class="premium-left">'
            '<div class="premium-row"><b>기존</b><div class="premium-barwrap">'
            '<div class="premium-bar" style="width:%.1f%%"></div></div>'
            '<div class="premium-value">%s원</div></div>'
            '<div class="premium-row"><b>변경 후</b><div class="premium-barwrap">'
            '<div class="premium-bar after" style="width:%.1f%%"></div></div>'
            '<div class="premium-value">%s원</div></div></div>'
            '<div class="increase-box"><strong>%s%s원</strong>'
            '<small>약 %s%% %s</small></div></div></section>'
            '<section class="panel"><div class="panel-title"><span class="number">3</span>'
            '보장 변화 구성 (항목 수)</div><div class="change-wrap"><div class="change-left">%s</div>'
            '<div class="summary"><div class="shield"><span>✓</span></div>'
            '<div class="summary-title">보장 강화 항목</div>'
            '<div class="summary-number">%d개</div>'
            '<div class="summary-text">고객님 보장이<br>더 든든해졌습니다!</div></div></div></section>'
            '</main>'
            '<div class="footer"><div><strong>MAKEONE</strong> 보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></div>'
            '</div></body></html>'
            % (CSS3, client, base_date, trs, bars,
               cmp_['prem_old'] / pmax * 100, format(int(cmp_['prem_old']), ','),
               cmp_['prem_new'] / pmax * 100, format(int(cmp_['prem_new']), ','),
               '+' if up else '-', format(abs(int(sv)), ','),
               abs(cmp_['save_pct']), '인상' if up else '절감',
               chg, strong, client, pg, totpg))

# ═══════════════════ 4 쪽 ═══════════════════


CSS4 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--deepGreen:#0e4b3e;--green:#15865f;--gold:#c6912c;
--goldSoft:#f7f0dc;--red:#bf3434;--line:#dfe5eb;--soft:#f4f7fb;--text:#142c49;--muted:#33404f}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text)}
.page{width:210mm;height:297mm;position:relative;background:#fff;overflow:hidden}
.top-curve{display:none}

.header{position:relative;z-index:2;padding:8.4mm 10.4mm 3.6mm;display:flex;
justify-content:space-between;align-items:flex-start}
.header>div:first-child{width:120mm;flex:none}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:800}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.4mm 0 0;color:var(--navy);font-size:22pt;line-height:1.18;white-space:nowrap}
.client{display:none}
.client strong{display:block;font-size:12pt;margin-bottom:1.6mm}
.client span{display:block;font-size:8pt;margin-top:1mm}
.client .page-step{margin-top:2mm;color:#ffc84d;font-weight:800}
.title-line{margin:0 10mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 23%,var(--navy) 23%)}
.content{padding:2.4mm 10mm 18mm}
.section-header{display:flex;align-items:center;gap:3.2mm;padding-bottom:2.4mm;
border-bottom:.5mm solid var(--navy)}
.section-header h2{margin:0;color:var(--navy);font-size:13pt}
.section-header span{color:#a97821;font-size:7.5pt;font-weight:800}
.coverage-grid{display:flex;gap:3.6mm;margin-top:2.6mm}
.coverage-panel{flex:1}
.coverage-panel h3{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:0 0 1.6mm;font-size:15pt;font-weight:900}
.coverage-panel.brain h3{color:var(--green)}
.coverage-panel.heart h3{color:var(--red)}
.gold-rule{height:.9mm;background:var(--gold);margin-bottom:2mm}
.coverage-table{width:100%;border-collapse:collapse;font-size:7.5pt}
.coverage-table th{padding:1.4mm 1mm;color:#fff;text-align:center;background:var(--deepGreen);
border-right:.2mm solid rgba(255,255,255,.15)}
.coverage-table th:last-child{background:#927224}
.coverage-table td{padding:1mm 1mm;border-bottom:.2mm solid #e1e6eb;text-align:center;
vertical-align:middle}
.coverage-table td:first-child{text-align:left}
.group-row td{background:#edf1f5;color:#183a4f;font-size:8.5pt;font-weight:900;
text-align:left;padding:1.1mm}
.highlight td{background:#fffaf0;border-top:.5mm solid #d3a13b;border-bottom:.5mm solid #d3a13b}
.code-name{display:block;font-weight:800;color:#20384c}
.code-number{display:block;margin-top:.6mm;color:#33404f}
.amount{display:inline-block;min-width:16mm;padding:.6mm 1.2mm;border:.2mm solid #d5dde5;
border-radius:1.2mm;background:#fff;font-size:10.5pt;font-weight:900;color:#18394c}
.dot{width:3.2mm;height:3.2mm;display:inline-block;border-radius:50%;
border:.5mm solid #cfd5db;background:#fff}
.dot.on{border-color:var(--green);background:var(--green)}
.hold{display:inline-block;margin-right:1mm;padding:.4mm 1mm;border-radius:4mm;
background:#d5a239;color:#fff;font-size:6pt;font-weight:900}
.legend{margin:2mm 0 0;font-size:7.5pt;color:#7c8794}
.legend .circle{display:inline-block;width:2.6mm;height:2.6mm;border-radius:50%;
margin-right:.8mm;border:.5mm solid #cfd5db}
.legend .circle.on{background:var(--green);border-color:var(--green)}
.legend strong{color:#b67c19}
.opinion{margin-top:4mm}
.opinion-title{font-size:10.5pt;font-weight:900;color:var(--navy);padding-bottom:1.6mm;
border-bottom:.5mm solid var(--gold)}
.opinion-box{margin-top:2mm;height:24mm;border:.3mm solid #c9d2dc;border-radius:1.8mm}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;
background:linear-gradient(90deg,#082d59,#123f75);color:#fff;padding:0 10mm;display:flex;
align-items:center;justify-content:space-between;font-size:8.5pt}
.footer strong{color:#ffcb57;font-size:9.5pt}
"""

# (구분, 질병명, 코드, 엑셀 담보명, 순환계, 산정특례)  ※순환계·산정특례는 제도라 고정
BRAIN4 = [
    ('출혈성 뇌혈관 (I60~62)', '뇌출혈', 'I60~62', '뇌출혈진단비', 1, 1),
    ('허혈성 뇌혈관 (I63~66)', '뇌졸중 · 뇌경색', 'I63 · 65 · 66', '뇌졸증진단비', 1, 1),
    ('기타 뇌혈관 (I64 · I67~69)', '뇌혈관', 'I64 · 67 · 68 · 69', '뇌혈관진단비', 1, 1),
    ('순환계 확장 · 선천', '뇌동맥류 · 정맥류', 'I71 · 72', None, 1, 0),
    ('순환계 확장 · 선천', '선천 뇌혈관기형', 'Q28.0~28.3', None, 0, 1),
    ('순환계 확장 · 선천', '외상성 뇌출혈', 'S06', '외상성뇌출혈', 0, 1),
]
HEART4 = [
    ('허혈성 심장질환 (I20~25)', '급성심근경색', 'I21~23', '허혈성 진단비', 1, 1, 1),
    ('허혈성 심장질환 (I20~25)', '협심증', 'I20', '@허혈성 진단비', 1, 1, 1),
    ('허혈성 심장질환 (I20~25)', '허혈성', 'I24 · 25', '허혈성 진단비', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심장판막', 'I05 · I34~37', '심장판막', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심근 · 심내막 염증', 'I30~33 · I40', '염증', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '빈맥', 'I47 · 48', None, 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '부정맥', 'I49', '부정맥', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심부전', 'I50', '심부전', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심근병증', 'I42~45', '심근병증', 1, 1, 1),
    ('순환계 확장 (2대+동맥류 · 정맥류 등)', '대동맥류 · 죽상경화', 'I70 · 71', None, 0, 1, 1),
    ('순환계 확장 (2대+동맥류 · 정맥류 등)', '동맥류 · 정맥류 등', '[확인]', None, 0, 1, 1),
    ('순환계 확장 (2대+동맥류 · 정맥류 등)', '선천 심장기형', 'Q20~25', None, 0, 0, 1),
]


def _d4(on):
    return '<span class="dot%s"></span>' % (' on' if on else '')


def _c4(v):
    if v:
        return ('<span class="hold">보장</span>'
                '<span class="amount">%s만</span>' % format(int(v), ','))
    return _d4(0)


def p4(cmp_, client='고객', base_date='', pg=4, totpg=7):
    cov = cmp_['new']['cov']

    def val(key):
        if not key: return 0
        return cov.get(key[1:] if key.startswith('@') else key, 0)

    br = ''
    last = None
    for grp, nm, cd, key, circ, spc in BRAIN4:
        if grp != last:
            br += '<tr class="group-row"><td colspan="4">%s</td></tr>' % grp
            last = grp
        v = val(key)
        br += ('<tr class="%s"><td><span class="code-name">%s</span>'
               '<span class="code-number">%s</span></td><td>%s</td><td>%s</td><td>%s</td></tr>'
               % ('highlight' if v else '', nm, cd, _c4(v), _d4(circ), _d4(spc)))

    ht = ''
    last = None
    for grp, nm, cd, key, sph, circ, spc in HEART4:
        if grp != last:
            ht += '<tr class="group-row"><td colspan="5">%s</td></tr>' % grp
            last = grp
        v = val(key)
        dot_only = bool(key) and key.startswith('@')
        ht += ('<tr class="%s"><td><span class="code-name">%s</span>'
               '<span class="code-number">%s</span></td><td>%s</td><td>%s</td><td>%s</td>'
               '<td>%s</td></tr>'
               % ('highlight' if (v and not dot_only) else '', nm, cd,
                  _d4(1) if dot_only and v else _c4(v),
                  _d4(sph), _d4(circ), _d4(spc)))

    lg = ('<div class="legend"><span class="circle on"></span> 보장 &nbsp;'
          '<span class="circle"></span> 미보장 &nbsp;'
          '<strong>[확인] 산정특례 HOLD · 노란행=보유</strong></div>')

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<article class="page"><div class="top-curve"></div>'
            '<header class="header"><div>'
            '<div class="brand">MAKEONE · 리모델링 리포트</div>'
            '<h1>%s <span style="color:#a4731e">고객님</span> 리모델링 리포트</h1></div>'
            '<div class="client"><strong>%s 고객님</strong>'
            '<span>제안 기준일 %s</span>'
            '<span class="page-step">04 담보별 보장범위</span></div></header>'
            '<div class="title-line"></div>'
            '<main class="content">'
            '<div class="section-header"><h2>담보별 보장범위 — 질병코드 커버</h2>'
            '<span>DISEASE-CODE COVERAGE · 각 축=개별 담보 · 각각 보상 · 산정특례 · 순환계</span></div>'
            '<div class="coverage-grid">'
            '<section class="coverage-panel brain"><h3>뇌 — 질병코드별 커버</h3>'
            '<div class="gold-rule"></div><table class="coverage-table"><thead><tr>'
            '<th>질병 (코드)</th><th>뇌혈관<br>진단비</th><th>순환계</th><th>산정<br>특례</th>'
            '</tr></thead><tbody>%s</tbody></table>%s</section>'
            '<section class="coverage-panel heart"><h3>심장 — 질병코드별 커버</h3>'
            '<div class="gold-rule"></div><table class="coverage-table"><thead><tr>'
            '<th>질병 (코드)</th><th>허혈성<br>진단비</th><th>심장<br>(특정)</th><th>순환계</th>'
            '<th>산정<br>특례</th></tr></thead><tbody>%s</tbody></table>%s</section>'
            '</div>'
            '<section class="opinion"><div class="opinion-title">리포트 의견</div>'
            '<div class="opinion-box"></div></section>'
            '</main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS4, client, client, base_date, br, lg, ht, lg, client, pg, totpg))

# ═══════════════════ 5 쪽 ═══════════════════


CSS5 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--gold:#c78d22;--gold2:#e8bd61;--green:#0f785e;
--green2:#138c68;--red:#d53b31;--blue:#1686c9;--gray:#98a5b4;--line:#dce4ec;--soft:#f8fafc;
--text:#132c49}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text)}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.top-curve{display:none}

.header{position:relative;z-index:2;display:flex;justify-content:space-between;
align-items:flex-start;padding:8.4mm 10.4mm 3.6mm}
.header>div:first-child{width:126mm;flex:none}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:800}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.6mm 0 0;color:var(--navy);font-size:23pt;line-height:1.15;
font-weight:900;white-space:nowrap}
.header h1 .gold{color:var(--gold)}
.page-no{display:none}
.page-no strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;display:block;font-size:26pt;line-height:1}
.page-no span{display:block;margin-top:1.4mm;font-size:8.5pt}
.title-line{margin:0 10mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 24%,var(--navy) 24%)}
.content{padding:2mm 9.6mm 15mm}
.columns{display:flex;gap:5mm}
.column{flex:1;min-width:0}
.big-card{border:.2mm solid #bcd8c9;border-radius:2.4mm;overflow:hidden;background:#fff;margin-bottom:1.8mm}
.big-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;height:9.6mm;display:flex;align-items:center;padding:0 3.2mm;color:#fff;
font-size:11.5pt;font-weight:900}
.big-title.cancer{background:linear-gradient(90deg,#b7861c,#d39b17)}
.big-title.brain{background:linear-gradient(90deg,#0e7258,#138669)}
.big-title::before{content:"";width:3.6mm;height:3.6mm;border-radius:50%;background:#fff;margin-right:1.6mm}
.diagnosis,.brain-table{padding:2mm 1.6mm 2.2mm;background:#f7fcf9}
.check-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--green);font-size:11pt;font-weight:900}
.small-desc{margin-top:.8mm;color:#33404f;font-size:7pt}
.amount-row{display:flex;gap:1.4mm;align-items:center;margin-top:1.4mm}
.amount-row .st{width:16mm;flex:none}
.amount-row .lb{flex:1}
.amount-row .vb{width:24mm;flex:none}
.amount-row .un{width:6mm;flex:none}
.status{height:5.6mm;white-space:nowrap;border-radius:1.2mm;display:flex;align-items:center;justify-content:center;
color:#fff;font-size:8pt;font-weight:900}
.status.join{background:var(--red);padding:0 1mm}
.status.gray{background:var(--gray)}
.label{font-size:7.5pt;font-weight:800;color:#294558}
.label.blue{color:var(--blue)}
.value-box{height:6.2mm;display:flex;justify-content:flex-end;align-items:center;
border:.3mm solid #c9d2dc;border-radius:1.2mm;background:#fff;padding:0 2mm;color:#153d40;
font-weight:900;font-size:11pt}
.value-box.red{color:#d52f28}
.value-box.blue{color:#1583c4}
.unit{font-size:7pt;color:#778593}
.small-card{border:.3mm solid #c9d2dc;border-radius:2mm;padding:2.2mm 2.6mm;margin-bottom:1.4mm;background:#fff}
.small-card h4{margin:0 0 .6mm;color:#123d42;font-size:9pt}
.small-card .desc{color:#33404f;font-size:7pt;line-height:1.35}
.small-card .special{color:#0d7158;font-size:7.5pt;font-weight:900;margin-top:.6mm}
.small-line{display:flex;gap:1.6mm;align-items:center;margin-top:1.1mm}
.small-line .st{width:16mm;flex:none}
.small-line .vb{flex:1}
.small-line .un{width:6mm;flex:none}
.brain-row{display:flex;gap:1.8mm;align-items:center;margin-top:1.2mm}
.brain-row .st{width:13mm;flex:none}
.brain-row .lb{width:20mm;flex:none;font-size:7.5pt;font-weight:800;color:#274958}
.brain-row .lb.blue{color:#1187ca}
.brain-row .vb{flex:1}
.brain-row .un{width:6mm;flex:none}
.brain-row.ind .lb{margin-left:14.8mm}
.customer-center{text-align:center;margin:2.4mm 0 3mm;color:var(--navy)}
.customer-center strong{display:block;font-size:13pt}
.customer-center .gold{color:var(--gold);font-size:9.5pt;font-weight:900;display:block;margin-top:1mm}
.customer-center .counter{margin-top:1.4mm;font-size:12pt;font-weight:900}
.customer-center small{color:#33404f;font-size:7pt}
.special-wrap{display:flex;gap:3.6mm;margin-top:1.4mm;padding:1.6mm;border:.5mm solid #dfbb5d;
border-radius:2.4mm;background:#fffcf5}
.special-card{flex:1;border:.3mm solid #c9d2dc;border-radius:1.8mm;padding:2mm 2.4mm;background:#fff}
.special-card h4{margin:0 0 .8mm;font-size:10pt;color:var(--navy)}
.special-card p{margin:0;color:#33404f;font-size:7pt}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;display:flex;align-items:center;
justify-content:space-between;padding:0 10mm;color:#fff;
background:linear-gradient(90deg,#082d59,#123f75);font-size:8.5pt}
.footer strong{color:#ffcb57;font-size:9.5pt}
"""


def _m5(v):
    return (format(int(v), ',') + '만') if v else ''


def _s5(on):
    return ('<div class="status join">가입</div>' if on
            else '<div class="status gray">상태</div>')


def p5(cmp_, client='고객', base_date='', pg=5, totpg=7):
    cov = cmp_['new']['cov']

    def g(k):
        return cov.get(k, 0)

    isch = g('허혈성 진단비')
    # ★2대 ↔ 순환계 — 같은 값이면 순환계가 답이고 2대는 비운다(지점장 확정)
    # ★동명 담보 2행이 dict에서 합산돼 2,000이 됐다(지점장 지적 2026.08.16).
    #   순환계 주요치료비는 <b>한 행의 값</b>이다 → all에서 대표값(최댓값)을 쓴다.
    _rows2 = [r[3] for r in cmp_.get('all', []) if str(r[1]).strip() == '2대 주요치료비']
    two, circ = (max(_rows2) if _rows2 else g('2대 주요치료비')), 0
    if two:
        circ, two = two, 0

    def row(lb, v, st=None, blue=False, ind=False):
        """★뇌·심 칸이 계속 찌그러졌다(지점장 반복 지적) — flex는 좁은 열에서 눌린다.
           <b>표</b>로 그려 배지·라벨·값 자리를 고정한다."""
        badge = ('<div class="status join" style="height:5.2mm;font-size:7.5pt">가입</div>'
                 if st else '')
        return ('<tr>'
                '<td style="width:11mm;padding:.7mm 0">%s</td>'
                '<td style="width:18mm;padding:.7mm 1mm;font-size:7.5pt;font-weight:800;'
                'color:%s;white-space:nowrap">%s</td>'
                '<td style="padding:.7mm 0"><div class="value-box%s">%s</div></td>'
                '<td style="width:6mm;padding:.7mm 0 .7mm 1.2mm;font-size:6.5pt;'
                'color:#4a5768">만원</td></tr>'
                % (badge, '#1187ca' if blue else '#274958', lb,
                   ' blue' if blue else '', _m5(v)))

    def arow(lb, v, join=False, red=False):
        return ('<div class="amount-row"><div class="st">%s</div>'
                '<div class="lb"><div class="label">%s</div></div>'
                '<div class="vb"><div class="value-box%s">%s</div></div>'
                '<div class="un"><span class="unit">만원</span></div></div>'
                % (_s5(True) if join else '', lb, ' red' if red else '', _m5(v)))

    def card(h4, desc, v, join=False, special='', red=False):
        return ('<div class="small-card"><h4>%s</h4><div class="desc">%s</div>%s'
                '<div class="small-line"><div class="st">%s</div>'
                '<div class="vb"><div class="value-box%s">%s</div></div>'
                '<div class="un"><span class="unit">만원</span></div></div></div>'
                % (h4, desc, ('<div class="special">%s</div>' % special) if special else '',
                   _s5(join), ' red' if red else '', _m5(v)))

    left = (arow('암진단비', g('일반암'), join=True)
            + arow('통합암진단비', g('통합암'))
            + arow('통합전이암진단비', g('통합전이암'))
            + arow('유사암진단비', g('유사암(갑.기.경.제)'))
            + arow('고액암진단비', g('고액암')))

    right = ('<table style="width:100%;border-collapse:collapse">' + row('뇌혈관', g('뇌혈관진단비'), st=True)
             + row('뇌졸중', g('뇌졸증진단비'), ind=True)
             + row('뇌출혈', g('뇌출혈진단비'), blue=True, ind=True)
             + row('심부전', g('심부전'), ind=True)
             + row('부정맥', g('부정맥'), ind=True)
             + row('허혈성', isch, ind=True)
             + row('급성심근경색', isch, ind=True) + '</table>')

    owned = sum(1 for v in (g('일반암'), g('뇌혈관진단비'), isch) if v)

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<article class="page"><div class="top-curve"></div>'
            '<header class="header"><div>'
            '<div class="brand">MAKEONE 리모델링 · 상담 워크시트</div>'
            '<h1>지금 고객의 <span class="gold">3대 주요치료비</span>는?</h1></div>'
            '<div class="page-no"><strong>%d</strong><span>상담 워크시트</span></div></header>'
            '<div class="title-line"></div>'
            '<main class="content"><div class="columns">'
            '<section class="column">'
            '<div class="big-card"><div class="big-title cancer">암 주요치료비</div>'
            '<div class="diagnosis"><div class="check-title">✓ 암 진단비</div>'
            '<div class="small-desc">걸렸을 때 일시금 (기본)</div>%s</div></div>'
            '%s%s%s%s%s%s</section>'
            '<section class="column">'
            '<div class="big-card"><div class="big-title brain">뇌 · 심장 주요치료비</div>'
            '<div class="brain-table"><div class="check-title">✓ 뇌 · 심 진단비</div>'
            '<div class="small-desc">뇌혈관 · 허혈성 일시금</div>%s</div></div>'
            ''
            '%s%s%s%s</section></div>'
            '<section class="special-wrap">'
            '<div class="special-card"><h4>산정특례 (뇌)</h4>'
            '<p>뇌혈관질환 I60~69 전체 · Q28 · S06</p>'
            '<div class="small-line"><div class="st">%s</div>'
            '<div class="vb"><div class="value-box">%s</div></div>'
            '<div class="un"><span class="unit">만원</span></div></div></div>'
            '<div class="special-card"><h4>산정특례 (심장)</h4>'
            '<p>심혈관질환 I20~50 · 판막 전체</p>'
            '<div class="small-line"><div class="st">%s</div>'
            '<div class="vb"><div class="value-box">%s</div></div>'
            '<div class="un"><span class="unit">만원</span></div></div></div>'
            '</section></main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS5, pg, left,
               card('암 주요치료비', '수술 · 방사선 · 약물 정액 (100만~)', g('암주요치료비'),
                    join=bool(g('암주요치료비')), special='암주요치료비Plus', red=True),
               card('암 통합치료비', '검사 · 수술 · 약물 통합 (진단서 전용)', 0),
               card('흥국화재 10억통장', '[갱신형] 플래티넘 건강 리셋플랜 II · 한도 10억원', g('10억 플랜')),
               card('하이클래스 (비급여)', '표적 · 면역 · 중입자 전액본인 커버', g('하이클래스(암)'),
                    special='비급여암주요치료비Plus', red=True),
               card('비급여 암 통합치료비', '비급여 검사 · 수술 · 약물 통합', 0),
               card('암 생활비', '치료 중 소득보상', 0),
               right,
               card('2대 주요치료비', '수술 · 혈전용해 · 중환자실 (100만~)', two, join=bool(two)),
               card('순환계 주요치료비', '부정맥 · 심부전 확대', circ, join=bool(circ),
                    special='신특정순환계질환주요치료비ⅢPlus', red=True),
               card('순환계 통합치료비', '검사 · 수술 · 약물 통합', 0),
               card('순환계 생활비', '치료 중 소득보상', 0),
               _s5(bool(g('산정특례뇌혈관'))), _m5(g('산정특례뇌혈관')),
               _s5(bool(g('산정특례심장'))), _m5(g('산정특례심장')),
               client, pg, totpg))

# ═══════════════════ 6 쪽 ═══════════════════


CSS6 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--gold:#c78d22;--green:#0b6559;--red:#d9342b;
--line:#dbe3eb;--soft:#f7f9fc;--text:#15334b;--muted:#7d8995}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text)}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.top-curve{display:none}

.header{position:relative;z-index:2;display:flex;justify-content:space-between;padding:8.4mm 10.4mm 3.6mm}
.header>div:first-child{width:130mm;flex:none}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:900}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.6mm 0 0;font-size:21pt;color:var(--navy);white-space:nowrap}
.header h1 span{color:var(--gold)}
.page-number{display:none}
.page-number strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:26pt}
.page-number div{font-size:8.5pt}
.line{margin:0 10mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 24%,var(--navy) 24%)}
.content{padding:2mm 9.6mm 15mm}
.columns{display:flex;gap:5mm}
.columns>section{flex:1;min-width:0}
.panel{border:.3mm solid #c9d2dc;border-radius:2.2mm;margin-bottom:1.4mm;overflow:hidden;background:#fff}
.panel-title{padding:1.4mm 3mm;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;padding:1.8mm 3mm;background:linear-gradient(90deg,#0a6459,#087166);color:#fff;
font-size:11pt;font-weight:900}
.panel-title:before{content:"■";margin-right:1.6mm}
.inner{padding:1.4mm 2.2mm}
.inner h3{margin:0 0 .6mm;font-size:10pt;color:#123f42}
.desc{color:var(--muted);font-size:7pt;margin-bottom:1.2mm}
.row{display:flex;align-items:center;gap:1.2mm;margin:.8mm 0}
.row .lb{width:24mm;flex:none}
.row .bx{flex:1}
.row .un{width:6mm;flex:none}
.row.two .lb{width:16mm}
.label{font-size:7.5pt;font-weight:800;color:#294c4a}
.box{height:5.4mm;border:.3mm solid #c9d2dc;border-radius:1.2mm;background:#fff;
display:flex;align-items:center;justify-content:flex-end;padding:0 1.8mm;font-weight:900;
color:#0b2340;font-size:10pt}
.unit{font-size:6.5pt;color:#33404f}
.red{color:var(--red)}
.sub-title{margin:1.2mm 0 .8mm;padding:1mm 1.4mm;background:#f0f4f8;font-size:7.5pt;
font-weight:900;color:#0d554d}
.sub-title:before{content:"■ ";color:#087166}
.customer{display:none}
.customer strong{display:block;font-size:13pt;color:var(--navy)}
.customer .gold{color:var(--gold);font-weight:900;margin-top:.8mm;font-size:9pt}
.customer .life{margin-top:1.6mm;font-size:11pt;font-weight:900;color:var(--navy)}
.customer small{color:#7d8995;font-size:7pt}
.bottom{display:flex;gap:5mm;margin-top:1mm}
.bottom>div{flex:1}
.footer{position:absolute;bottom:0;left:0;right:0;height:13mm;display:flex;align-items:center;
justify-content:space-between;padding:0 10mm;color:#fff;
background:linear-gradient(90deg,#082d59,#123f75);font-size:8.5pt}
.footer strong{color:#ffca55;font-size:9.5pt}
"""


def _m6(v):
    # ★시안 표기는 「17,600만」이다 — 억으로 바꾸지 않는다
    n = int(v or 0)
    return (format(n, ',') + '만') if n else ''


def p6(cmp_, client='고객', base_date='', pg=6, totpg=7):
    cov, old = cmp_['new']['cov'], cmp_['old']['cov']

    def g(k):
        return cov.get(k, 0)

    def r1(lb, key, extra=''):
        return ('<div class="row"><div class="lb"><div class="label">%s</div></div>'
                '<div class="bx"><div class="box">%s%s</div></div>'
                '<div class="un"><span class="unit">만원</span></div></div>'
                % (lb, _m6(g(key)) if key else '', extra))

    def r2(l1, k1, l2, k2):
        return ('<div class="row two"><div class="lb"><div class="label">%s</div></div>'
                '<div class="bx"><div class="box">%s</div></div>'
                '<div class="un"><span class="unit">만원</span></div>'
                '<div class="lb"><div class="label">%s</div></div>'
                '<div class="bx"><div class="box">%s</div></div>'
                '<div class="un"><span class="unit">만원</span></div></div>'
                % (l1, _m6(g(k1)) if k1 else '', l2, _m6(g(k2)) if k2 else ''))

    d = g('상해사망') - old.get('상해사망', 0)
    inj = ('<span class="red">&nbsp;+%s만</span>' % format(int(d), ',')) if d > 0 else ''
    isch = g('허혈성 진단비')

    left = ('<div class="panel"><div class="panel-title">사망 · 진단비 · 수술</div>'
            '<div class="inner"><h3>사망</h3><div class="desc">평생 유지 · 종신형</div>'
            + r1('종신 사망', '일반사망') + r1('질병 사망', '질병사망(80세)')
            + ('<div class="row"><div class="lb"><div class="label">상해 사망</div></div>'
               '<div class="bx"><div class="box">%s%s</div></div>'
               '<div class="un"><span class="unit">만원</span></div></div>'
               % (_m6(old.get('상해사망', 0)), inj))
            + r1('교통상해 사망', '교통상해사망')
            + '</div></div>'
            '<div class="panel"><div class="inner"><h3>진단비</h3>'
            '<div class="desc">암 · 뇌 · 심 진단 일시금</div>'
            '<div class="sub-title">암</div>'
            + r2('고액암', '고액암', '일반암', '일반암')
            + r1('유사암', '유사암(갑.기.경.제)')
            + '<div class="sub-title">뇌</div>'
            + r2('뇌혈관', '뇌혈관진단비', '뇌졸중', '뇌졸증진단비')
            + r1('뇌출혈', '뇌출혈진단비')
            + '<div class="sub-title">심장</div>'
            + ('<div class="row two"><div class="lb"><div class="label">허혈성</div></div>'
               '<div class="bx"><div class="box">%s</div></div>'
               '<div class="un"><span class="unit">만원</span></div>'
               '<div class="lb"><div class="label">협심증</div></div>'
               '<div class="bx"><div class="box">%s</div></div>'
               '<div class="un"><span class="unit">만원</span></div></div>' % (_m6(isch), _m6(isch)))
            + r2('급성심근', None, '심부전', '심부전')
            + r2('염증', '염증', '부정맥', '부정맥')
            + '</div></div>'
            '<div class="panel"><div class="inner"><h3>수술</h3>'
            '<div class="desc">질병 · 상해 구분</div>'
            '<div class="sub-title">질병</div>'
            + r2('질병수술비', '질병수술비', '1~5종', '질병 종수술비(1-5종)')
            + r2('N대수술비', 'n대수술비', '뇌혈관수술', '뇌혈관수술비')
            + r2('허혈성수술', '허혈성수술비', '암수술비', '암수술')
            + '<div class="sub-title">상해</div>'
            + r2('상해수술비', '상해수술비', '1~5종', '상해 종수술비(1-5종)')
            + r2('중대상해수술', '중대한상해수술비', '창상봉합', '창상봉합술')
            + r2('골절수술비', '골절수술비', '5대골절수술', '5대골절수술비')
            + '</div></div>')

    right = ('<div class="panel"><div class="panel-title">후유장해 · 상해 · 일당</div>'
             '<div class="inner"><h3>후유장해</h3>'
             '<div class="desc">상해 · 질병 각 3% ~ 80%</div>'
             + r1('상해 후유 3%', '상해후유3%') + r1('상해 후유 80%', '상해후유80%')
             + r1('질병 후유 3%', '질병후유3%') + r1('질병 후유 80%', '질병후유80%')
             + '</div></div>'
             '<div class="panel"><div class="inner"><h3>상해</h3>'
             '<div class="desc">골절 · 화상 · 깁스 · 응급실</div>'
             + r1('골절', '골절(치아파절제외)') + r1('5대 골절', '5대골절진단비')
             + r1('화상진단비', '화상진단비') + r1('중대화상진단비', '중증화상진단비')
             + r1('반깁스', '반깁스') + r1('깁스', '깁스진단비') + r1('응급실', '응급실(응급)')
             + '</div></div>'
             ''
             '<div class="panel"><div class="inner"><h3>일당</h3>'
             '<div class="desc">입원 · 수술 · 중환자실</div>'
             + r1('질병 입원일당', '질병일당') + r1('상해 입원일당', '상해일당')
             + r1('질병 수술일당', '질병수술일당') + r1('상해 수술일당', '상해수술일당')
             + r1('질병 중환자실', '질병중환자실') + r1('상해 중환자실', '상해중환자실')
             + r1('1인실(종합/상급)', '1인실 종합병원')
             + '</div></div>')

    # ★연금·종신·저축은 계약에서 판별한다(마스터에 담보 행이 없다 · 지점장 확인 2026.08.16)
    import remodel as _rm
    _k = _rm.contract_kinds(cmp_.get('new', {}).get('contracts', []))

    def kline(lb, lst):
        if lst:
            nmk = ' / '.join(dict.fromkeys(c['company'] for c in lst))
            pr = sum(int(c.get('premium') or 0) for c in lst)
            val = ('<span style="color:#0b6559;font-weight:900">가입</span>&nbsp;'
                   + nmk + ('&nbsp;· %s원' % format(pr, ',') if pr else ''))
        else:
            val = '<span style="color:#98a5b4">미가입</span>'
        return ('<div class="row"><div class="lb"><div class="label">%s</div></div>'
                '<div class="bx"><div class="box" style="justify-content:flex-start;'
                'font-size:8pt">%s</div></div>'
                '<div class="un"></div></div>' % (lb, val))

    bottom = ('<div class="bottom">'
              '<div class="panel"><div class="inner"><h3>연금</h3>'
              '<div class="desc">노후 소득 · 계약에서 판별</div>'
              + kline('연금', _k['연금']) + kline('저축', _k['저축']) + '</div></div>'
              '<div class="panel"><div class="inner"><h3>종신 · 사망</h3>'
              '<div class="desc">평생 보장 · 계약에서 판별</div>'
              + kline('종신', _k['종신']) + r1('일반사망', '일반사망') + '</div></div></div>')

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<article class="page"><div class="top-curve"></div>'
            '<header class="header"><div>'
            '<div class="brand">MAKEONE 리모델링 · 평생 지키는 준비</div>'
            '<h1>바뀌지 않는 담보 — <span>은퇴 후에도 평생</span></h1></div>'
            '<div class="page-number"><strong>%d</strong><div>비갱신 추천</div></div></header>'
            '<div class="line"></div>'
            '<main class="content"><div class="columns">'
            '<section>%s</section><section>%s</section></div>%s</main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS6, pg, left, right, bottom, client, pg, totpg))

# ═══════════════════ 7 쪽 ═══════════════════


CSS7 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--gold:#c78916;--gold2:#e7bb59;--green:#0f6957;
--line:#dbe3eb;--soft:#f7f9fc;--cream:#fffaf1;--text:#132c49}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text)}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.top-curve{display:none}

.header{position:relative;z-index:2;display:flex;justify-content:space-between;
align-items:flex-start;padding:8.4mm 10.4mm 3.6mm}
.header>div:first-child{width:142mm;flex:none}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:900}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.6mm 0 0;color:var(--navy);font-size:16.5pt;white-space:nowrap;line-height:1.18;font-weight:900}
.header h1 .gold{color:var(--gold)}
.page-no{display:none}
.page-no strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;display:block;font-size:26pt;line-height:1}
.page-no span{display:block;margin-top:1.6mm;font-size:8.5pt}
.title-line{margin:0 10mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 24%,var(--navy) 24%)}
.content{padding:4.4mm 10mm 16mm}
.greeting{text-align:center;padding:5mm 0 7mm}
.laurel{display:none}
.customer-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:1mm;color:#0b2340;font-size:28pt;font-weight:900}
.greeting-message{margin-top:2.4mm;color:var(--navy);font-size:13pt;line-height:1.55;font-weight:800}
.section-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:2mm;padding-bottom:1.6mm;border-bottom:.9mm solid var(--navy);
color:var(--navy);font-size:13pt;font-weight:900}
.summary-grid{display:flex;gap:4.4mm;margin-top:2.8mm}
.summary-card{flex:1;border:.3mm solid #c9d2dc;border-radius:2.2mm;overflow:hidden;background:#fff}
.summary-head{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;height:9.4mm;display:flex;align-items:center;justify-content:center;color:#fff;
font-size:11pt;font-weight:900}
.summary-head.green{background:linear-gradient(90deg,#0b6757,#0c7764)}
.summary-head.navy{background:linear-gradient(90deg,#0b3264,#123f75)}
.summary-head.gold{background:linear-gradient(90deg,#b57b0f,#d19817)}
.summary-body{padding:9mm 3mm 8mm;text-align:center}
.icon-circle{width:17mm;height:17mm;line-height:17mm;margin:0 auto 2.6mm;border-radius:50%;
font-size:20pt;background:#f0f4f8}
.summary-card:nth-child(1) .icon-circle{color:var(--green);background:#eef8f2}
.summary-card:nth-child(2) .icon-circle{color:var(--navy);background:#f0f4fa}
.summary-card:nth-child(3) .icon-circle{color:var(--gold);background:#fff7e7}
.summary-body h3{margin:0;font-size:12.5pt;color:var(--navy)}
.summary-body p{margin:3.4mm 0 5.4mm;font-size:8.5pt;color:#33404f;line-height:1.7}
.metric{height:12mm;border:.3mm solid #c9d2dc;border-radius:1.8mm;display:flex;align-items:center;
justify-content:space-around;padding:0 3mm;color:var(--navy);font-size:8pt;font-weight:800}
.metric strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:24pt;color:var(--navy)}
.summary-card:nth-child(3) .metric strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold)}
.points{border:.3mm solid #c9d2dc;border-radius:2mm;margin-top:3.4mm;display:flex;padding:5mm 0}
.point{flex:1;text-align:center;padding:0 3.6mm;position:relative}
.point:not(:last-child)::after{content:"";position:absolute;right:0;top:1.6mm;height:24mm;
border-right:.2mm dashed #cbd4de}
.point-icon{width:12mm;height:12mm;line-height:12mm;margin:0 auto 2mm;border-radius:50%;
background:#f2f5f9;font-size:14pt;color:var(--navy)}
.point:nth-child(1) .point-icon{color:var(--green);background:#eef8f2}
.point:nth-child(3) .point-icon{color:var(--gold);background:#fff7e7}
.point b{display:block;font-size:10pt;color:var(--navy)}
.point p{margin:1.4mm 0 0;color:#33404f;font-size:7.5pt;line-height:1.45}
.closing{margin-top:5.4mm;border:.2mm solid #ead6a7;border-radius:2.2mm;
background:#fff;padding:4mm 5mm;text-align:left;height:46mm}
.closing-main{color:#b4760b;font-size:11.5pt;font-weight:900}
.closing-sub{margin-top:1.6mm;color:var(--navy);font-size:10.5pt;font-weight:800}
.hand{margin-top:2.4mm;font-size:13pt;color:#173c73}
.contact{margin-top:3.6mm;padding-top:3mm;border-top:.2mm solid #e6d6ae;display:flex;gap:6mm}
.contact-box{flex:1;display:flex;gap:2.4mm;align-items:center;text-align:left}
.contact-icon{width:11mm;height:11mm;line-height:11mm;text-align:center;flex:none;border-radius:50%;
background:#f0f3f8;font-size:14pt;color:var(--navy)}
.contact-lines{flex:1}
.contact-row{display:flex;gap:2mm;font-size:8pt;color:#50657b;margin-bottom:1.8mm}
.contact-row b{width:26mm;flex:none;line-height:1.25}
.contact-value{flex:1;border-bottom:.2mm solid #cfd7df;min-height:4mm}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;padding:0 10mm;
background:linear-gradient(90deg,#082d59,#123f75);color:#fff;display:flex;align-items:center;
justify-content:space-between;font-size:8.5pt}
.footer strong{color:#ffcb57;font-size:9.5pt}
"""


def p7(cmp_, client='고객', base_date='', pg=7, totpg=7):
    cov = cmp_['new']['cov']
    gap = len(cmp_['up']) + len(cmp_['add'])
    owned = sum(1 for k in ('일반암', '뇌혈관진단비', '허혈성 진단비') if cov.get(k, 0))
    cts = cmp_['new']['contracts']
    # ★★★★★v424 (지점장 확정 2026.08.16): 비갱신 <b>가입율</b> = 전체 보험 기준.
    #   단 <b>저축 · 연금 · 화재보험은 모수에서 제외</b>한다(보장성 보험이 아니다).
    #   실측(박미정): 8건 중 연금·저축 2건 제외 → 모수 6건 · 비갱신 5건 = 83%.
    _base = [c for c in cts
             if not any(x in str(c.get('product') or '') for x in ('저축', '연금', '화재'))]
    nonren = sum(1 for c in _base if '비갱신' in (c.get('renewal') or ''))
    rate = round(nonren / len(_base) * 100) if _base else 0

    cards = [
        ('green', '① 담보 점검 결과', '☑', '보장 공백 최소화',
         '필수 보장 분석을 통해<br>누락 보장 보완 제안',
         '<span>보장 보완 항목</span><strong>%d</strong><span>개</span>' % gap, ''),
        ('navy', '② 3대 주요치료비', '♢', '핵심 치료비 강화',
         '암 · 뇌 · 심장 주요치료비를<br>업그레이드하여 든든하게 준비',
         '<span>3대 치료비 보유</span><strong>%d / 3</strong>' % owned, ''),
        ('gold', '③ 평생 보장 설계', '♡', '평생 유지 · 비갱신 중심',
         '은퇴 후에도 걱정 없이<br>평생 보장 설계 완성',
         '<span>비갱신 가입율</span><strong>%d</strong><span>%%</span>' % rate,
         ' style="color:var(--gold)"'),
    ]
    sg = ''.join('<div class="summary-card"><div class="summary-head %s">%s</div>'
                 '<div class="summary-body">'
                 '<h3%s>%s</h3><p>%s</p><div class="metric">%s</div></div></div>'
                 % (cl, hd, st, h3, p, mt)
                 for cl, hd, ic, h3, p, mt, st in cards)

    pts = [('✓', '보장 안정성', '필수 보장을 빈틈없이<br>구성하여 안정적으로'),
           ('↗', '경제적 효율성', '불필요한 중복은 줄이고<br>필요한 보장은 강화'),
           ('♡', '평생 동반 보장', '비갱신 중심으로<br>평생 보장 유지'),
           ('☷', '체계적 관리', '정기 점검으로<br>지속적인 보장 관리')]
    pt = ''.join('<div class="point"><div class="point-icon">%s</div><b>%s</b><p>%s</p></div>'
                 % t for t in pts)

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>' + CSS7 +
            '</style></head><body><article class="page"><div class="top-curve"></div>'
            '<header class="header"><div>'
            '<div class="brand">MAKEONE LIFE PLAN · 평생 지키는 준비</div>'
            '<h1>MAKEONE LIFE PLAN 분석을 마무리하며<br>'
            '당신의 보장은 <span class="gold">더 단단해집니다</span></h1></div>'
            '<div class="page-no"><strong>' + str(pg) + '</strong><span>클로징</span></div></header>'
            '<div class="title-line"></div><main class="content">'
            '<section class="greeting"><div class="laurel">❧ ★★★ ❧</div>'
            '<div class="customer-title">' + client + ' 고객님</div>'
            '<div class="greeting-message">바뀌는 세상, 바뀌지 않는 보장으로<br>'
            '평생 곁에서 지켜드리겠습니다.</div></section>'
            '<div class="section-title">이번 LIFE PLAN 핵심 요약</div>'
            '<section class="summary-grid">' + sg + '</section>'
            '<div class="section-title" style="margin-top:4.4mm">고객님께 드리는 추천 포인트</div>'
            '<section class="points">' + pt + '</section>'
            '<section class="closing">'
            '<div class="closing-main">지금 이 선택이, 미래의 당신과 가족을 지킵니다.</div>'
            '<div class="closing-sub">항상 곁에서 든든한 동반자가 되겠습니다.</div>'
            '<div class="hand">감사합니다. 언제나 건강하시고 행복하세요!</div>'
            '<div class="contact">'
            '<div class="contact-box"><div class="contact-icon">♙</div><div class="contact-lines">'
            '<div class="contact-row"><b>담당 컨설턴트</b><div class="contact-value"></div></div>'
            '<div class="contact-row"><b>연락처</b><div class="contact-value"></div></div>'
            '</div></div>'
            '<div class="contact-box"><div class="contact-icon">▣</div><div class="contact-lines">'
            '<div class="contact-row"><b>다음 보장 점검 예정</b>'
            '<div class="contact-value"></div></div></div></div>'
            '</div></section></main>'
            '<footer class="footer">'
            '<div><strong>MAKEONE LIFE PLAN</strong>&nbsp;보장분석 자동화</div>'
            '<div>' + client + ' 고객님 · ' + str(pg) + ' / ' + str(totpg) + '</div></footer>'
            '</article></body></html>')



# ═══════════════════ 8 쪽 · 자산 · 연금 ═══════════════════

def p8(cmp_, client='고객', base_date='', pg=7, totpg=8):
    """★자산 · 재무 (지점장 확정 2026.08.16)
       [선저축|선지출] → [지금 시작하세요 전폭] → [현재의 나|미래의 나] → [이미 준비된 것 전폭]
       ㆍ보유 계약은 <b>최대 20건</b>까지 유동적으로 늘어난다(연금·종신이 다건인 고객 대비).
       ㆍ진단서와 <b>같은 구조</b>다 — 한 사람에게 두 문서가 다르게 보이면 안 된다."""
    import remodel as _rm

    def _nonrenew_rate(cs):
        """★비갱신 가입율(지점장 확정 2026.08.16) — <b>전체 보험</b> 기준.
           단 <b>저축 · 연금 · 화재보험은 모수에서 제외</b>한다(보장성 보험이 아니다)."""
        base = [c for c in (cs or [])
                if not any(x in str(c.get('product') or '') for x in ('저축', '연금', '화재'))]
        if not base:
            return 0, 0, 0
        nr = [c for c in base if '비갱신' in str(c.get('renewal') or '')]
        return len(nr), len(base), round(len(nr) * 100 / len(base))

    k = _rm.contract_kinds(cmp_.get('new', {}).get('contracts', []))
    cov = cmp_.get('new', {}).get('cov', {})
    dth = int(cov.get('일반사망', 0))
    prem = int(cmp_.get('prem_new', 0))
    NAVY, GOLD, GREEN, LINE = '#06203f', '#c5a052', '#0e7258', '#c3ccd8'

    def blank(lb):
        return ('<div style="display:flex;align-items:flex-end;gap:2mm;margin:0">'
                '<div style="width:28mm;flex:none;font-size:6.8pt;font-weight:800;color:#33404f">%s</div>'
                '<div style="flex:1;height:2.8mm;border-bottom:.3mm solid %s"></div>'
                '<div style="width:6mm;flex:none;font-size:7.5pt;color:#4a5768">원</div></div>'
                % (lb, LINE))

    def head(t):
        return ('<div style="font-size:9pt;font-weight:900;color:%s;'
                'border-bottom:.4mm solid %s;padding-bottom:.3mm;margin-top:.6mm;font-size:7.2pt">%s</div>'
                % (NAVY, GOLD, t))

    def col(title, sub, bg, body):
        return ('<div style="flex:1;border:.3mm solid %s;border-radius:2.2mm;overflow:hidden">'
                '<div style="background:%s;color:#fff;padding:1.2mm 3mm;font-size:9pt;font-weight:900">'
                '%s<span style="font-size:7.5pt;font-weight:400;margin-left:2mm">%s</span></div>'
                '<div style="padding:1mm 3mm 1.2mm">%s</div></div>'
                % (LINE, bg, title, sub, body))

    now = (head('소득 · 지출')
           + ''.join(blank(x) for x in ['월 평균 소득', '월 평균 생활비', '월 저축 · 투자액',
                                        '월 부채 상환액', '월 잉여자금'])
           + head('부채')
           + ''.join(blank(x) for x in ['주택담보대출', '신용 · 기타대출', '총 부채 합계']))
    _gl = ['노후 준비', '안정적인 자산 관리', '자녀 교육자금',
           '월 현금 흐름 개선', '주택 마련 · 이사', '부채 상환']

    def _chk(t):
        return ('<td style="width:3.2mm;padding:.15mm 0"><div style="width:2.6mm;height:2.6mm;'
                'border:.3mm solid %s"></div></td>'
                '<td style="padding:.15mm 1.6mm .15mm 1.2mm;font-size:6.8pt;font-weight:800;'
                'color:#33404f;white-space:nowrap">%s</td>' % (NAVY, t))

    # ★목표 6개를 <b>2열 3행</b>으로 접어 세로를 절반으로 (지점장 2026.08.16)
    goals = ('<table style="width:100%;border-collapse:collapse">'
             + ''.join('<tr>' + _chk(_gl[i]) + _chk(_gl[i + 3]) + '</tr>' for i in range(3))
             + '</table>')
    fut = (head('보유 자산')
           + ''.join(blank(x) for x in ['예금 · 적금', '주식 · 펀드 · ETF', '부동산(시가)', '총 자산 합계'])
           + head('은퇴 후') + blank('필요 생활비(월)')
           + head('재무 목표') + '<div style="margin-top:.6mm">' + goals + '</div>')

    # ── 보유 계약 표 : 최대 20건까지 유동
    def amt(c):
        lp, pm = int(c.get('lump_sum') or 0), int(c.get('premium') or 0)
        if lp:
            return format(lp, ',') + '원'
        if pm:
            return format(pm, ',') + '원/월'
        return '<span style="color:#98a5b4;font-weight:400">상담 시 확인</span>'

    MAXROW = 10   # ★한 장에 담는 계약 수 상한(지점장 확정 2026.08.16). 넘으면 다음 장으로.

    def rows(lst, kind):
        # ★★★★★v425 (제49조): MAXROW 초과분이 <b>조용히 잘리던 것</b>을 소리나게 바꾼다.
        out = ''
        _cut = max(0, len(lst or []) - MAXROW)
        if _cut:
            print('[v425 잘림] %s %d건 초과 — 표에 명시' % (kind, _cut))
        for c in (lst or []):
            v3 = (('%s만원' % format(dth, ',')) if dth else '—') if kind == 'whole' \
                 else (str(c.get('contract_date') or '') or '—')
            out += ('<tr><td style="padding:.55mm 1.6mm;font-size:7pt;border-bottom:.25mm solid #e7ebef">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:7pt;border-bottom:.25mm solid #e7ebef">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:8pt;font-weight:900;color:%s;'
                    'text-align:right;border-bottom:.25mm solid #e7ebef">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:7pt;text-align:center;'
                    'border-bottom:.25mm solid #e7ebef">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:7pt;text-align:center;'
                    'border-bottom:.25mm solid #e7ebef">%s</td></tr>'
                    % (c['company'], str(c['product'])[:26], NAVY, amt(c), v3,
                       str(c.get('pay_term') or '') or '—'))
        if _cut:
            out += ('<tr><td colspan="5" style="padding:1.6mm;text-align:center;color:#b07d0e;'
                    'font-weight:900;font-size:7.5pt;border-bottom:.25mm solid #e7ebef">'
                    '외 %d건 — 다음 장 참조</td></tr>' % _cut)
        if not out:
            out = ('<tr><td colspan="5" style="padding:2.4mm;text-align:center;color:#98a5b4;'
                   'font-size:8pt;border-bottom:.25mm solid #e7ebef">보유 계약 없음</td></tr>')
        return out

    def tbl_all():
        """★표 3개를 <b>하나로</b> 합친다 — 머리글이 셋이면 그만큼 세로를 먹는다.
           구분 열을 두고 연금 · 종신 · 저축을 한 표에 싣는다(최대 20건)."""
        # ★★★★★v425 (제49조): 상한 초과분이 <b>조용히 잘리던 것</b>을 소리나게 바꾼다.
        #   지점장 확정 = 한 장에 10건(제48조 2항). 넘으면 표에 명시하고 로그를 남긴다.
        body = ''
        n = 0
        _tot = sum(len(x or []) for x in (k['연금'], k['종신'], k['저축']))
        _cut = 0
        if _tot > MAXROW:
            # ★자르지 않는다 — 전부 싣고 넘치면 다음 장으로 간다(제47조).
            print('[v425 다장] 재무 계약 %d건 — 한 장 %d건 기준 %d장 예상'
                  % (_tot, MAXROW, (_tot + MAXROW - 1) // MAXROW))
        for lbl, lst, kind in [('연금', k['연금'], 'pension'),
                               ('종신', k['종신'], 'whole'),
                               ('저축', k['저축'], 'pension')]:
            for c in (lst or []):
                n += 1
                v3 = (('%s만원' % format(dth, ',')) if dth else '—') if kind == 'whole' \
                     else (str(c.get('contract_date') or '') or '—')
                body += ('<tr><td style="padding:.55mm 1.6mm;font-size:7pt;font-weight:900;'
                         'color:%s;border-bottom:.25mm solid #e7ebef">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7pt;border-bottom:.25mm solid #e7ebef">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7pt;border-bottom:.25mm solid #e7ebef">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:8pt;font-weight:900;color:%s;'
                         'text-align:right;border-bottom:.25mm solid #e7ebef">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7pt;text-align:center;'
                         'border-bottom:.25mm solid #e7ebef">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7pt;text-align:center;'
                         'border-bottom:.25mm solid #e7ebef">%s</td></tr>'
                         % (NAVY, lbl, c['company'], str(c['product'])[:24], NAVY, amt(c), v3,
                            str(c.get('pay_term') or '') or '—'))
        if _cut:
            body += ('<tr><td colspan="6" style="padding:1.6mm;text-align:center;color:#b07d0e;'
                     'font-weight:900;font-size:7.5pt;border-bottom:.25mm solid #e7ebef">'
                     '외 %d건 — 다음 장 참조</td></tr>' % _cut)
        if not body:
            body = ('<tr><td colspan="6" style="padding:2.4mm;text-align:center;color:#98a5b4;'
                    'font-size:8pt">보유 계약 없음</td></tr>')
        th = ''.join('<th style="padding:.9mm 1.6mm;background:#eef3f9;font-size:6.8pt;'
                     'text-align:%s;color:%s">%s</th>' % (al, NAVY, t)
                     for t, al in [('구분', 'left'), ('보험사', 'left'), ('상품명', 'left'),
                                   ('가입금액', 'right'), ('가입날짜 · 담보금액', 'center'),
                                   ('납입기간', 'center')])
        return ('<table style="width:100%%;border-collapse:collapse;background:#fff;'
                'border:.3mm solid %s"><tr>%s</tr>%s</table>' % (LINE, th, body))


    have = (tbl_all()
            + '<div style="display:flex;gap:3mm;margin-top:1.6mm">'
              '<div style="flex:1;background:%s;padding:1.4mm;text-align:center;color:#fff">'
              '<div style="font-size:8pt;color:#b9c6d6">사망보장 합계</div>'
              '<div style="font-size:12pt;font-weight:900;margin-top:.6mm">%s</div></div>'
              '<div style="flex:1;background:%s;padding:2mm;text-align:center;color:#fff">'
              '<div style="font-size:8pt;color:#cfe8de">리모델링 후 월 보험료</div>'
              '<div style="font-size:12pt;font-weight:900;margin-top:.6mm">%s원</div></div></div>'
              % (NAVY, ('%s만원' % format(dth, ',')) if dth else '—', GREEN, format(prem, ',')))

    def side(ok, title, sub, flow, items, c, bg):
        ic = '✓' if ok else '✗'
        cells = []
        for n, (a, v) in enumerate(flow):
            if n:
                cells.append('<td style="width:6mm;text-align:center;font-size:13pt;'
                             'font-weight:900;color:%s">→</td>' % c)
            cells.append('<td style="text-align:center;padding:1.2mm 1mm;border:.35mm solid %s;'
                         'border-radius:1.8mm;background:#fff">'
                         '<div style="font-size:7.5pt;font-weight:900;color:#33404f">%s</div>'
                         '<div style="font-size:12pt;font-weight:900;color:%s;margin-top:.6mm">%s</div>'
                         '</td>' % (c, a, c, v))
        return ('<div style="flex:1;border:.5mm solid %s;border-radius:2.4mm;overflow:hidden;'
                'background:%s"><div style="background:%s;color:#fff;padding:2mm 3mm;'
                'text-align:center"><div style="font-size:13pt;font-weight:900">%s %s</div>'
                '<div style="font-size:7.5pt;margin-top:.5mm">%s</div></div>'
                '<div style="padding:1.8mm 2.4mm">'
                '<table style="width:100%%;border-collapse:separate;border-spacing:0"><tr>%s</tr></table>'
                '<div style="margin-top:1.8mm;text-align:center;font-size:8pt;font-weight:800;'
                'color:%s;line-height:1.5">%s</div></div></div>'
                % (c, bg, c, ic, title, sub, ''.join(cells), c, ' · '.join(items)))

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>' + CSS2 +
            # ★v424 다장 : overflow를 풀어 표가 다음 장으로 흐르게 하고,
            #   머리·꼬리는 fixed로 매 장 반복시킨다(계약 20건 이상 대비).
            '.page{height:auto;min-height:297mm;overflow:visible}'
            '.header,.footer{position:fixed;left:0;right:0;background:#fff;z-index:9}'
            '.header{top:0}.footer{bottom:0;background:linear-gradient(90deg,#082d59,#123f75)}'
            '.content{padding-top:30mm;padding-bottom:18mm}' +
            '</style><style>@page{size:A4 portrait;margin:0}.page{height:auto !important;overflow:visible !important}header.header{position:fixed;top:0;left:0;right:0;background:#fff;z-index:9}footer.footer{position:fixed;bottom:0;left:0;right:0;z-index:9}main.content{padding-top:32mm !important;padding-bottom:18mm !important}</style></head><body><article class="page">'
            '<header class="header"><div>'
            '<div class="eyebrow">MAKEONE · ASSET &amp; FINANCE</div>'
            '<h1 style="white-space:nowrap;font-size:19pt">당신은 미래의 자신에게 돈을 보내고 있나요?</h1>'
            '<div class="title-line"></div></div>'
            '<div class="title-line"></div></div></header>'
            '<main class="content">'
            '<div style="display:flex;gap:4mm;margin-top:1mm">'
            + side(True, '선저축 후지출', '저축이 먼저, 소비는 나중에',
                   [('월급', '100%'), ('저축', '70%'), ('생활비', '30%')],
                   ['경제적 자유', '삶의 선택권', '마음의 여유', '행복한 은퇴'], GREEN, '#f7fcf9')
            + side(False, '선지출 후후회', '쓰고 남으면 저축? 남는 게 없다',
                   [('월급', '100%'), ('지출', '100%'), ('남는 것', '0%')],
                   ['불안과 스트레스', '원하는 것 포기', '후회', '불행한 노후'], '#cc5656', '#fdf5f4')
            + '</div>'
            + '</div>'
            '<div style="margin:1.4mm -10mm;background:#06203f;padding:1.6mm;text-align:center">'
            '<div style="font-size:8pt;font-weight:800;color:#b9c6d6">'
            '당신의 선택이 10년 후 당신의 모습을 만듭니다</div>'
            '<div style="font-size:13pt;font-weight:900;color:#fff;margin:.6mm 0;'
            'letter-spacing:-.03em">지금 <span style="color:#e7c274">시작하세요</span></div>'
            '<div style="font-size:7.5pt;color:#d6dee8">'
            '① 소득 파악하기　② 저축 먼저 설정하기(자동이체)　③ 남은 돈으로 생활하기</div></div>'
            '<div style="display:flex;gap:4mm">'
            + col('현재의 나', '오늘 · 함께 적습니다', '#0b3264', now)
            + col('미래의 나', '내일 · 목표를 정합니다', '#b57b0f', fut)
            + '</div>'
            '<section style="margin-top:2.6mm;border:.4mm solid ' + NAVY + ';border-radius:2.2mm;'
            'overflow:hidden;display:flex;flex-direction:column;">'
            '<div style="background:' + NAVY + ';color:#fff;padding:2.4mm 3mm;text-align:center;'
            'font-size:11.5pt;font-weight:900">이미 준비된 것'
            '<span style="font-size:8pt;font-weight:400;margin-left:2mm">보험 계약에서 자동으로</span>'
            '</div><div style="padding:3.4mm 3mm 4mm;background:#f4f7fb;flex:1">'
            + have + '</div></section>'
            '</main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>' + client + ' 고객님 · 리모델링 리포트 · ' + str(pg) + ' / ' + str(totpg) + '</div>'
            '</footer></article></body></html>')


def build(cmp_, client='고객', base_date='', total=8):
    """7쪽 HTML을 순서대로 돌려준다. 페이지 번호·분모는 실제 장수에서 온다(하드코딩 금지)."""
    return [p1(client, base_date, 1, total, cmp_),
            p2(cmp_, client, base_date, 2, total),
            p3(cmp_, client, base_date, 3, total),
            p4(cmp_, client, base_date, 4, total),
            p5(cmp_, client, base_date, 5, total),
            p6(cmp_, client, base_date, 6, total),
            p8(cmp_, client, base_date, 7, total),
            p7(cmp_, client, base_date, 8, total)]
