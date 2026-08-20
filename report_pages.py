# -*- coding: utf-8 -*-
"""★★★★★ report_pages.py — 리모델링 리포트 7쪽 (지점장 시안 정본 2026.08.15)
# 각인: v521-heartfix-20260820

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
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:#1c2430;font-weight:600}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}
.tbar{position:absolute;left:0;top:0;width:210mm;height:14mm;background:#0b2340}
.tgold{position:absolute;left:0;top:14mm;width:210mm;height:2.4mm;background:#7e6528}
.fbar{position:absolute;left:0;bottom:0;width:210mm;height:6mm;background:#0b2340}
.body{position:absolute;left:16mm;right:16mm;top:30mm;font-weight:600}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:13pt;font-weight:800;color:#0b2340;letter-spacing:.18em}
.bln{width:24mm;height:1.6mm;background:#7e6528;margin-top:3mm}
.mark{position:absolute;right:0;top:14mm;font-size:44pt;font-weight:900;color:#f2f4f7;letter-spacing:.06em}
.eyebrow{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:26mm;font-size:9.5pt;font-weight:800;color:#524014;letter-spacing:.28em}
.title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:4mm;font-size:44pt;font-weight:900;color:#0b2340;line-height:1.14;letter-spacing:-.03em}
.rule{width:34mm;height:1.8mm;background:#7e6528;margin-top:7mm}
.sub{margin-top:5mm;font-size:11pt;font-weight:700;color:#0b2340}
.namebox{margin-top:11mm;background:#f4f7fb;border-left:1.4mm solid #0b2340;padding:7mm 8mm;
display:flex;align-items:baseline;gap:6mm}
.namebox .nm{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:40pt;font-weight:900;color:#0b2340}
.namebox .sfx{font-size:13pt;font-weight:800;color:#524014}
.stats{display:flex;gap:5mm;margin-top:9mm}
.stat{flex:1;border-top:1.4mm solid #b08d38;border-left:.38mm solid #626d78;
border-right:.38mm solid #626d78;border-bottom:.38mm solid #626d78;padding:4mm 5mm 5mm}
.stat .k{font-size:8.5pt;font-weight:800;color:#1e2a38}
.stat .v{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3mm;font-size:20pt;font-weight:900;color:#06203f}
.inbox{margin-top:9mm;border:.38mm solid #626d78;padding:5mm 6mm}
.inbox .t{font-size:8.5pt;font-weight:900;color:#524014}
.inbox .l{margin-top:3mm;font-size:10.5pt;font-weight:800;color:#0b2340;line-height:1.8}
.hr{margin-top:12mm;border-top:.2mm solid #686b6e}
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
--softBlue:#f4f8fd;--softGreen:#f1faf5;--line:#5c636b;--text:#132b48;--muted:#455363}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text);font-weight:600}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}   /* ★v487 제98조 원복 */
.top-curve{display:none}

.header{position:relative;z-index:2;padding:8.4mm 10.4mm 4.4mm;display:flex;justify-content:space-between}
.header>div:first-child{width:118mm;flex:none}   /* ★제목이 3줄로 접히던 원인 */
.eyebrow{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:800;letter-spacing:.04em}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.4mm 0 3.6mm;color:#06203f;font-weight:900;font-size:21pt;white-space:nowrap;line-height:1.16;letter-spacing:-.04em}
.title-line{width:100mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 24%,var(--navy) 24%)}
.client{display:none}
.client strong{font-size:12pt;display:block;margin-bottom:1.6mm}
.client span{display:block;font-size:8.4pt;margin-top:1mm;color:#455363}
.client .step{color:#80672e;font-weight:800;margin-top:2mm}
.content{padding:0 10mm 17mm}   /* ★v487 제98조 원복 */
.summary-card{margin-top:1.6mm;border:.5pt solid #626d78;border-radius:3.2mm;padding:5.2mm 6mm;
display:flex}
.summary-col{flex:1;min-height:34mm;padding:2.4mm 5mm}
.summary-col:first-child{border-right:.4pt solid #676a6e}
.summary-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:#06203f;font-size:14pt;font-weight:900}
.summary-sub{margin-top:2.8mm;font-size:9.5pt;color:#1e2a38}
.big-gold{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3.6mm;color:#95690a;font-size:34pt;font-weight:900}
.big-navy{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3.6mm;color:#06203f;font-size:32pt;font-weight:900}
.summary-note{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:3.8mm;font-size:12pt;font-weight:900;color:#06203f}
.summary-gold{margin-top:4.8mm;color:var(--gold);font-size:11.5pt;font-weight:900}
.section-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:6.8mm 4mm 3.4mm;color:#06203f;font-size:15pt;font-weight:900;border-left:1.6mm solid #b08d38;padding-left:3mm}
.compare{display:flex;gap:2.4mm;align-items:center;padding:0 4mm}
.premium-card{flex:1;min-height:30mm;padding:3.2mm 4.4mm;border-radius:2.6mm}
.premium-card.before{background:var(--softBlue);border:.4pt solid #66727f}
.premium-card.after{background:var(--softGreen);border:.4pt solid #607668}
.premium-card b{display:block;font-size:10pt;margin-bottom:2.4mm}
.premium-card.after b{color:var(--green)}
.insurers{font-size:8.4pt;color:#1e2a38}
.premium-number{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:5.2mm;font-size:22pt;font-weight:900;color:var(--navy)}
.after .premium-number{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--green)}
.arrow{width:11mm;flex:none;text-align:center;color:var(--gold);font-size:22pt}
.contract-head{display:flex;justify-content:space-between;align-items:flex-end;margin:5.4mm 4mm 2mm}
.contract-head h3{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:0;font-size:15pt;color:#06203f;font-weight:900;border-left:1.6mm solid #b08d38;padding-left:3mm}
.contract-head span{font-size:8.5pt;color:#1e2a38;font-weight:700}
.report-table{width:calc(100% - 8mm);margin:0 4mm;border-collapse:collapse;font-size:8.5pt}
.report-table th{padding:2.2mm 1.8mm;text-align:left;color:#96671e;border-bottom:.4pt solid #686a6d}
.report-table td{padding:2.2mm 1.8mm;border-bottom:.4pt solid #70757b}
.report-table tbody tr:nth-child(odd) td{background:#eef3f9}
.report-table tbody tr.new td{background:#d7f0e2;font-weight:800}
.report-table tbody tr.del td{background:#7b5f5c;font-weight:800;color:#fff}   /* ★v501 제103조 — 갈색 배경에 초록 글자는 안 보인다 → 배경 진하게 + 글자 흰색 */
.report-table tbody tr.del td:nth-child(3),.report-table tbody tr.del td:nth-child(4){color:#fff}
.status-del{color:#fff;font-weight:900}
.report-table th:nth-child(3),.report-table th:nth-child(4),
.report-table td:nth-child(3),.report-table td:nth-child(4){text-align:right}
.report-table td:nth-child(3){color:#1e2a38}
.report-table td:nth-child(4){font-weight:800;color:var(--navy)}
.report-table th:nth-child(5),.report-table td.st-cell{text-align:center;padding-left:3.2mm}   /* ★v507 제108조 */
.report-table tbody tr.del td:nth-child(5){color:#fff}
.status-new{color:var(--green);font-weight:900}
.total-row td{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:10.5pt;font-weight:900;padding-top:3.4mm;color:var(--navy)!important}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;
background:linear-gradient(90deg,#082d59,#123f75);color:#fff;display:flex;align-items:center;
justify-content:space-between;padding:0 10mm;font-size:8.5pt}
.footer strong{color:#80662b;font-size:9.5pt}
"""


def _m2(v):
    return format(int(v or 0), ',') + '원'


def p2(cmp_, client='고객', base_date='', pg=2, totpg=7):
    sv = cmp_['save_m']
    up = sv < 0
    kl = {(c['company'], c['product']) for c in cmp_['kill']}
    # ★★★★★v521 제118조 (지점장 실측 2026.08.19
    #   «기존보험 11만원에서 2번째 엑셀에서 5만원으로 줄여도 전과 후에 금액이 동일하게 나온다»)
    #   [결함] 유지 계약을 `bf = af = c['premium']`로 <b>한 값에 묶었다</b>. c는 <b>기존 엑셀</b>의
    #     계약이므로 최종 엑셀에서 보험료를 낮춰도 「후」가 <b>기존값 그대로</b> 찍힌다.
    #     게다가 표 합계(prem_new)는 최종 엑셀에서 오므로 <b>계약 합 ≠ 합계</b>가 된다(제2조 등식 위반).
    #   [수정] 「후」는 <b>최종 엑셀의 같은 회사·상품 보험료</b>에서 읽는다. 없으면 그때만 기존값.
    _newprem = {(c['company'], c['product']): c['premium'] for c in cmp_['new']['contracts']}
    rows = [(c, 'keep') for c in cmp_['old']['contracts']] + [(c, 'new') for c in cmp_['prop']]
    trs = ''
    for c, kind in rows:
        if kind == 'new':
            bf, af, tg, cls = 0, c['premium'], '신규', 'status-new'
        elif (c['company'], c['product']) in kl:
            bf, af, tg, cls = c['premium'], 0, '삭제', 'status-del'   # ★v501 제103조
        else:
            bf = c['premium']
            af = _newprem.get((c['company'], c['product']), c['premium'])   # ★v521 제118조
            tg, cls = '유지', ''
        # ★★★★★v507 제108조 — 「후 (변경 후)」와 「상태」가 <b>같은 칸</b>에 있어 붙어 보였다.
        #   비교 엑셀 제89조와 같이 <b>상태를 오른쪽 끝 별도 칸</b>으로 뺀다.
        trs += ('<tr class="%s"><td>%s</td><td>%s</td><td>%s</td>'
                '<td>%s</td><td class="st-cell"><span class="%s">%s</span></td></tr>'
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
            '<th style="width:17%%">전 (기존)</th><th style="width:17%%">후 (변경 후)</th>'
            '<th style="width:9%%">상태</th>'   # ★★★★★v507 제108조 (지점장 2026.08.19 「전/후+상태가 붙어있다」)
            '</tr></thead><tbody>%s</tbody>'
            '<tfoot><tr class="total-row"><td colspan="2">합계</td><td>%s</td><td>%s</td><td></td></tr></tfoot>'
            '</table></main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS2 + _P2FILL, client, base_date,
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


# ★★★★★v502 제103조 (지점장 지시 2026.08.19 「2·3페이지도 여백 없이 하자. 여백 있으면 디테일이 아주 떨어진다」)
#   [실측] 김순자 2쪽 8.8% · 서은옥 2쪽 19.6%.  3쪽은 이미 1.8%(CSS3 세로분산 적용됨).
#   [원인] <b>CSS2는 2쪽·8쪽 공용</b>이다. v490에서 8쪽을 원복하면서 <b>2쪽의 세로분산까지 같이 빠졌다</b>.
#   [수정] <b>2쪽 전용 오버라이드</b>로 세로분산을 되살린다. 8쪽은 원복 상태 그대로 둔다.
#     ★계약 수에 따라 표 길이가 달라지므로 <b>고정 확대는 위험</b>하다(5·6쪽에서 넘친 전례).
#       블록이 여럿인 2쪽은 <b>남는 높이를 블록 사이로 나누는 방식</b>이 맞다(제88조).
_P2FILL = ('.page{display:flex;flex-direction:column}'
           '.content{flex:1;display:flex;flex-direction:column;justify-content:space-between}')


CSS3 = """
@page{size:A4;margin:0}
:root{--navy:#0b2d58;--navy2:#123f79;--gold:#c79532;--gold-light:#8f7847;--green:#159263;
--red:#cc5656;--text:#17273c;--gray:#1e2a38;--line:#76787b;--soft:#f5f8fb}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text);font-weight:600}
.report{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden;display:flex;flex-direction:column}   /* ★v479 제88조 */
.top-shape{display:none}

.header{position:relative;z-index:2;display:flex;justify-content:space-between;
align-items:flex-start;padding:6mm 13mm 2mm}
.brand-line{font-size:9.5pt;color:var(--gold);font-weight:800;letter-spacing:.08em}
h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:1.5mm 0 2mm;color:#06203f;font-size:23pt;font-weight:900;line-height:1.18;letter-spacing:-.04em}
.gold-line{display:flex;gap:1.5mm}
.gold-line span:first-child{width:23mm;height:1.8mm;background:var(--gold);border-radius:5mm}
.gold-line span:last-child{width:12mm;height:1.8mm;background:#92866e;border-radius:5mm}
.customer{display:none}
.customer strong{display:block;font-size:13.5pt;margin-bottom:2mm}
.customer span{display:block;margin-top:1.2mm;font-size:8.5pt}
.customer .step{color:#806a3c;font-weight:800;margin-top:2.5mm}
.content{padding:1mm 13mm 19mm;flex:1;display:flex;flex-direction:column;justify-content:space-between}   /* ★v479 제88조 */
.panel{border:.9pt solid #676d73;background:#fff;border-radius:4mm;margin-bottom:2.8mm;overflow:hidden;box-shadow:0 .6mm 1.4mm rgba(11,45,89,.10)}
.panel-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;display:inline-block;padding:2.6mm 6mm;
border-radius:0 0 3mm 0;color:#fff;background:linear-gradient(90deg,var(--navy),var(--navy2));
font-size:13pt;font-weight:900}
.number{display:inline-block;width:5.4mm;height:5.4mm;line-height:5.4mm;text-align:center;
border-radius:50%;color:var(--navy);background:#fff;font-size:8.5pt;font-weight:900;margin-right:1.5mm}
.table-wrap{padding:.6mm 5mm 1.4mm}
table{width:100%;border-collapse:collapse;font-size:10pt}
th,td{padding:.5mm 3mm;border-bottom:.4pt solid #727375;text-align:right}
th{color:#96671e;background:#fffaf2;font-weight:800;font-size:9pt}
th:first-child,td:first-child{text-align:left}
tbody tr:last-child td{border-bottom:0}
.old{color:#1e2a38}
.new{color:var(--navy);font-weight:800}
.diff{color:var(--green);font-weight:800}.diff.up{color:#0b6559}.diff.down{color:#c0392b}.diff.same{color:#36404b}
.circle-icon{display:inline-block;width:7mm;height:7mm;line-height:7mm;text-align:center;
margin-right:2mm;border-radius:50%;background:#fff6e7;color:var(--gold);font-size:9pt;font-weight:900}
.chart{padding:1.6mm 7mm 2mm}
.legend{text-align:right;margin-bottom:1mm;color:#454c59;font-size:8.5pt}
.legend span{margin-left:4mm}
.legend i{display:inline-block;width:3mm;height:3mm;margin-right:1.2mm}
.legend .gray{background:#f0762a}   /* ★v502 — 범례도 주황 */
.legend .navy{background:var(--navy)}
.bar-row{display:flex;align-items:center;gap:4mm;margin:0.5mm 0}
.bar-label{width:42mm;font-weight:700;font-size:9.5pt;flex:none}
.bar-group{flex:1}
.bar-line{display:flex;align-items:center;gap:2mm;height:3.4mm}
.bar{height:3.4mm;min-width:.5mm}
/* ★★★★★v501 제103조 (지점장 지시 2026.08.19 「가로그래프 색을 극과 극이되도록 해라」)
   구: 전 #6e7780(중간 회색) · 후 남색 — <b>둘 다 어두워 구분이 약했다</b>.
   신: 전 <b>밝은 회색</b>(테두리로 형태 유지) · 후 <b>진한 남색</b> — 명도 차이를 최대로. */
.bar-before{background:#f0762a}   /* ★v502 제103조 — 지점장 「남색에 주황색」 · 전(기존)=주황 */
.bar-after{background:linear-gradient(90deg,#04203f,#0d3a69)}   /* 후(변경 후)=남색 */
.value{font-size:8.5pt;color:#1e2a38;white-space:nowrap}
.value.after{color:var(--navy);font-weight:800}
.premium-wrap{display:flex;gap:6mm;align-items:center;padding:2mm 7mm 2mm}
.premium-left{flex:1}
.premium-row{display:flex;align-items:center;gap:3mm;margin:1.6mm 0}
.premium-row b{width:22mm;font-size:11pt;flex:none}
.premium-barwrap{flex:1}
.premium-bar{height:7mm;background:#f0762a}   /* ★v502 — 기존 막대 주황 */
.premium-bar.after{background:linear-gradient(90deg,#0b2d58,#164679)}
.premium-value{width:26mm;text-align:right;font-weight:800;font-size:10.5pt;flex:none}
.increase-box{width:40mm;flex:none;text-align:center;border:.4pt solid #646b73;border-radius:4mm;
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
.change-bar.gray{background:#85898d}
.change-num{width:12mm;font-size:9pt;font-weight:700;flex:none}
.summary{width:56mm;flex:none;text-align:center;padding:2mm 4mm;border-radius:7mm;
border:.4pt solid #686b6f;background:radial-gradient(circle at 50% 10%,#fff,#f3f7fb)}
.shield{width:7mm;height:8.4mm;margin:0 auto .6mm;position:relative;
background:linear-gradient(180deg,#896e34,#96670d);
clip-path:polygon(50% 0,92% 18%,82% 70%,50% 100%,18% 70%,8% 18%)}
.shield span{position:absolute;left:0;right:0;top:1.4mm;color:#fff;font-size:10pt;font-weight:900}
.summary-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--navy);font-size:11.5pt;font-weight:800}
.summary-number{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:24pt;line-height:1;color:var(--navy);font-weight:900;margin:1mm 0 1.5mm}
.summary-text{font-size:8.5pt;color:#1e2a38}
.footer{position:absolute;left:0;right:0;bottom:0;height:14mm;padding:0 13mm;display:flex;
align-items:center;justify-content:space-between;
background:linear-gradient(90deg,#092b55,#154477);color:#fff;font-size:9pt}
.footer strong{color:#806937}
"""

ICONS3 = ['♥', '◈', '✚', '▦', '✥', '▣', '◆', '●']


def _mw(v):
    v = float(v or 0)
    return ('%s만원' % format(int(v), ',')) if v else '0만원'


def p3(cmp_, client='고객', base_date='', pg=3, totpg=7):
    # ★★★★★v470 제76조 (지점장 사진 2026.08.17 — 3쪽이 <b>텅 비어</b> 올라갔다)
    #   변화가 없는 고객이면 `up + add`가 0개라 표도 차트도 <b>빈칸</b>이 된다.
    #   ★고객은 「변화」만 보러 온 게 아니다. <b>지금 무엇을 갖고 있는지</b>를 봐야 한다.
    #   → 변화 항목이 6개가 안 되면 <b>보유 핵심 담보</b>로 채운다. 빈 쪽을 내보내지 않는다.
    rw = (cmp_['up'] + cmp_['add'])[:6]
    if len(rw) < 6:
        _KEY = ['일반암', '뇌혈관진단비', '허혈성 진단비', '급성심근경색', '뇌졸증진단비',
                '상해사망', '일반사망', '질병사망(80세)', '상해후유3%', '실손 입원']
        _seen = {x[0] for x in rw}
        _cov = (cmp_.get('new') or {}).get('cov', {}) or {}
        _old = (cmp_.get('old') or {}).get('cov', {}) or {}
        for _k in _KEY:
            if len(rw) >= 6:
                break
            if _k in _seen:
                continue
            try: _n = int(float(_cov.get(_k, 0) or 0))
            except Exception: _n = 0
            try: _o = int(float(_old.get(_k, 0) or 0))
            except Exception: _o = 0
            if _n or _o:
                rw.append((_k, _o, _n, _n - _o)); _seen.add(_k)
    vmax = max([max(o, n) for _n, o, n, _d in rw] + [1])
    sv = cmp_['save_m']
    up = sv < 0
    pmax = max(cmp_['prem_old'], cmp_['prem_new'], 1)
    cnt = [('보장 증가', len(cmp_['up']), ''), ('신규 추가', len(cmp_['add']), 'green'),
           ('보장 감소', len(cmp_['down']), 'gray'), ('삭제', len(cmp_['delete']), 'gray')]
    cmax = max([c[1] for c in cnt] + [1])

    trs = ''
    for i, (nm, o, n, d) in enumerate(rw):
        # ★★★★★v472 제77조 (육안검수 실측 2026.08.17) — <b>줄었는데 초록 +로 찍혔다.</b>
        #   실측: 일반암 전 4,000만 → 후 1,000만인데 증감이 「3,000만」 초록.
        #   `abs(d)`로 절댓값만 쓰고 마이너스 부호를 버렸다. <b>보장이 준 것을 늘었다고 보여줬다.</b>
        #   → 부호를 살리고 색도 가른다. 늘면 초록 `+`, 줄면 빨강 `−`.
        _cls = 'diff up' if d > 0 else ('diff down' if d < 0 else 'diff same')
        _sgn = '+' if d > 0 else ('−' if d < 0 else '')
        trs += ('<tr><td><span class="circle-icon">%s</span>%s</td>'
                '<td class="old">%s</td><td class="new">%s</td>'
                '<td class="%s">%s%s</td></tr>'
                % (ICONS3[i % len(ICONS3)], nm, _mw(o), _mw(n),
                   _cls, _sgn, _mw(abs(d))))

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
--goldSoft:#f7f0dc;--red:#bf3434;--line:#737679;--soft:#f4f7fb;--text:#142c49;--muted:#1e2a38}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text);font-weight:600}
.page{width:210mm;height:297mm;position:relative;background:#fff;overflow:hidden}
.top-curve{display:none}

.header{position:relative;z-index:2;padding:8.4mm 10.4mm 3.6mm;display:flex;
justify-content:space-between;align-items:flex-start}
.header>div:first-child{width:120mm;flex:none}
.brand{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold);font-size:9.5pt;font-weight:800}
.header h1{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:2.4mm 0 0;color:var(--navy);font-size:22pt;line-height:1.18;white-space:nowrap}
.client{display:none}
.client strong{display:block;font-size:12pt;margin-bottom:1.6mm}
.client span{display:block;font-size:8.4pt;margin-top:1mm}
.client .page-step{margin-top:2mm;color:#806426;font-weight:800}
.title-line{margin:0 10mm;height:.9mm;background:linear-gradient(90deg,var(--gold) 0 23%,var(--navy) 23%)}
.content{padding:2.4mm 10mm 18mm}
.section-header{display:flex;align-items:center;gap:3.2mm;padding-bottom:2.4mm;
border-bottom:.5mm solid var(--navy)}
.section-header h2{margin:0;color:var(--navy);font-size:13pt}
.section-header span{color:#654813;font-size:8pt;font-weight:800}
.coverage-grid{display:flex;gap:3.6mm;margin-top:2.6mm}
.coverage-panel{flex:1}
.coverage-panel h3{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin:0 0 1.6mm;font-size:15pt;font-weight:900}
.coverage-panel.brain h3{color:var(--green)}
.coverage-panel.heart h3{color:var(--red)}
.gold-rule{height:.9mm;background:var(--gold);margin-bottom:2mm}
.coverage-table{width:100%;border-collapse:collapse;font-size:8pt}
.coverage-table th{padding:2.6mm 1mm;   /* ★v504 제105조 */color:#fff;text-align:center;background:var(--deepGreen);
border-right:.2mm solid rgba(255,255,255,.15)}
.coverage-table th:last-child{background:#927224}
.coverage-table td{padding:2.7mm 1mm;   /* ★v504 제105조 — 의견 칸 자리를 표에 돌린다 */border-bottom:.2mm solid #6c6e71;text-align:center;
vertical-align:middle}
.coverage-table td:first-child{text-align:left}
.group-row td{background:#edf1f5;color:#183a4f;font-size:8.5pt;font-weight:900;
text-align:left;padding:1.1mm}
.highlight td{background:#fffaf0;border-top:.5mm solid #97732a;border-bottom:.5mm solid #97732a}
.code-name{display:block;font-weight:800;color:#20384c}
.code-number{display:block;margin-top:.6mm;color:#1e2a38}
.amount{display:inline-block;min-width:16mm;padding:.6mm 1.2mm;border:.2mm solid #666a6e;
border-radius:1.2mm;background:#fff;font-size:10.5pt;font-weight:900;color:#18394c}
.dot{width:3.2mm;height:3.2mm;display:inline-block;border-radius:50%;
border:.5mm solid #73777a;background:#fff}
.dot.on{border-color:var(--green);background:var(--green)}
.hold{display:inline-block;margin-right:1mm;padding:.4mm 1mm;border-radius:4mm;
background:#997429;color:#fff;font-size:6.8pt;font-weight:900}
.legend{margin:2mm 0 0;font-size:8pt;color:#4a5158}
.legend .circle{display:inline-block;width:2.6mm;height:2.6mm;border-radius:50%;
margin-right:.8mm;border:.5mm solid #73777a}
.legend .circle.on{background:var(--green);border-color:var(--green)}
.legend strong{color:#6d4a0f}
.opinion{margin-top:4mm}
.opinion-title{font-size:10.5pt;font-weight:900;color:var(--navy);padding-bottom:1.6mm;
border-bottom:.5mm solid var(--gold)}
.opinion-box{margin-top:2mm;height:34mm;border:.38mm solid #626d78;border-radius:1.8mm}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;
background:linear-gradient(90deg,#082d59,#123f75);color:#fff;padding:0 10mm;display:flex;
align-items:center;justify-content:space-between;font-size:8.5pt}
.footer strong{color:#80662b;font-size:9.5pt}
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
# ★★★★★v521 제122조 (지점장 실측 2026.08.19 «심장도 엉망이고»)
#   [결함] 허혈성 심장질환 3행이 <b>전부 「허혈성 진단비」 한 칸만</b> 보고 있었다.
#     급성심근경색 1,000 · 협심증 500이 실려 있어도 <b>허혈성 진단비 행이 비어 있으면 셋 다 ○(미보장)</b>.
#     [실측 김순자] 롯데 특정심장Ⅰ 1,000(급성심근) · 특정Ⅱ 500(협심증)인데 표는 <b>전부 미보장</b>.
#   [수정] ㉠급성심근경색 → 「급성심근경색」 행  ㉡협심증 → 「협심증」 행
#         ㉢허혈성(I24·25) → <b>제84조에 따라 협심증 행을 본다</b>(심장의 허혈성은 다 협심증) ·
#           금액은 협심증 줄에 이미 찍히므로 여기는 <b>점만</b>(@).
#         ㉣빈맥(I47·48)은 마스터 43행에 <b>있다</b> — key=None(영구 미보장)은 오답이었다.
#         ㉤심근병증 코드 I42~45 → <b>I42·43</b>. I44·I45는 방실차단·전도장애로 별개다.
HEART4 = [
    ('허혈성 심장질환 (I20~25)', '급성심근경색', 'I21~23', '급성심근경색', 1, 1, 1),
    ('허혈성 심장질환 (I20~25)', '협심증', 'I20', '협심증', 1, 1, 1),
    # ★제123조 — ①단독 허혈성진단비가 있으면 그 금액 ②없으면 묶음(협심증 행) 근거로 점만
    ('허혈성 심장질환 (I20~25)', '허혈성', 'I24 · 25', ('허혈성 진단비', '@협심증'), 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심장판막', 'I05 · I34~37 · I39', '심장판막', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심근 · 심내막 염증', 'I30~33 · I38 · I40 · I41', '염증', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '빈맥', 'I47 · 48', '빈맥', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '부정맥', 'I49', '부정맥', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심부전', 'I50', '심부전', 1, 1, 1),
    ('심장특정 (판막 · 염증 · 부정맥 · 심근)', '심근병증', 'I42 · 43', '심근병증', 1, 1, 1),
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

    # ★★★★★v521 제123조 (지점장 재확정 2026.08.19
    #   «허혈성은 <b>허혈성 단독</b>이다. 심장Ⅰ·Ⅱ 등에는 다 협심증이다»)
    #   ⇒ 「허혈성」 줄은 <b>2층</b>이다.
    #     ①담보명이 허혈성인 <b>단독 담보</b>가 있으면 → 그 값을 <b>금액으로</b> 표시.
    #     ②없고 묶음(특정심장Ⅰ·Ⅱ 등)이 I24·I25를 품고 있으면 → 그 값은 <b>협심증 행</b>에 있으므로
    #       여기는 <b>점만</b>(@) 찍는다. 금액을 두 줄에 쓰면 이중계상으로 보인다(제84조).
    #   구현 = key에 <b>대안 목록</b>을 허용한다. 앞에서부터 값이 있는 것을 쓴다.
    def _pick(key):
        """key(문자열 또는 대안 튜플) → (금액, 점만여부)"""
        for k in ((key,) if isinstance(key, str) or key is None else tuple(key)):
            if not k:
                continue
            _v = cov.get(k[1:] if k.startswith('@') else k, 0)
            if _v:
                return _v, k.startswith('@')
        return 0, False

    def val(key):
        return _pick(key)[0]

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
        v, dot_only = _pick(key)          # ★v521 제123조
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
            '<h1>%s <span style="color:#624512">고객님</span> 리모델링 리포트</h1></div>'
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
            # ★★★★★v504 제105조 (지점장 지시 2026.08.19 「4페이지 짤린다 · 리포트의견+칸 없애도
            #   된다 · 차라리 표 2개를 풀로 채워도 된다」) — <b>의견 제목·박스 삭제</b>.
            #   확보한 자리(제목 6mm + 박스 34mm ≈ 40mm)를 <b>표 2개</b>에 돌린다.
            ''
            '</main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS4, client, client, base_date, br, lg, ht, lg, client, pg, totpg))

# ═══════════════════ 5 쪽 ═══════════════════


CSS5 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--gold:#c78d22;--gold2:#e8bd61;--green:#0f785e;
--green2:#138c68;--red:#d53b31;--blue:#1686c9;--gray:#5b6b7d;--line:#717679;--soft:#f8fafc;
--text:#132c49}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text);font-weight:600}
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
.big-card{border:.2mm solid #697870;border-radius:2.4mm;overflow:hidden;background:#fff;margin-bottom:1.8mm}
.big-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;height:9.6mm;display:flex;align-items:center;padding:0 3.2mm;color:#fff;
font-size:11.5pt;font-weight:900}
.big-title.cancer{background:linear-gradient(90deg,#836014,#976f10)}
.big-title.brain{background:linear-gradient(90deg,#0e7258,#138669)}
.big-title::before{content:"";width:3.6mm;height:3.6mm;border-radius:50%;background:#fff;margin-right:1.6mm}
.diagnosis,.brain-table{padding:2mm 1.6mm 2.2mm;background:#f7fcf9}
.check-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--green);font-size:11pt;font-weight:900}
.small-desc{margin-top:.8mm;color:#1e2a38;font-size:7.6pt}
.amount-row{display:flex;gap:1.4mm;align-items:center;margin-top:1.4mm}
.amount-row .st{width:16mm;flex:none}
.amount-row .lb{flex:1}
.amount-row .vb{width:24mm;flex:none}
.amount-row .un{width:6mm;flex:none}
.status{height:5.6mm;white-space:nowrap;border-radius:1.2mm;display:flex;align-items:center;justify-content:center;
color:#fff;font-size:8.4pt;font-weight:900}
.status.join{background:var(--red);padding:0 1mm}
.status.gray{background:var(--gray)}
.label{font-size:8pt;font-weight:800;color:#294558}
.label.blue{color:var(--blue)}
.value-box{height:6.2mm;display:flex;justify-content:flex-end;align-items:center;
border:.38mm solid #626d78;border-radius:1.2mm;background:#fff;padding:0 2mm;color:#153d40;
font-weight:900;font-size:11pt}
.value-box.red{color:#d52f28}
.value-box.blue{color:#0c4e75}
.unit{font-size:7.6pt;color:#474f58}
.small-card{border:.38mm solid #626d78;border-radius:2mm;padding:2.2mm 2.6mm;margin-bottom:1.4mm;background:#fff}
.small-card h4{margin:0 0 .6mm;color:#123d42;font-size:9pt}
.small-card .desc{color:#1e2a38;font-size:7.6pt;line-height:1.35}
.small-card .special{color:#0d7158;font-size:8pt;font-weight:900;margin-top:.6mm}
.small-line{display:flex;gap:1.6mm;align-items:center;margin-top:1.1mm}
.small-line .st{width:16mm;flex:none}
.small-line .vb{flex:1}
.small-line .un{width:6mm;flex:none}
.brain-row{display:flex;gap:1.8mm;align-items:center;margin-top:1.2mm}
.brain-row .st{width:13mm;flex:none}
.brain-row .lb{width:20mm;flex:none;font-size:8pt;font-weight:800;color:#274958}
.brain-row .lb.blue{color:#0a5179}
.brain-row .vb{flex:1}
.brain-row .un{width:6mm;flex:none}
.brain-row.ind .lb{margin-left:14.8mm}
.customer-center{text-align:center;margin:2.4mm 0 3mm;color:var(--navy)}
.customer-center strong{display:block;font-size:13pt}
.customer-center .gold{color:var(--gold);font-size:9.5pt;font-weight:900;display:block;margin-top:1mm}
.customer-center .counter{margin-top:1.4mm;font-size:12pt;font-weight:900}
.customer-center small{color:#1e2a38;font-size:7.6pt}
.special-wrap{display:flex;gap:3.6mm;margin-top:1.4mm;padding:1.6mm;border:.5mm solid #7c6833;
border-radius:2.4mm;background:#fffcf5}
.special-card{flex:1;border:.38mm solid #626d78;border-radius:1.8mm;padding:2mm 2.4mm;background:#fff}
.special-card h4{margin:0 0 .8mm;font-size:10pt;color:var(--navy)}
.special-card p{margin:0;color:#1e2a38;font-size:7.6pt}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;display:flex;align-items:center;
justify-content:space-between;padding:0 10mm;color:#fff;
background:linear-gradient(90deg,#082d59,#123f75);font-size:8.5pt}
.footer strong{color:#80662b;font-size:9.5pt}
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
        badge = ('<div class="status join" style="height:5.2mm;font-size:8pt">가입</div>'
                 if st else '')
        return ('<tr>'
                '<td style="width:11mm;padding:.7mm 0">%s</td>'
                '<td style="width:18mm;padding:.7mm 1mm;font-size:8pt;font-weight:800;'
                'color:%s;white-space:nowrap">%s</td>'
                '<td style="padding:.7mm 0"><div class="value-box%s">%s</div></td>'
                '<td style="width:6mm;padding:.7mm 0 .7mm 1.2mm;font-size:7.2pt;'
                'color:#1e2a38">만원</td></tr>'
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
             # ★★★★★v521 제120조 (지점장 지적 2026.08.19 «리포트 5페이지에 협심증칸이 없다»)
             #   마스터 41행 <b>협심증 = 심장 블록 첫 행</b>인데 5쪽 뇌·심 진단비 표에만 없었다.
             #   제84조에 따라 <b>심장의 허혈성은 전부 협심증 행</b>으로 떨어진다 —
             #   롯데 특정심장질환Ⅱ · KB/한화/현대 특정Ⅰ 등 <b>묶음 대부분의 착지점</b>이다.
             #   칸이 없으면 값이 있어도 상담지에서 사라진다(제1조 등식1 위반).
             #   ★위치도 <b>마스터 행 순서</b>(협심증→심부전→부정맥)를 따른다.
             + row('협심증', g('협심증'), ind=True)
             + row('심부전', g('심부전'), ind=True)
             + row('부정맥', g('부정맥'), ind=True)
             # ★v472 제77조 — 5쪽도 같은 결함이었다. 급성심근 칸에 <b>허혈성 값</b>을 넣고 있었다.
             + row('허혈성', isch, ind=True)
             + row('급성심근경색', g('급성심근경색'), ind=True) + '</table>')

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
--line:#717579;--soft:#f7f9fc;--text:#15334b;--muted:#5a626b}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text);font-weight:600}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden}   /* ★v487 제98조 */
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
.content{padding:2mm 9.6mm 15mm}   /* ★v487 제98조 */
.columns{display:flex;gap:5mm}
.columns>section{flex:1;min-width:0}
.panel{border:.38mm solid #626d78;border-radius:2.2mm;margin-bottom:.8mm;overflow:hidden;background:#fff}
.panel-title{padding:1.4mm 3mm;font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;padding:1.8mm 3mm;background:linear-gradient(90deg,#0a6459,#087166);color:#fff;
font-size:11pt;font-weight:900}
.panel-title:before{content:"■";margin-right:1.6mm}
.inner{padding:.8mm 1.8mm}
.inner h3{margin:0 0 .4mm;font-size:9.6pt;color:#123f42}
.desc{color:var(--muted);font-size:7.4pt;margin-bottom:.7mm}
.row{display:flex;align-items:center;gap:1.2mm;margin:.35mm 0}
.row .lb{width:24mm;flex:none}
.row .bx{flex:1}
.row .un{width:6mm;flex:none}
.row.two .lb{width:16mm}
.label{font-size:8pt;font-weight:800;color:#294c4a}
/* ★★★★★v465 제72조 2항 (육안검수 실측) — 6쪽 「일반암 14,000만」이 두 줄로 깨졌다.
   진하기를 올리면서 글자가 커져 좁은 칸(2칸 행)을 넘쳤다.
   ★<b>금액 칸은 절대 줄바꿈하지 않는다.</b> 넘치면 글자를 줄여서라도 한 줄로 둔다. */
.box{height:4.5mm;border:.38mm solid #626d78;border-radius:1.2mm;background:#fff;
display:flex;align-items:center;justify-content:flex-end;padding:0 1.4mm;font-weight:900;
color:#0b2340;font-size:10pt;white-space:nowrap;overflow:hidden}
.row.two .box{font-size:8.6pt;padding:0 1mm}
.row.two .un{width:5mm}
.unit{font-size:7.2pt;color:#1e2a38}
.red{color:var(--red)}
.sub-title{margin:.8mm 0 .5mm;padding:.7mm 1.2mm;background:#f0f4f8;font-size:8pt;
font-weight:900;color:#0d554d}
.sub-title:before{content:"■ ";color:#087166}
.customer{display:none}
.customer strong{display:block;font-size:13pt;color:var(--navy)}
.customer .gold{color:var(--gold);font-weight:900;margin-top:.8mm;font-size:9pt}
.customer .life{margin-top:1.6mm;font-size:11pt;font-weight:900;color:var(--navy)}
.customer small{color:#4b5259;font-size:7.6pt}
.bottom{display:flex;gap:5mm;margin-top:1mm}
.bottom>div{flex:1}
.footer{position:absolute;bottom:0;left:0;right:0;height:13mm;display:flex;align-items:center;
justify-content:space-between;padding:0 10mm;color:#fff;
background:linear-gradient(90deg,#082d59,#123f75);font-size:8.5pt}
.footer strong{color:#80652a;font-size:9.5pt}
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
            # ★★★★★v472 제77조 (지점장 지적 2026.08.17 「리포트가 뇌심 입력물이다 안 나온다」)
            #   ★결함 2개 — 실측으로 잡았다.
            #   ① 협심증 칸에 <b>허혈성 값</b>을 넣고 있었다(`_m6(isch)`가 두 번).
            #      사공호 실측: 협심증 1,000만인데 허혈성 0 → <b>빈칸</b>으로 나왔다.
            #   ② 급성심근 칸의 키가 <b>None</b>이라 무슨 값이 있어도 늘 빈칸이었다.
            #      사공호 실측: 급성심근경색 1,000만.
            #   ★제28조(허혈성 단독)는 <b>보유 배지</b> 규칙이다 — 칸에 값을 넣는 것과 다르다.
            #     칸은 <b>그 이름의 담보 값</b>을 넣는다.
            + r2('허혈성', '허혈성 진단비', '협심증', '협심증')
            + r2('급성심근', '급성심근경색', '심부전', '심부전')
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
            + '</div></div>'
            # ★★★★★v466 제74조 — 운전자는 <b>왼쪽 칸</b>에 둔다.
            #   오른쪽(일당 아래)에 넣었더니 6쪽이 2장으로 넘쳐 <b>연금·종신 칸이 통째로 사라졌다</b>(실측).
            #   라벨은 진단서와 같은 이름을 쓴다.
            '<div class="panel"><div class="inner"><h3>운전자</h3>'
            '<div class="desc">벌금 · 합의금 · 변호사 · 부상치료</div>'
            + r2('대인', '대인', '대물', '대물')
            + r2('합의금', '합의금', '6주미만', '6주미만')
            + r2('변호사', '변호사', '자부상', '자부상')
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
             # ★★★★★v466 제74조 (지점장 지적 2026.08.17 「리모델링 리포트 운전자·간병인이 빠졌다.
             #   진단서에서 카피해라」) — 엑셀에는 있는데(간병인 92행·운전 118행) 리포트만 없었다.
             + r1('간병인', '간병인') + r1('간병인지원일당', '간병인지원일당')
             + r1('간호통합병동', '간호통합병동')
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
            val = '<span style="color:#36404b">미가입</span>'
        return ('<div class="row"><div class="lb"><div class="label">%s</div></div>'
                '<div class="bx"><div class="box" style="justify-content:flex-start;'
                'font-size:8.4pt">%s</div></div>'
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
            '<section>%s</section><section class="rcol">%s</section></div>%s</main>'   # ★v506 제107조 — 오른쪽 열 식별
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS6 + _P6BIG, pg, left, right, bottom, client, pg, totpg))

# ═══════════════════ 7 쪽 ═══════════════════


# ★★★★★v501 제103조 (지점장 지시 2026.08.19 「4-6페이지 여백이 많다 세로칸 싸이즈를 늘려라」)
#   CSS6은 <b>6쪽·7쪽 공용</b>이라 직접 키우면 7쪽이 이중으로 커진다(제99조 실패 이력).
#   ⇒ 6쪽 전용 오버라이드를 따로 둔다. 7쪽은 `_P9BIG`, 6쪽은 `_P6BIG`.
#   ★★★★★v501 실측 정정 — <b>5·6쪽은 확대하면 넘친다</b>.
#     김순자(계약 7건·담보 101개) 기준 5쪽 여유 3.3% · 6쪽 여유 0.3%뿐이다.
#     키웠더니 <b>5쪽이 2장 · 6쪽이 3장</b>이 됐다(실측). 서은옥은 여유가 있었지만
#     <b>계약이 많으면 넘친다</b> — 8쪽 원복과 같은 이유(지점장 「계약이 많을 수도 있다」).
#   ⇒ <b>4쪽만 확대</b>하고 5·6쪽은 원복한다. `_P6BIG`은 빈 값으로 둔다(자리는 남겨 둔다).
_P6BIG = ('.rcol .row{margin:1.05mm 0}'          # ★★★★★v506 제107조 (지점장 2026.08.19)
          '.rcol .box{height:6.4mm}'             #   「6페이지 오른쪽 표 3개 세로칸 조금씩만 더 늘려줘」
          '.rcol .inner{padding:1.7mm 1.8mm}'    #   왼쪽 열은 이미 꽉 찼다 — <b>오른쪽만</b> 키운다.
          '.rcol .panel{margin-bottom:1.6mm}')   #   페이지 높이는 긴 쪽(왼쪽)이 정하므로 넘치지 않는다.


CSS7 = """
@page{size:A4 portrait;margin:0}
:root{--navy:#082d59;--navy2:#123f75;--gold:#c78916;--gold2:#e7bb59;--green:#0f6957;
--line:#717579;--soft:#f7f9fc;--cream:#fffaf1;--text:#132c49}
*{box-sizing:border-box}
body{margin:0;background:#fff;font-family:"Noto Sans CJK KR","Noto Sans KR",sans-serif;color:var(--text);font-weight:600}
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden;display:flex;flex-direction:column}   /* ★v479 제88조 세로 분산 */
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
.content{padding:4.4mm 10mm 16mm;flex:1;display:flex;flex-direction:column;justify-content:space-between}   /* ★v479 제88조 */
.greeting{text-align:center;padding:5mm 0 7mm}
.laurel{display:none}
.customer-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:1mm;color:#0b2340;font-size:28pt;font-weight:900}
.greeting-message{margin-top:2.4mm;color:var(--navy);font-size:13pt;line-height:1.55;font-weight:800}
.section-title{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;margin-top:2mm;padding-bottom:1.6mm;border-bottom:.9mm solid var(--navy);
color:var(--navy);font-size:13pt;font-weight:900}
.summary-grid{display:flex;gap:4.4mm;margin-top:2.8mm}
.summary-card{flex:1;border:.38mm solid #626d78;border-radius:2.2mm;overflow:hidden;background:#fff}
.summary-head{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;height:9.4mm;display:flex;align-items:center;justify-content:center;color:#fff;
font-size:11pt;font-weight:900}
.summary-head.green{background:linear-gradient(90deg,#0b6757,#0c7764)}
.summary-head.navy{background:linear-gradient(90deg,#0b3264,#123f75)}
.summary-head.gold{background:linear-gradient(90deg,#b57b0f,#966d10)}
.summary-body{padding:9mm 3mm 8mm;text-align:center;font-weight:600}
.icon-circle{width:17mm;height:17mm;line-height:17mm;margin:0 auto 2.6mm;border-radius:50%;
font-size:20pt;background:#f0f4f8}
.summary-card:nth-child(1) .icon-circle{color:var(--green);background:#eef8f2}
.summary-card:nth-child(2) .icon-circle{color:var(--navy);background:#f0f4fa}
.summary-card:nth-child(3) .icon-circle{color:var(--gold);background:#fff7e7}
.summary-body h3{margin:0;font-size:12.5pt;color:var(--navy)}
.summary-body p{margin:3.4mm 0 5.4mm;font-size:8.5pt;color:#1e2a38;line-height:1.7}
.metric{height:12mm;border:.38mm solid #626d78;border-radius:1.8mm;display:flex;align-items:center;
justify-content:space-around;padding:0 3mm;color:var(--navy);font-size:8.4pt;font-weight:800}
.metric strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;font-size:24pt;color:var(--navy)}
.summary-card:nth-child(3) .metric strong{font-family:"Noto Sans CJK KR Black","Noto Sans CJK KR",sans-serif;color:var(--gold)}
.points{border:.38mm solid #626d78;border-radius:2mm;margin-top:3.4mm;display:flex;padding:5mm 0}
.point{flex:1;text-align:center;padding:0 3.6mm;position:relative}
.point:not(:last-child)::after{content:"";position:absolute;right:0;top:1.6mm;height:24mm;
border-right:.2mm dashed #71767c}
.point-icon{width:12mm;height:12mm;line-height:12mm;margin:0 auto 2mm;border-radius:50%;
background:#f2f5f9;font-size:14pt;color:var(--navy)}
.point:nth-child(1) .point-icon{color:var(--green);background:#eef8f2}
.point:nth-child(3) .point-icon{color:var(--gold);background:#fff7e7}
.point b{display:block;font-size:10pt;color:var(--navy)}
.point p{margin:1.4mm 0 0;color:#1e2a38;font-size:8pt;line-height:1.45}
.closing{margin-top:5.4mm;border:.2mm solid #83775d;border-radius:2.2mm;
background:#fff;padding:4mm 5mm;text-align:left;height:46mm}
.closing-main{color:#6c4606;font-size:11.5pt;font-weight:900}
.closing-sub{margin-top:1.6mm;color:var(--navy);font-size:10.5pt;font-weight:800}
.hand{margin-top:2.4mm;font-size:13pt;color:#173c73}
.contact{margin-top:3.6mm;padding-top:3mm;border-top:.2mm solid #807761;display:flex;gap:6mm}
.contact-box{flex:1;display:flex;gap:2.4mm;align-items:center;text-align:left}
.contact-icon{width:11mm;height:11mm;line-height:11mm;text-align:center;flex:none;border-radius:50%;
background:#f0f3f8;font-size:14pt;color:var(--navy)}
.contact-lines{flex:1}
.contact-row{display:flex;gap:2mm;font-size:8.4pt;color:#303c49;margin-bottom:1.8mm}
.contact-row b{width:26mm;flex:none;line-height:1.25}
.contact-value{flex:1;border-bottom:.2mm solid #73787c;min-height:4mm}
.footer{position:absolute;left:0;right:0;bottom:0;height:13mm;padding:0 10mm;
background:linear-gradient(90deg,#082d59,#123f75);color:#fff;display:flex;align-items:center;
justify-content:space-between;font-size:8.5pt}
.footer strong{color:#80662b;font-size:9.5pt}
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

# ★★★★★v488 제99조 (지점장 지적 2026.08.19 「7페이지 심각하지?」) — <b>7쪽 전용 확대</b>.
#   [실측] 김순자 7쪽 하단 <b>28.3%</b>가 빈다. 정상범(보유계약 없음)은 <b>절반</b>이 빈다.
#   [지시] 「억지로 벌린 공백은 삭제하고 <b>입력칸의 세로폭을 더 늘려라</b>」
#   ★1차 실패 — CSS6을 직접 키웠더니 <b>6쪽이 2장으로 넘쳤다</b>(CSS6은 6쪽·7쪽 공용).
#     ⇒ <b>7쪽 전용 오버라이드(EXTRA9)</b>에만 넣는다. 공용 CSS는 손대지 않는다.
_P9BIG = ('.content{padding:2mm 9.6mm 15mm}'
          '.panel{margin-bottom:2.6mm}'
          '.inner{padding:2.4mm 1.8mm}'
          '.row{margin:1.7mm 0}'
          '.box{height:10.2mm}'
          '.ctb th{padding:2.9mm 1.6mm}'
          '.ctb td{padding:2.8mm 1.6mm}')

EXTRA9 = '.ctb{width:100%;border-collapse:collapse;margin:1.4mm 0 2.4mm;font-size:8pt}.ctb th{background:#0b2340;color:#fff;font-weight:900;padding:1.4mm 1.6mm;border:.3mm solid #0b2340}.ctb td{border:.3mm solid #626d78;padding:1.3mm 1.6mm;color:#1e2a38}.ctb td.g{background:#eef3f9;font-weight:900;text-align:center}.ctb td.hi{color:#b3261e;font-weight:900}.ctb td.ok{color:#0b6559;font-weight:900}.ctb .tn{font-weight:400}.sect{margin:1.6mm 0 1.2mm;padding:1.6mm 3mm;background:linear-gradient(90deg,#0a6459,#087166);color:#fff;font-size:10.5pt;font-weight:900;border-radius:0 3mm 3mm 0;display:inline-block}' + _P9BIG


def p9(cmp_, client='고객', base_date='', pg=7, totpg=9):
    """★★★★★v467 제74조 2항 (지점장 지적 2026.08.17 「PPT도 운전자·간병인 페이지 없는데」)

    ★내 잘못: 6쪽에 <b>작은 칸 6개</b>를 끼워넣고 「넣었다」고 보고했다.
      지점장이 말한 것은 <b>진단서 10쪽을 페이지째 카피하라</b>는 것이었다.
      진단서 10쪽 = 운전자보험 담보 + 자동차보험 대비표 + 간병비 담보 + 간병인 비교표.
    ★리포트와 진단서는 <b>같은 구조</b>여야 한다 — 한 사람에게 두 문서가 다르면 안 된다.
    """
    old = (cmp_.get('old') or {}).get('cov', {}) or {}
    new = (cmp_.get('new') or {}).get('cov', {}) or {}

    def line(lb, key):
        """보유(검정) + 제안 증가분(레드)을 한 칸에. 원천은 비교 결과다."""
        try: o = int(float(old.get(key, 0) or 0))
        except Exception: o = 0
        try: n = int(float(new.get(key, 0) or 0))
        except Exception: n = 0
        val = (format(o, ',') + '만') if o else ''
        if n > o:
            add = '<span class="red">+%s만</span>' % format(n - o, ',')
            val = (val + '&nbsp;' + add) if val else add
        return ('<div class="row"><div class="lb"><div class="label">%s</div></div>'
                '<div class="bx"><div class="box">%s</div></div>'
                '<div class="un"><span class="unit">만원</span></div></div>' % (lb, val))

    drive_l = ('<div class="panel"><div class="inner"><h3>벌금 · 합의금</h3>'
               '<div class="desc">교통사고 형사 · 행정</div>'
               + line('대인 벌금', '대인') + line('대물 벌금', '대물')
               + line('합의금', '합의금') + line('6주미만 합의금', '6주미만')
               + '</div></div>')
    drive_r = ('<div class="panel"><div class="inner"><h3>변호사 · 위로금</h3>'
               '<div class="desc">기타 지원</div>'
               + line('변호사비', '변호사') + line('자동차부상위로금', '자부상')
               + '</div></div>')

    care_l = ('<div class="panel"><div class="inner"><h3>간병인 지원</h3>'
              '<div class="desc">보험사 파견</div>'
              + line('간병인지원일당', '간병인지원일당') + line('간호통합병동', '간호통합병동')
              + '</div></div>')
    care_r = ('<div class="panel"><div class="inner"><h3>간병인 사용</h3>'
              '<div class="desc">직접 고용</div>'
              # ★★★★★v510 제109조 (지점장 지시 2026.08.19 「7페이지 질병 간병인일당 →
              #   간호통합병동으로 변경」) — 구 키 `질병간병인일당`은 <b>마스터에 행이 없어</b>
              #   값이 항상 빈칸이었다(실측 마스터 간병 행 = 간병인 · 간병인지원일당 · 간호통합병동 3개).
              + line('간병인일당', '간병인') + line('간호통합병동', '간호통합병동')
              + '</div></div>')

    cmp_tbl = ('<table class="ctb"><tr><th style="width:18%">구분</th>'
               '<th>자동차보험 <span class="tn">(의무)</span></th>'
               '<th>운전자보험 <span class="tn">(선택)</span></th></tr>'
               '<tr><td class="g">보장 대상</td><td>타인의 피해</td><td class="hi">운전자 본인</td></tr>'
               '<tr><td class="g">책임 종류</td><td>민사 배상</td><td class="hi">형사 · 행정</td></tr>'
               '<tr><td class="g">주요 보장</td><td>대인 · 대물</td>'
               '<td class="hi">벌금 · 형사합의금 · 변호사선임비</td></tr></table>')

    care_tbl = ('<table class="ctb"><tr><th style="width:20%">구분</th>'
                '<th>간병인지원일당</th><th>간병인사용일당</th></tr>'
                '<tr><td class="g">방식</td><td>보험사가 간병인 <b>직접 배정</b></td>'
                '<td>내가 <b>직접 고용</b> 후 정액 지급</td></tr>'
                '<tr><td class="g">간병인 선택</td><td>불가 (배정 · 교체 가능)</td>'
                '<td class="ok"><b>간병인 or 지인</b> 가능</td></tr>'
                '<tr><td class="g">갱신</td><td>5 · 10 · 15 · 20년 갱신</td>'
                '<td class="ok">비갱신 선택 가능</td></tr>'
                '<tr><td class="g">납입면제</td><td>갱신형이라 <b>갱신 끝나면 다시 납부</b></td>'
                '<td class="ok"><b>납입면제 가능</b></td></tr>'
                '<tr><td class="g">인건비 상승</td><td>간병인 배정 (교체 가능)</td>'
                '<td>체증형 (5년 10%↑)</td></tr></table>')

    return ('<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>%s%s</style></head><body>'
            '<article class="page"><div class="top-curve"></div>'
            '<header class="header"><div>'
            '<div class="brand">MAKEONE 리모델링 · 일상 리스크</div>'
            '<h1>운전자 · 간병 — <span>일상 리스크 대비</span></h1></div>'
            '<div class="page-number"><strong>%d</strong><div>운전자·간병</div></div></header>'
            '<div class="line"></div>'
            '<main class="content">'
            '<div class="sect">■ 운전자보험 담보</div>'
            '<div class="columns"><section>%s</section><section>%s</section></div>'
            '%s'
            '<div class="sect">■ 간병비 담보</div>'
            '<div class="columns"><section>%s</section><section>%s</section></div>'
            '%s'
            '</main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>%s 고객님 · 리모델링 리포트 · %d / %d</div></footer>'
            '</article></body></html>'
            % (CSS6, EXTRA9, pg, drive_l, drive_r, cmp_tbl, care_l, care_r, care_tbl,
               client, pg, totpg))


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
    NAVY, GOLD, GREEN, LINE = '#06203f', '#7e6528', '#0e7258', '#626d78'

    def blank(lb):
        return ('<div style="display:flex;align-items:flex-end;gap:2mm;margin:0">'
                '<div style="width:28mm;flex:none;font-size:7.4pt;font-weight:800;color:#1e2a38">%s</div>'
                '<div style="flex:1;height:4.6mm;border-bottom:.38mm solid %s"></div>'   # ★v503 제104조
                '<div style="width:6mm;flex:none;font-size:8pt;color:#1e2a38">원</div></div>'
                % (lb, LINE))

    def head(t):
        return ('<div style="font-size:9pt;font-weight:900;color:%s;'
                'border-bottom:.4mm solid %s;padding-bottom:.3mm;margin-top:.6mm;font-size:7.8pt">%s</div>'
                % (NAVY, GOLD, t))

    def col(title, sub, bg, body):
        return ('<div style="flex:1;border:.38mm solid %s;border-radius:2.2mm;overflow:hidden">'
                '<div style="background:%s;color:#fff;padding:1.2mm 3mm;font-size:9pt;font-weight:900">'
                '%s<span style="font-size:8pt;font-weight:400;margin-left:2mm">%s</span></div>'
                '<div style="padding:1.8mm 3mm 2.0mm">%s</div></div>'   # ★v503 제104조
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
                'border:.38mm solid %s"></div></td>'
                '<td style="padding:.15mm 1.6mm .15mm 1.2mm;font-size:7.4pt;font-weight:800;'
                'color:#1e2a38;white-space:nowrap">%s</td>' % (NAVY, t))

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
        return '<span style="color:#36404b;font-weight:400">상담 시 확인</span>'

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
            out += ('<tr><td style="padding:.55mm 1.6mm;font-size:7.6pt;border-bottom:.32mm solid #70757b">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:7.6pt;border-bottom:.32mm solid #70757b">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:8.4pt;font-weight:900;color:%s;'
                    'text-align:right;border-bottom:.32mm solid #70757b">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:7.6pt;text-align:center;'
                    'border-bottom:.32mm solid #70757b">%s</td>'
                    '<td style="padding:.55mm 1.6mm;font-size:7.6pt;text-align:center;'
                    'border-bottom:.32mm solid #70757b">%s</td></tr>'
                    % (c['company'], str(c['product'])[:26], NAVY, amt(c), v3,
                       str(c.get('pay_term') or '') or '—'))
        if _cut:
            out += ('<tr><td colspan="5" style="padding:1.6mm;text-align:center;color:#c88d20;'
                    'font-weight:900;font-size:8pt;border-bottom:.32mm solid #70757b">'
                    '외 %d건 — 다음 장 참조</td></tr>' % _cut)
        if not out:
            out = ('<tr><td colspan="5" style="padding:2.4mm;text-align:center;color:#36404b;'
                   'font-size:8.4pt;border-bottom:.32mm solid #70757b">보유 계약 없음</td></tr>')
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
                body += ('<tr><td style="padding:.55mm 1.6mm;font-size:7.6pt;font-weight:900;'
                         'color:%s;border-bottom:.32mm solid #70757b">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7.6pt;border-bottom:.32mm solid #70757b">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7.6pt;border-bottom:.32mm solid #70757b">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:8.4pt;font-weight:900;color:%s;'
                         'text-align:right;border-bottom:.32mm solid #70757b">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7.6pt;text-align:center;'
                         'border-bottom:.32mm solid #70757b">%s</td>'
                         '<td style="padding:.55mm 1.6mm;font-size:7.6pt;text-align:center;'
                         'border-bottom:.32mm solid #70757b">%s</td></tr>'
                         % (NAVY, lbl, c['company'], str(c['product'])[:24], NAVY, amt(c), v3,
                            str(c.get('pay_term') or '') or '—'))
        if _cut:
            body += ('<tr><td colspan="6" style="padding:1.6mm;text-align:center;color:#c88d20;'
                     'font-weight:900;font-size:8pt;border-bottom:.32mm solid #70757b">'
                     '외 %d건 — 다음 장 참조</td></tr>' % _cut)
        if not body:
            body = ('<tr><td colspan="6" style="padding:2.4mm;text-align:center;color:#36404b;'
                    'font-size:8.4pt">보유 계약 없음</td></tr>')
        th = ''.join('<th style="padding:.9mm 1.6mm;background:#eef3f9;font-size:7.4pt;'
                     'text-align:%s;color:%s">%s</th>' % (al, NAVY, t)
                     for t, al in [('구분', 'left'), ('보험사', 'left'), ('상품명', 'left'),
                                   ('가입금액', 'right'), ('가입날짜 · 담보금액', 'center'),
                                   ('납입기간', 'center')])
        return ('<table style="width:100%%;border-collapse:collapse;background:#fff;'
                'border:.38mm solid %s"><tr>%s</tr>%s</table>' % (LINE, th, body))


    have = (tbl_all()
            + '<div style="display:flex;gap:3mm;margin-top:1.6mm">'
              '<div style="flex:1;background:%s;padding:1.4mm;text-align:center;color:#fff">'
              '<div style="font-size:8.4pt;color:#6c7178">사망보장 합계</div>'
              '<div style="font-size:12pt;font-weight:900;margin-top:.6mm">%s</div></div>'
              '<div style="flex:1;background:%s;padding:2mm;text-align:center;color:#fff">'
              '<div style="font-size:8.4pt;color:#67746f">리모델링 후 월 보험료</div>'
              '<div style="font-size:12pt;font-weight:900;margin-top:.6mm">%s원</div></div></div>'
              % (NAVY, ('%s만원' % format(dth, ',')) if dth else '—', GREEN, format(prem, ',')))

    def side(ok, title, sub, flow, items, c, bg):
        ic = '✓' if ok else '✗'
        cells = []
        for n, (a, v) in enumerate(flow):
            if n:
                cells.append('<td style="width:6mm;text-align:center;font-size:13pt;'
                             'font-weight:900;color:%s">→</td>' % c)
            cells.append('<td style="text-align:center;padding:1.2mm 1mm;border:.42mm solid %s;'
                         'border-radius:1.8mm;background:#fff">'
                         '<div style="font-size:8pt;font-weight:900;color:#1e2a38">%s</div>'
                         '<div style="font-size:12pt;font-weight:900;color:%s;margin-top:.6mm">%s</div>'
                         '</td>' % (c, a, c, v))
        return ('<div style="flex:1;border:.5mm solid %s;border-radius:2.4mm;overflow:hidden;'
                'background:%s"><div style="background:%s;color:#fff;padding:2mm 3mm;'
                'text-align:center"><div style="font-size:13pt;font-weight:900">%s %s</div>'
                '<div style="font-size:8pt;margin-top:.5mm">%s</div></div>'
                '<div style="padding:1.8mm 2.4mm">'
                '<table style="width:100%%;border-collapse:separate;border-spacing:0"><tr>%s</tr></table>'
                '<div style="margin-top:1.8mm;text-align:center;font-size:8.4pt;font-weight:800;'
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
            '<div style="font-size:8.4pt;font-weight:800;color:#6c7178">'
            '당신의 선택이 10년 후 당신의 모습을 만듭니다</div>'
            '<div style="font-size:13pt;font-weight:900;color:#fff;margin:.6mm 0;'
            'letter-spacing:-.03em">지금 <span style="color:#73613a">시작하세요</span></div>'
            '<div style="font-size:8pt;color:#6b6f74">'
            '① 소득 파악하기　② 저축 먼저 설정하기(자동이체)　③ 남은 돈으로 생활하기</div></div>'
            '<div style="display:flex;gap:4mm">'
            + col('현재의 나', '오늘 · 함께 적습니다', '#0b3264', now)
            + col('미래의 나', '내일 · 목표를 정합니다', '#b57b0f', fut)
            + '</div>'
            '<section style="margin-top:2.6mm;border:.4mm solid ' + NAVY + ';border-radius:2.2mm;'
            'overflow:hidden;display:flex;flex-direction:column;">'
            '<div style="background:' + NAVY + ';color:#fff;padding:2.4mm 3mm;text-align:center;'
            'font-size:11.5pt;font-weight:900">이미 준비된 것'
            '<span style="font-size:8.4pt;font-weight:400;margin-left:2mm">보험 계약에서 자동으로</span>'
            '</div><div style="padding:3.4mm 3mm 4mm;background:#f4f7fb;flex:1">'
            + have + '</div></section>'
            '</main>'
            '<footer class="footer"><div><strong>MAKEONE</strong>&nbsp;보장분석 자동화</div>'
            '<div>' + client + ' 고객님 · 리모델링 리포트 · ' + str(pg) + ' / ' + str(totpg) + '</div>'
            '</footer></article></body></html>')


def build(cmp_, client='고객', base_date='', total=9):
    """7쪽 HTML을 순서대로 돌려준다. 페이지 번호·분모는 실제 장수에서 온다(하드코딩 금지)."""
    return [p1(client, base_date, 1, total, cmp_),
            p2(cmp_, client, base_date, 2, total),
            p3(cmp_, client, base_date, 3, total),
            p4(cmp_, client, base_date, 4, total),
            p5(cmp_, client, base_date, 5, total),
            p6(cmp_, client, base_date, 6, total),
            # ★v467 제74조 2항 — 운전자·간병은 <b>전용 페이지</b>다(진단서 10쪽과 같은 구조).
            p9(cmp_, client, base_date, 7, total),
            p8(cmp_, client, base_date, 8, total),
            p7(cmp_, client, base_date, 9, total)]
