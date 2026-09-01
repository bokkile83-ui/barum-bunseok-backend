# -*- coding: utf-8 -*-
"""GA채널 회사별 비교표 (설계사 참고용) — 2표/페이지. 보장설명서 부록."""
import html as _h

# ★각인 v622-dollar-20260901
def _tbl(head, rows, hl=1):
    th="".join(f"<th>{_h.escape(str(c))}</th>" for c in head)
    body=""
    for r in rows:
        c=f'<td class="gh">{_h.escape(str(r[0]))}</td>'
        for i,v in enumerate(r[1:],1):
            cls=' class="ghl"' if hl and i==hl else ''
            c+=f'<td{cls}>{_h.escape(str(v))}</td>'
        body+=f"<tr>{c}</tr>"
    return f'<table class="gtb"><tr>{th}</tr>{body}</table>'

def _blk(num, title, sub, src, head, rows, bullets, hl=1, note=""):
    b="<br>".join(f"• {x}" for x in bullets)
    nt=f'<div class="gnote">{note}</div>' if note else ''
    return (f'<div class="gblk"><div class="gsec"><span class="gn">{num}</span>{_h.escape(title)}'
            f'<span class="gsub">{_h.escape(sub)} · {src}</span></div>'
            f'{_tbl(head,rows,hl)}{nt}'
            f'<div class="gtalk">{b}</div></div>')

# ── 11개 표 정의 ──
B01=_blk("01","유병자보험 개요","손보 간편심사 인수","p130",
 "구분|삼성|현대|DB|KB|메리츠|한화|흥국|롯데|NH|하나".split("|"),
 [["3.0.5","O","O","O","O","O","O","O","O","-","-"],
  ["3.1.5","O","O*","O","O","O","O","O","O","-","O"],
  ["3.5.5","O","O","O","O","O","O*","O","O","O","O"],
  ["3.10.5↑","O","O","10","-","O","O","O*","O","10","O"],
  ["5년고지","5대","6대","6대","6대","6대","6대","6대","6대","6대","6대"],
  ["예외질환","194","500","333","163","3천","2900","3.5천","4806","558","405"],
  ["최대인수","무제한","5","5","5","5","4","5","5","5","5"]],
 ["유병력 고객 핵심=인수유형 — 삼성화재 5년고지 5대·최대인수 무제한으로 문턱 최저",
  "현대해상 O*=조건부 → 청약 전 인수기준 확인 / 고당지 고객은 롯데 3.6.10(7월) 검토"])

B02=_blk("02","비급여 통합치료비 (암)","급여 끝, 비급여 시작","p68",
 "구분|삼성생명|동양생명|삼성화재|현대해상|메리츠".split("|"),
 [["통합치료비","암통치1억","비급여1억","비급여1억","하이클래스8천","비급여1억"],
  ["고액항암약물","3천","3천","3천","3천","3천"],
  ["다빈치로봇","1천/2백","1천/2백","1천/5백","1천/2백","1천/2백"],
  ["비급여암치료","1천/2백","1천/2백","1천/2백","1천/2백","1천/2백"],
  ["40세 남/여","29,301/18,963","37,047/21,531","25,247/17,784","27,730/17,030","24,500/17,970"],
  ["50세 남/여","37,154/20,323","47,201/22,972","32,274/19,054","36,610/16,870","31,470/17,260"]],
 ["국민건보 암 부담률 75%(−5.2%p)→비급여 본인부담 25% 공백. 고액항암·다빈치·비급여암 정액 보장으로 실손 공백 메움",
  "담보값 5사 대동소이 → 보험료·가입한도·병원규모(종합이상)로 비교"])

B03=_blk("03","암주요치료비 (항암호르몬)","한도↓ 치료↑","p67",
 ["구분","삼성생명","푸본현대 마이픽"],
 [["암주요치료비","일반1천/2백·상급1천/2백","일반1천/2백·상급1천/2백"],
  ["치료횟수","연1회","연1회 (일반2천 가능)"],
  ["40세 남/여","39,168/37,422","39,776/38,746"],
  ["50세 남/여","51,648/40,328","51,516/41,732"]],
 ["전립선암 66%·유방암 70~80% 항암호르몬 치료 → 수술·방사선 후 수년 지속, 기존 한도 축소 추세로 별도 확보 필요",
  "담보값 동일(1천/2백) → 푸본 일반2천 확대가 차별점, 보험료로 결정"])

B04=_blk("04","항암방사선·약물 + 중입자","신치료 정액 보장","p71",
 "구분|삼성생명|한화생명|미래에셋|흥국생명|라이나|ABL|삼성화재".split("|"),
 [["항암약물","5천/1천","5천/1천","3천/6백","5천/5백","5천","3천/6백","5천/2천"],
  ["항암방사선","5천/1천","5천/1천","5천/1천","5천/5백","5천","5천/1천","5천/2천"],
  ["중입자방사선","5천","5천","5천","5천","5천","5천","5천"],
  ["40세 남/여","50,665/49,085","47,715/48,830","35,478/38,519","36,103/36,063","47,280/47,420","46,520/41,450","73,092/63,516"]],
 ["항암약물·방사선 5천/1천 + 중입자 5천이 시장 표준 → 담보값 큰 차이 없음",
  "중입자=고가 신치료, 회당 부담 큼 → 정액 5천이 핵심. 차별화는 보험료·기본계약(삼성화재 1억형은 P↑)"],
 note="? 삼성화재 50세 보험료 원본 일부 잘림. 단위 만원/원, 40세 20년납 100세만기.")

B05=_blk("05","순환계 통합치료","진단→재활 원스톱","p88",
 "구분|삼성생명|한화생명|미래에셋|동양생명|삼성화재|메리츠".split("|"),
 [["가입한도","연1억","연1억","연5천","연1억","연1억","연1억"],
  ["MRI/PET/CT","5/5/10","20/20/20","5/5/10","20/20/20","10/10/10","5/5/10"],
  ["수술","2.5천","2천","1천","2천","2천","2천"],
  ["혈전용해","3천","2.5천","1천","2.5천","2천","2.5천"],
  ["혈전제거","3천","5백","-","1천","2천","-"],
  ["에크모","2천","2천","2천","2천","2천","2천"],
  ["재활(입/외)","2(연15)","5(연30)","2(연15)","2(연15)","50(연1)","2(연15)"]],
 ["진단→검사→치료(혈전용해·제거·에크모)→재활 원스톱. 혈전제거·용해 한도 회사 편차 큼→뇌심 고위험 고객 우선 확인",
  "삼성생명 혈전용해·제거 각 3천으로 두터움 / 미래에셋 한도↓ 대신 보험료 경쟁력"])

B06=_blk("06","치매 표적약물 허가치료비","레켐비 등 신약","p104",
 "구분|삼성생명|한화생명|교보생명|KB라이프|동양생명|DB생명|농협생명|라이나".split("|"),
 [["총보장","2,200","2,080","2,500","3,200","2,080","2,700","1천","3,500"],
  ["회차별(7회)","5백/1천","1천","1천","3,200","1천","1천","1천","1천"],
  ["급여CDR검사","10","10","-","9","-","-","-","-"],
  ["MRI/PET","20","10","20","50","20","30","20","-"],
  ["최저보험료","2만","1만","3만","1만","3만","1만","2만","2만"]],
 ["표적치매약물=신약 치료비, 회차별 한도(1/7/13/19회) 지급. 차별점=동시가입·급여CDR/MRI·요양담보 유무",
  "총보장 라이나3,500·KB3,200·교보2,500 상위 / KB·미래·신한 7월 개정·출시 → 재비교 필요"])

B07=_blk("07","간병 (사용일당/요양병원)","주요사 비교","p110",
 "구분|KDB|KB라이프|흥국생명|NH농협|현대해상|DB손보|한화손보|흥국화재".split("|"),
 [["사용일당 180↓","20/5/5","15/7/7","15/6/7","20/5/7","20/5/7","15/5/7","20/6/7","20/5/7"],
  ["181↑","-","10/-/3","20/-/7","20/-/7","20/-/7","10/-/7","20/6/7","15/-/7"],
  ["40세 남/여","32,949/40,180","29,447/36,883","32,444/41,854","58,340/73,330","36,060/47,070","41,230/52,100","43,917/48,103","56,397/64,766"]],
 ["기준선=간병인 사용일당 15만 or 요양병원. 181일↑(장기간병) 유무가 핵심 — KDB는 장기구간 없어 초기형",
  "NH농협 요양특화라 두텁고 보험료↑ / 고령·독거·부부 고객에 장기간병 리스크로 접근"],
 note="사용일당=(초기/중기/장기)만원. 사무직 20년납 100세만기.")

B08=_blk("08","간병인 지원일당 (20년갱신)","KB손보 vs 메리츠","p111",
 ["구분","KB손보 (탑클래스3N5)","메리츠화재 (통합간편)"],
 [["기본계약","상해사망1백","상해사망·상해80% 1백"],
  ["상해입원일당(1-180)","1만원","5천원"],
  ["질병입원일당(1-180)","1만원","5천원"],
  ["요양성특정(181)","1만원","5천원 (+간호간병통합 7/7)"],
  ["40세 남/여","9,761/15,208","10,020/18,980"],
  ["50세 남/여","19,680/27,482","17,980/31,580"],
  ["60세 남/여","37,867/51,562","33,990/55,440"]],
 ["둘 다 20년 갱신형 → 갱신 시 인상 가능, 고객에게 갱신구조 안내 필수",
  "40·50대 KB 남자 유리, 60세 항목별 역전 → 나이 기준 비교. 물가상승 대비 '20년 고정' 소구"],
 note="간병 콜센터 KB 1522-8213 / 메리츠 1688-0090.")

B09=_blk("09","생보 수술비","질병수술+종수술+순환계2천","p120",
 "구분|삼성생명|한화생명|신한|미래에셋|동양생명|KB라이프|DB생명|ABL|라이나".split("|"),
 [["질병수술 전체","20","60","180","100","170","160","240","290","180"],
  ["일반","20","20","40","20","20","20","40","20","30"],
  ["종수술1~2종","20/30","20/30","20/40","20/30","30/40","20/40","15/30","20/40","20/30"],
  ["4~5종","3백/1천","3백/5백","5백/1천","7천/1천","8백/1천","1천/2천","3백/5백","5백/1천","1백/3백"],
  ["순환계","2천","순환1억","2천/1천","순환5천","순환1억","2천","뇌심2천","2천","2천"],
  ["남/여(원)","63,316/48,536","94,562/73,631","96,126/82,381","한도조정","84,971/71,863","95,426/73,390","76,066/57,396","91,506/87,422","98,887/77,798"]],
 ["공식=질병수술비(특정포함)+종수술 플랜+순환계2천 3단. 종수술1~5종=난이도, 5종에 순환계 결합 많음",
  "질병수술 전체 20~290 편차 큼(ABL290·DB240). 일반/경증제외/특정제외/상급종합 세부조건 함께 봐야"])

B10=_blk("10","손보 수술비","질병수술+종수술+순환계2천","p122",
 "구분|삼성화재|현대해상|DB손보|KB손보|메리츠|한화손보|흥국화재|롯데손보|하나손보".split("|"),
 [["질병수술 전체","200","180","200","190","200","300","150","200","200"],
  ["일반","30","30","30","30","20","20","20","30","30"],
  ["특정제외","70","50","70","60","20","150","20","70","70"],
  ["상급종합","100","100","100","100","100","100","100","100","100"],
  ["종수술1~2종","30/50","20/50","20/30","20/50","20/30","30/30","30/50","30/50","30/40"],
  ["4~5종","3백/5백","5백/1천","1천/1천","5백/5백","1백/4백","5백/1천","5백/1천","4백/5백","5백/1천"],
  ["순환계","2천/5백","2천/1천","2천/1천","2천","순환1억","2천/1천","2천/1천","2천","2천"],
  ["남(원)","96,764","92,100","112,250","86,775","91,940","142,021","148,123","77,303","139,918"]],
 ["생보와 동일 공식이나 손보는 순환계 (질병/상해) 분리표기 많음 → 어느 쪽인지 확인",
  "질병수술 전체 한화300 최고·흥국150 최저(단 보험료도 높음). 롯데 200에 7만원대=가성비"])

B11=_blk("11","손보 주요 수술별 보장금액","상급종합·감액후","p123",
 "수술(코드)|삼성화재|현대|DB|KB|메리츠|한화|흥국|롯데|하나".split("|"),
 [["백내장(1종)","150","150","150","150","140","180","160","160","160"],
  ["자궁근종(1종)","230","200","220","210","220","330","180","230","230"],
  ["충수염(2종)","250","230","230","240","230","330","200","250","250"],
  ["무릎관절(2종)","250","230","230","240","230","180","200","250","250"],
  ["척추수술(3종)","500","380","300","390","300","600","350","400","600"],
  ["갑상선암(3종)","500","380","300","390","300","600","350","400","600"],
  ["뇌개두술(5종)","2,700","3,180","3,200","2,790","2,600","3,300","3,150","2,700","3,200"],
  ["심장카테터(3종)","2,500","2,380","2,300","2,390","2,300","2,600","2,350","2,400","2,600"]],
 ["다빈도 실전수술(백내장·충수염·무릎)로 체감 비교. 뇌개두·심장카테터는 순환계 결합 2,600~3,300",
  "상급종합·감액후 기준. 초기(감액)·병원등급·코드 따라 변동 → 실제 청약 설계값 확인"],
 note="? 셀=원본 화면 잘림 미확정.")

PAGES=[[B01,B02,B03,B04],[B05,B06,B07,B08],[B09,B10,B11]]

GA_CSS="""
.gap-cover{background:#fff;color:#1B2A4A;height:297mm;padding:55mm 20mm;position:relative;border-top:6mm solid #C9A24B;}
.gap-cover .gt{font-size:30pt;font-weight:800;line-height:1.25;color:#1B2A4A;}
.gap-cover .gt b{color:#9c7c2e;}
.gap-cover .gd{font-size:11pt;color:#4A5A72;margin-top:8mm;line-height:1.7;}
.gap-cover .gwarn{position:absolute;bottom:24mm;left:20mm;right:20mm;font-size:8.5pt;color:#9c7c2e;border-top:1.5pt solid #C9A24B;padding-top:4mm;line-height:1.5;}
.gpg{padding:6mm 8mm 10mm;}
.gpg3 .gtb{font-size:6.8pt;margin-bottom:1.8mm;}
.gpg3 .gtb th{padding:1.2mm 0.6mm;font-size:6.4pt;}
.gpg3 .gtb td{padding:1.2mm 0.6mm;}
.gpg3 .gsec{font-size:9.6pt;margin-bottom:1.6mm;}
.gblk{margin-bottom:2.6mm;}
.gsec{font-size:9pt;font-weight:800;color:#1B2A4A;border-left:3pt solid #C9A24B;padding-left:2.2mm;margin-bottom:1mm;}
.gsec .gn{background:#1B2A4A;color:#fff;font-size:8pt;padding:0.6mm 1.6mm;border-radius:2mm;margin-right:2mm;}
.gsec .gsub{font-size:7.5pt;color:#6B7A90;font-weight:600;margin-left:2mm;}
.gtb{width:100%;border-collapse:collapse;font-size:5.9pt;margin-bottom:1.6mm;}
.gtb th{background:#1B2A4A;color:#fff;padding:0.8mm 0.5mm;text-align:center;font-weight:700;font-size:6.4pt;border:0.3pt solid #1B2A4A;}
.gtb td{border:0.3pt solid #D8DEE8;padding:0.7mm 0.5mm;text-align:center;color:#2B3A52;}
.gtb td.gh{background:#EEF1F6;font-weight:800;color:#1B2A4A;text-align:left;padding-left:1.4mm;}
.gtb td.ghl{color:#9c7c2e;font-weight:800;}
.gtb tr:nth-child(even) td{background:#FAFBFD;}
.gnote{font-size:5.4pt;color:#6B7A90;line-height:1.3;margin-bottom:0.7mm;}
.gtalk{background:#EAF2FB;border-left:2.4pt solid #2E5A88;border-radius:1.2mm;padding:1.2mm 2mm;font-size:6pt;line-height:1.35;color:#1B2A4A;}
"""



# ═══ 암 보장률 인포메이션 3쪽 (지점장 지시 2026.08.22 «표지 뒤») ═══
# ★_GA_BODY(회사별 비교표 본문 3장)는 계속 False. 이 3쪽은 별개다.
_CS = """
<style>
.ci{padding:0 11mm;}
.ci .lead{font-size:8.6pt;line-height:1.75;color:#333;margin:3.2mm 0 4mm;}
.ci h1{font-size:21pt;font-weight:900;color:#1B2A4A;line-height:1.25;margin-top:2mm;}
.ci h1 .r{color:#C0392B;}
.ci .kpi{width:100%;border-collapse:separate;border-spacing:2.4mm 0;margin-bottom:5mm;}
.ci .kpi td{width:25%;border:0.8pt solid #DDE3EC;border-radius:2mm;padding:3mm 2mm;text-align:center;background:#FBFCFE;}
.ci .kpi .kl{font-size:7pt;color:#6B7A90;font-weight:700;}
.ci .kpi .kv{font-size:17pt;font-weight:900;color:#1B2A4A;margin:1.2mm 0 0.8mm;}
.ci .kpi .kd{font-size:7pt;font-weight:700;}
.ci .dn{color:#C0392B;} .ci .up{color:#1F7A4D;} .ci .fl{color:#6B7A90;}
.ci .st{font-size:10pt;font-weight:800;color:#1B2A4A;margin:0 0 2.4mm;padding-left:3mm;border-left:3.2pt solid #C9A24B;}
.ci .st span{font-size:8pt;color:#9c7c2e;font-weight:700;}
.ci table.d{width:100%;border-collapse:collapse;font-size:7.6pt;margin-bottom:2mm;}
.ci table.d th{background:#1B2A4A;color:#fff;padding:1.8mm 1mm;font-weight:700;border:0.5pt solid #1B2A4A;}
.ci table.d td{padding:1.7mm 1mm;border:0.5pt solid #DDE3EC;text-align:center;color:#333;}
.ci table.d td.l{text-align:left;padding-left:2.6mm;font-weight:700;color:#1B2A4A;}
.ci table.d tr.hi td{background:#FDF2F2;font-weight:800;color:#C0392B;}
.ci table.d tr.hi td.l{color:#C0392B;}
.ci .src{font-size:6.6pt;color:#8892A4;margin:1.4mm 0 4.6mm;}
.ci .box{border:0.9pt solid #C9A24B;background:#FDFBF4;border-radius:2mm;padding:3.4mm 4mm;font-size:8.2pt;line-height:1.8;color:#333;}
.ci .box b.t{color:#9c7c2e;}
.ci .bar{width:100%;border-collapse:collapse;font-size:8pt;margin-bottom:1.5mm;}
.ci .bar td{padding:1.1mm 0;vertical-align:middle;}
.ci .bar .yr{width:13mm;font-weight:800;color:#1B2A4A;font-size:8.4pt;}
.ci .bar .tr{background:#EEF2F7;height:6.2mm;border-radius:1mm;position:relative;}
.ci .bar .fi{height:6.2mm;border-radius:1mm;background:#1B3A63;}
.ci .bar .fr{background:#C0392B;}
.ci .bar .vv{width:16mm;text-align:right;font-weight:900;font-size:9pt;color:#1B2A4A;}
.ci .bar .vr{color:#C0392B;}
.ci .ln{width:100%;border-collapse:collapse;font-size:7.4pt;text-align:center;margin-bottom:1.5mm;}
.ci .ln td{border:0.5pt solid #DDE3EC;padding:1.5mm 0.4mm;color:#333;}
.ci .ln td.h{background:#F4F6FA;font-weight:800;color:#1B2A4A;}
.ci .ln td.e{background:#FDF2F2;font-weight:900;color:#C0392B;}
</style>
"""

def _hd(sub):
    return ('<div class="top"><div class="eb">MAKEONE · 보장분석 인포메이션</div>'
            '<div class="nm">암 <b>보장률 리포트</b></div>'
            f'<div class="pgn"><b>@@PN@@</b>{sub}</div><div class="bar"></div></div>')

def _ft(r):
    return (f'<div class="ft"><b>MAKEONE</b> 보장분석 자동화'
            f'<span class="r">출처: 국민건강보험공단 보도자료(2025.12.30) · {r}</span></div>')


def _p1():
    kp = [("암질환 보장률", "75.0%", '<span class="dn">▼ 1.3%p</span>'),
          ("암 비급여 본인부담", "16.0%", '<span class="dn">▲ 1.2%p</span>'),
          ("4대 중증질환", "81.0%", '<span class="dn">▼ 0.8%p</span>'),
          ("전체 보장률", "64.9%", '<span class="fl">전년 동일</span>')]
    kh = "".join(f'<td><div class="kl">{a}</div><div class="kv">{b}</div>'
                 f'<div class="kd">{c}</div></td>' for a, b, c in kp)
    rows = [("전체", "81.8", "8.3", "9.9", "81.0 (-0.8)", "8.3", "10.7 (+0.8)", 0),
            ("암질환", "76.3", "8.9", "14.8", "75.0 (-1.3)", "9.0", "16.0 (+1.2)", 1),
            ("뇌혈관질환", "88.2", "6.7", "5.1", "87.9 (-0.3)", "6.4", "5.7 (+0.6)", 0),
            ("심장질환", "90.0", "5.1", "4.9", "90.3 (+0.3)", "4.9", "4.8 (-0.1)", 0),
            ("희귀·중증난치", "89.0", "8.4", "2.6", "89.3 (+0.3)", "8.1", "2.6", 0)]
    tb = ""
    for r in rows:
        cls = ' class="hi"' if r[7] else ''
        tb += (f'<tr{cls}><td class="l">{r[0]}</td>'
               + "".join(f"<td>{v}</td>" for v in r[1:7]) + "</tr>")
    yrs = ["2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
    val = ["63.4", "62.6", "62.7", "63.8", "64.2", "65.3", "64.5", "65.7", "64.9", "64.9"]
    ln = ('<table class="ln"><tr>' + "".join(f'<td class="h">{y}</td>' for y in yrs) + "</tr><tr>"
          + "".join(f'<td{" class=e" if i >= 8 else ""}>{v}</td>' for i, v in enumerate(val)) + "</tr></table>")
    return ('<div class="pg">' + _hd("암 보장률") + '<div class="body ci">'
            + _CS +
            '<h1>암, 보장률은 내려가고 <span class="r">비급여는 올라갔다</span></h1>'
            '<div class="lead">2024년 건강보험 보장률은 <b>64.9%</b>로 전년과 같지만, <b>암질환만 1.3%p 하락</b>했다. '
            '법정 본인부담은 그대로인데 <b>비급여 본인부담이 14.8% → 16.0%</b>로 뛴 것이 원인이다.</div>'
            f'<table class="kpi"><tr>{kh}</tr></table>'
            '<div class="st">4대 중증질환 산정특례대상자 건강보험 보장률 <span>단위: %</span></div>'
            '<table class="d"><tr><th rowspan="2">구 분</th><th colspan="3">2023년</th><th colspan="3">2024년</th></tr>'
            '<tr><th>보장률</th><th>법정<br>본인부담</th><th>비급여<br>본인부담</th>'
            '<th>보장률</th><th>법정<br>본인부담</th><th>비급여<br>본인부담</th></tr>'
            f'{tb}</table>'
            '<div class="src">4대 중증질환 보장률은 현금급여(본인부담상한제 사후환급금)를 포함함 · 원문 보도자료 붙임2</div>'
            '<div class="st">연도별 건강보험 보장률 <span>10년간 63.4% → 64.9%, 1.5%p 상승에 그침</span></div>'
            f'{ln}'
            '<div class="src">출처: 국민건강보험공단 보도자료 붙임1 「연도별 건강보험 보장률」</div>'
            '<div class="box"><b class="t">읽는 법 —</b> 뇌혈관 <b>87.9%</b> · 심장 <b>90.3%</b>는 <b>10만원 중 1만원 안쪽</b>만 내 돈이다. '
            '그런데 암은 <b>75.0%</b>, 즉 <b>4분의 1이 내 돈</b>이고 그중 대부분(<b>16.0%</b>)이 '
            '<b>건강보험이 한 푼도 안 내는 비급여</b>다.</div>'
            '</div>' + _ft("1 / 3") + '</div>')


def _p2():
    a = [("2021", 80.2, 0), ("2022", 75.7, 0), ("2023", 76.3, 0), ("2024", 75.0, 1)]
    b = [("2021", 11.0, 0), ("2022", 15.0, 0), ("2023", 14.8, 0), ("2024", 16.0, 1)]

    def _bars(data, base):
        s = '<table class="bar">'
        for y, v, hi in data:
            w = v / base * 100.0
            fc = ' fr' if hi else ''
            vc = ' vr' if hi else ''
            s += (f'<tr><td class="yr">{y}</td><td><div class="tr">'
                  f'<div class="fi{fc}" style="width:{w:.1f}%"></div></div></td>'
                  f'<td class="vv{vc}">{v:.1f}%</td></tr>')
        return s + "</table>"
    rows = [("2017", "52.5조 (7.5%)", "16.9조 (7.2%)", "14.3조 (6.6%)", "83.7조 (7.3%)", 0),
            ("2018", "59.5조 (13.3%)", "18.3조 (8.1%)", "15.5조 (8.3%)", "93.3조 (11.4%)", 0),
            ("2019", "66.3조 (11.5%)", "20.3조 (11.2%)", "16.6조 (7.0%)", "103.3조 (10.7%)", 0),
            ("2020", "67.1조 (1.2%)", "20.1조 (−1.2%)", "15.6조 (−6.2%)", "102.8조 (−0.5%)", 0),
            ("2021", "71.6조 (6.8%)", "22.1조 (9.9%)", "17.3조 (11.3%)", "111.1조 (8.1%)", 0),
            ("2022", "79.2조 (10.5%)", "23.7조 (7.5%)", "17.6조 (1.8%)", "120.6조 (8.5%)", 0),
            ("2023", "86.3조 (9.0%)", "26.5조 (11.7%)", "20.2조 (14.4%)", "133.0조 (10.3%)", 0),
            ("2024", "90.0조 (4.3%)", "26.8조 (1.0%)", "21.8조 (8.1%)", "138.6조 (4.2%)", 1)]
    tb = ""
    for r in rows:
        cls = ' class="hi"' if r[5] else ''
        tb += f'<tr{cls}><td class="l">{r[0]}</td>' + "".join(f"<td>{v}</td>" for v in r[1:5]) + "</tr>"
    return ('<div class="pg">' + _hd("암 4년 추이") + '<div class="body ci">'
            + _CS +
            '<h1>암 보장률 <span class="r">4년째 내리막</span></h1>'
            '<div class="st">암질환 보장률 연도별 추이 <span>4대 중증질환 산정특례대상자</span></div>'
            + _bars(a, 100.0) +
            '<div class="src">2021년 80.2% → 2024년 75.0% · <b>4년간 5.2%p 하락</b></div>'
            '<div class="st">같은 기간 암 비급여 본인부담률 <span>막대는 50% 기준 축</span></div>'
            + _bars(b, 50.0) +
            '<div class="src">2021년 11.0% → 2024년 <b>16.0%, 1.45배</b></div>'
            '<div class="st">연도별 부문별 진료비 규모 변화 <span>괄호는 전년 대비 증가율</span></div>'
            '<table class="d"><tr><th>연도</th><th>보험자부담금</th><th>법정 본인부담금</th>'
            '<th>비급여 진료비</th><th>총 진료비</th></tr>'
            f'{tb}</table>'
            '<div class="src">출처: 국민건강보험공단 보도자료 2쪽 「연도별 부문별 진료비 규모 변화」</div>'
            '<div class="box"><b class="t">핵심 —</b> 급여는 <b>4.3%</b> 늘 때 <b>비급여는 8.1%</b> 늘었다. '
            '공단이 부담을 늘려도 <b>비급여가 더 빨리 커져</b> 보장률이 제자리인 구조다. 그 한가운데에 <b>암</b>이 있다.</div>'
            '</div>' + _ft("2 / 3") + '</div>')


def _p3():
    kp = [("요양병원 암(산정특례)", "36.3%", '<span class="dn">▼ 1.0%p (’23 37.3)</span>'),
          ("약국 암(주상병)", "77.7%", '<span class="dn">▼ 4.0%p (’23 81.7)</span>'),
          ("상급종합 전체", "72.2%", '<span class="up">▲ 1.4%p</span>'),
          ("요양병원 전체", "67.3%", '<span class="dn">▼ 1.5%p</span>')]
    kh = "".join(f'<td><div class="kl">{a}</div><div class="kv">{b}</div>'
                 f'<div class="kd">{c}</div></td>' for a, b, c in kp)
    rows = [("상급종합", "72.2 (+1.4)", "19.1", "8.7", 0),
            ("종합병원", "66.7 (+0.6)", "21.7", "11.6", 0),
            ("병원", "51.1 (+0.9)", "18.1", "30.8", 0),
            ("요양병원", "67.3 (-1.5)", "18.5", "14.2 (+1.9)", 1),
            ("의원", "57.5 (+0.2)", "20.1", "22.4", 0),
            ("약국", "69.1 (-0.3)", "28.0", "2.9 (+0.5)", 1)]
    tb = ""
    for r in rows:
        cls = ' class="hi"' if r[4] else ''
        tb += f'<tr{cls}><td class="l">{r[0]}</td>' + "".join(f"<td>{v}</td>" for v in r[1:4]) + "</tr>"
    L = [("1", "백혈병"), ("2", "췌장의 악성신생물"), ("4", "기타 림프·조혈 악성신생물"),
         ("5", "뇌의 악성신생물"), ("6", "기관·기관지·폐의 악성신생물")]
    R = [("8", "비호지킨 림프종"), ("11", "기타 소화기관의 악성신생물"),
         ("12", "간 및 간내담관의 악성신생물"), ("13", "식도의 악성신생물"), ("25", "유방의 악성신생물")]
    tw = '<table style="width:100%;border-collapse:collapse"><tr><td style="width:49%;vertical-align:top">'
    for side in (L, R):
        tw += '<table class="d"><tr><th style="width:16%">순위</th><th>질환</th></tr>'
        for n, d in side:
            tw += f'<tr><td>{n}</td><td class="l">{d}</td></tr>'
        tw += '</table>'
        if side is L:
            tw += '</td><td style="width:2%"></td><td style="width:49%;vertical-align:top">'
    tw += '</td></tr></table>'
    return ('<div class="pg">' + _hd("치료 장소별") + '<div class="body ci">'
            + _CS +
            '<h1>어디서 치료받느냐가 <span class="r">내 돈을 정한다</span></h1>'
            '<div class="st">암 환자가 실제로 겪는 보장률</div>'
            f'<table class="kpi"><tr>{kh}</tr></table>'
            '<div class="src">공단 설명 — 요양병원·약국 보장률 하락은 「암질환 중심으로 비급여 진료비가 증가」한 것이 원인 · 원문 3쪽</div>'
            '<div class="st">요양기관 종별 건강보험 보장률 (2024년) <span>단위: % · 현금급여 미포함</span></div>'
            '<table class="d"><tr><th>요양기관 종별</th><th>건강보험 보장률</th>'
            '<th>법정 본인부담률</th><th>비급여 본인부담률</th></tr>'
            f'{tb}</table>'
            '<div class="src">원문 3쪽</div>'
            '<div class="st">1인당 중증·고액진료비 상위 30위 <span>절반 이상이 암</span></div>'
            f'{tw}'
            '<div class="src">상위 30위 내 질환 보장률 80.2%(-0.7%p) · 50위 내 78.5%(-0.5%p) · 원문 붙임3</div>'
            '<div class="box"><b class="t">상담 포인트 —</b><br>'
            '• 암 진단 후 <b>요양병원</b>으로 옮기면 보장률이 <b>36.3%</b>까지 떨어진다. 내 돈이 <b>3분의 2</b>다.<br>'
            '• 암 <b>약값</b>도 보장률이 1년 만에 <b>81.7% → 77.7%</b>로 떨어졌다.<br>'
            '• 공단이 못 내는 <b>비급여 16.0%</b>는 <b>진단비·주요치료비</b>로만 메울 수 있다.</div>'
            '</div>' + _ft("3 / 3") + '</div>')


def cancer_info_pages():
    return _p1() + _p2() + _p3()


def ga_pages_html():
    out=[GA_CSS.join(['<style>','</style>'])]
    # ★★★★★v553 (지점장 지시 2026.08.22, 영구): <b>회사별 비교표 표지도 삭제</b>한다.
    #   지점장 원문: "이건 진작버린거야" → "표지도". v328에서 본문 3장을 뺀 데 이어 표지까지 뺀다.
    #   ★HTML·PAGES 데이터는 지우지 않는다(되돌릴 때 이 스위치만 True로).
    _GA_COVER = False
    if _GA_COVER:
      out.append('<div class="pg gacover">'
        '<div class="top itop"><div class="eb">BARUM · 회사별 비교표</div>'
        '<div class="nm">회사별 <b>비교표 총정리</b></div>'
        '<div class="pgn"><b>·</b>부록</div><div class="bar"></div></div>'
        '<div class="gap-cover"><div class="gt">회사별 비교표<br><b>총정리</b></div>'
        '<div class="gd">유병자 · 비급여 통합치료비 · 항암 · 순환계 · 치매/간병 · 생·손보 수술비<br>11개 비교표 · 회사별 담보·한도·보험료<br>자료기준일 2026.06.30</div>'
        '<div class="gwarn">⚠ 본 자료는 사내교육 정리본이며 보험안내자료로 사용할 수 없습니다. 수치는 원본(폰 캡처) 판독 기반이므로, 청약·비교설명 전 반드시 최신 상품설명서로 대조 확인 요망. ? 표시 셀은 원본 화면 잘림으로 미확정.</div>'
        '</div>'
        '<div class="ft"><b>MAKEONE</b> GA채널 비교표 · 사내참고용(보험안내자료 아님)<span class="r">회사별 비교표 · 표지</span></div></div>')
    # ★★★★★v328 (지점장 지시 2026.08.02, 영구): <b>회사별 비교표 1·2·3 본문 3장을 삭제</b>한다.
    #   지점장 원문: "보험인포메이션 회사비교 1.2.3페이지는 삭제 요청했는데 그대로있다".
    #   ★<b>표지(gacover)는 지시에 없어 남긴다</b> — 뺄지는 지점장 확정 후 반영(임의 확장 금지).
    #   ★PAGES 데이터는 지우지 않는다(되돌릴 때 이 루프만 되살리면 된다).
    _GA_BODY = False
    if _GA_BODY:
        for i,grp in enumerate(PAGES,1):
            inner="".join(grp)
            out.append(f'<div class="pg"><div class="top"><div class="eb">BARUM · 회사별 비교표 (설계사 참고용)</div>'
              f'<div class="nm">회사별 비교표 <b>총정리</b></div>'
              f'<div class="pgn"><b>{i}</b>회사별 비교표</div><div class="bar"></div></div>'
              f'<div class="body gpg gpg{i}">{inner}</div>'
              f'<div class="ft"><b>MAKEONE</b> GA채널 비교표 · 사내참고용(보험안내자료 아님)<span class="r">회사별 비교표 {i}</span></div></div>')
    return "".join(out)

if __name__=='__main__':
    from weasyprint import HTML
    base="""*{margin:0;padding:0;box-sizing:border-box;font-family:'Noto Sans KR','맑은 고딕',sans-serif;}
    @page{size:A4;margin:0;} .pg{width:210mm;height:297mm;position:relative;page-break-after:always;background:#fff;}
    .top{background:#fff;color:#1B2A4A;padding:8mm 11mm 4mm;border-bottom:2.2pt solid #C9A24B;position:relative;}
    .eb{font-size:9pt;letter-spacing:2px;color:#9c7c2e;font-weight:700;} .nm{font-size:18pt;font-weight:800;color:#1B2A4A;margin-top:1.5mm;} .nm b{color:#9c7c2e;}
    .pgn{position:absolute;right:11mm;top:8mm;text-align:right;font-size:8pt;color:#6B7A90;} .pgn b{display:block;font-size:17pt;color:#1B2A4A;font-weight:800;}
    .bar{display:none;} .ft{position:absolute;bottom:0;left:0;width:100%;padding:3mm 11mm;background:#1B2A4A;color:#9FB0C6;font-size:7.5pt;} .ft b{color:#C9A24B;} .ft .r{float:right;}"""
    HTML(string=f"<html><head><meta charset='utf-8'><style>{base}</style></head><body>{ga_pages_html()}</body></html>").write_pdf('/mnt/user-data/outputs/GA비교표_설명서편입_7p.pdf')
    print("생성 완료")
