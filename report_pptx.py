# -*- coding: utf-8 -*-
# BARUM 보장 진단서 PPT 생성기 v4
#  - 순서: 표지 → 보장현황(1) → AI진단·갱신/비갱신·보험료(2) → 부위별 충족률(3) → 핵심보장 CI(4)
#  - 50대 큼직 + 레드 섞기 + 표지(이름/직책 자가입력)
#  - ★계약 多 → AI진단 페이지(갱신/비갱신·보험료) 자동 여러장 분할 (김진구 6계약은 1장 유지)
import os, tempfile, math
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Mm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY=RGBColor(0x0B,0x23,0x40); NAVY2=RGBColor(0x16,0x36,0x5C); DARK=RGBColor(0x2A,0x33,0x40)
GOLD=RGBColor(0xC5,0xA0,0x52); GOLDL=RGBColor(0xE6,0xC8,0x78); GOLDD=RGBColor(0x9C,0x7C,0x32)
INK=RGBColor(0x1C,0x24,0x30); MUT=RGBColor(0x6B,0x76,0x86); LINE=RGBColor(0xD9,0xDE,0xE6)
GOOD=RGBColor(0x1F,0x7A,0x4D); GAP=RGBColor(0xC0,0x44,0x4C); BLUE=RGBColor(0x14,0x56,0xB0)
WHITE=RGBColor(0xFF,0xFF,0xFF); FOOT=RGBColor(0x9F,0xB0,0xC6)
CMTBG=RGBColor(0xF6,0xF8,0xFB); NOTEBG=RGBColor(0xF2,0xEE,0xE2); GAPBG=RGBColor(0xFB,0xEC,0xEC)
BADGEBG={'full':RGBColor(0xE4,0xF0,0xEA),'part':RGBColor(0xFB,0xF1,0xD8),'gap':RGBColor(0xF7,0xE0,0xE2)}
BADGEC ={'full':GOOD,'part':GOLDD,'gap':GAP}; BADGET={'full':'충실','part':'일부','gap':'취약'}
F="맑은 고딕"
_DONDIR=tempfile.mkdtemp()

def _hex(c): return '#%02X%02X%02X'%(c[0],c[1],c[2])
def _dc(p): return GOOD if p>=70 else (GOLDD if p>=40 else GAP)

def _donut(pct,fn):
    col=_hex(_dc(pct))
    fig,ax=plt.subplots(figsize=(1.7,1.7),dpi=170)
    ax.pie([max(pct,0.5),max(100-pct,0)],colors=[col,'#EAEEF4'],startangle=90,counterclock=False,
           wedgeprops=dict(width=0.32,edgecolor='none'))
    ax.text(0,0,f"{pct}%",ha='center',va='center',fontsize=22,fontweight='bold',color='#0B2340')
    ax.set(aspect='equal'); fig.patch.set_alpha(0)
    plt.savefig(fn,transparent=True,bbox_inches='tight',pad_inches=0); plt.close(); return fn

def _rect(s,x,y,w,h,fill=None,line=None,radius=False,lw=0.75):
    sp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,Mm(x),Mm(y),Mm(w),Mm(h))
    sp.shadow.inherit=False
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb=fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(lw)
    if radius:
        try: sp.adjustments[0]=0.06
        except: pass
    return sp

def _tf(s,x,y,w,h,anchor=MSO_ANCHOR.TOP,wrap=True):
    tb=s.shapes.add_textbox(Mm(x),Mm(y),Mm(w),Mm(h)); tf=tb.text_frame
    tf.word_wrap=wrap; tf.vertical_anchor=anchor
    tf.margin_left=Mm(1.8); tf.margin_right=Mm(1.8); tf.margin_top=Mm(0.8); tf.margin_bottom=Mm(0.8)
    return tf

def _p(tf,first=False,align=PP_ALIGN.LEFT,sa=2,sb=0,ls=None):
    p=tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment=align; p.space_after=Pt(sa); p.space_before=Pt(sb)
    if ls: p.line_spacing=ls
    return p

def _r(p,t,sz,c=INK,b=False,it=False):
    rn=p.add_run(); rn.text=t; f=rn.font; f.size=Pt(sz); f.bold=b; f.italic=it; f.color.rgb=c; f.name=F; return rn

def _head(s,cust,pgn,pgt):
    _rect(s,0,0,210,25,fill=NAVY); _rect(s,0,24.2,210,1.4,fill=GOLD)
    tf=_tf(s,11,4,150,18)
    p=_p(tf,first=True); _r(p,"MAKEONE · 보장분석 리포트",9,GOLDL,b=True)
    p2=_p(tf,sa=0); _r(p2,f"{cust} ",21,WHITE,b=True); _r(p2,"고객님",21,GOLDL,b=True); _r(p2," 보장 진단서",21,WHITE,b=True)
    tf2=_tf(s,150,4,49,18)
    pa=_p(tf2,first=True,align=PP_ALIGN.RIGHT); _r(pa,str(pgn),22,WHITE)
    pb=_p(tf2,align=PP_ALIGN.RIGHT,sa=0); _r(pb,pgt,9,FOOT)

def _foot(s,cust,extra=""):
    _rect(s,0,288,210,9,fill=NAVY)
    tf=_tf(s,11,288,120,9,anchor=MSO_ANCHOR.MIDDLE)
    p=_p(tf,first=True); _r(p,"MAKEONE",8.5,GOLD,b=True); _r(p," 보장분석 자동화",8.5,FOOT)
    tf2=_tf(s,120,288,79,9,anchor=MSO_ANCHOR.MIDDLE)
    p2=_p(tf2,first=True,align=PP_ALIGN.RIGHT); _r(p2,f"{cust} 고객님{extra}",8.5,FOOT)

def _sect(s,x,y,w,title,en=""):
    tf=_tf(s,x,y,w,8); p=_p(tf,first=True); _r(p,title,14,NAVY,b=True)
    if en: _r(p,"  "+en,9,GOLDD,b=True)
    _rect(s,x,y+7.4,w,0.6,fill=NAVY)

def _slide(prs): return prs.slides.add_slide(prs.slide_layouts[6])

# ════════════ 표지 ════════════
def _cover(prs,rep):
    s=_slide(prs)
    _rect(s,0,0,210,46,fill=NAVY); _rect(s,0,44.6,210,1.6,fill=GOLD)
    tf=_tf(s,16,10,178,32,anchor=MSO_ANCHOR.MIDDLE)
    p=_p(tf,first=True); _r(p,"MAKEONE · 보장분석 리포트",10.5,GOLDL,b=True)
    p2=_p(tf,sa=0,sb=3); _r(p2,f"{rep['client']} ",30,WHITE,b=True); _r(p2,"고객님",30,GOLDL,b=True); _r(p2," 보장 진단서",30,WHITE,b=True)
    _rect(s,16,250,178,0.4,fill=LINE)
    tfs=_tf(s,16,255,178,30)
    ps=_p(tfs,first=True); _r(ps,"MAKEONE",18,GOLD,b=True)
    pn=_p(tfs,sb=7); _r(pn,"________________________     ",14,RGBColor(0xB0,0xB0,0xB0)); _r(pn,"( 이름 )",11,MUT)
    pt=_p(tfs,sb=5); _r(pt,"________________________     ",14,RGBColor(0xB0,0xB0,0xB0)); _r(pt,"( 직책 )",11,MUT)

# ════════════ 1 보장현황 ════════════
def _s1(prs,rep):
    s=_slide(prs); _head(s,rep['client'],1,"보장 현황")
    cards=[("보유 계약",f"{rep['n_contract']}","건",NAVY),("월 납입보험료",f"{rep['premium']:,}","원",NAVY),
           ("갱신 / 비갱신",f"{rep['renew']} / {rep['nonrenew']}","",NAVY),("보장 공백",f"{rep['gap_count']}","영역",GAP)]
    w=44; g=3.3; x0=11
    for i,(k,v,u,vc) in enumerate(cards):
        x=x0+i*(w+g); _rect(s,x,29,w,17,fill=WHITE,line=LINE,radius=True)
        tf=_tf(s,x,29.3,w,16.4,anchor=MSO_ANCHOR.MIDDLE)
        p=_p(tf,first=True,align=PP_ALIGN.CENTER); _r(p,k,10,MUT,b=True)
        p2=_p(tf,align=PP_ALIGN.CENTER,sa=0); _r(p2,v,18,vc,b=True)
        if u: _r(p2,f" {u}",10,MUT)
    _sect(s,11,49,188,"보장 현황","CATEGORY COVERAGE")
    cov=rep['coverage'][:10]; cols=2; cw=92; g2=4; x0=11; y0=59.5; rowh=42.5; vgap=2.4
    for i,c in enumerate(cov):
        col=i%cols; row=i//cols; x=x0+col*(cw+g2); y=y0+row*(rowh+vgap); st=c['status']
        body=GAPBG if st=='gap' else WHITE
        _rect(s,x,y,cw,rowh,fill=body,line=LINE,radius=True)
        _rect(s,x,y,3.2,rowh,fill=BADGEC[st])
        tf=_tf(s,x+5.5,y+2,cw-28,9); p=_p(tf,first=True); _r(p,c['name'],15,NAVY,b=True)
        _rect(s,x+cw-23,y+2.4,20,7.2,fill=BADGEBG[st],radius=True)
        tfb=_tf(s,x+cw-23,y+2.2,20,7.2,anchor=MSO_ANCHOR.MIDDLE); pb=_p(tfb,first=True,align=PP_ALIGN.CENTER); _r(pb,BADGET[st],11,BADGEC[st],b=True)
        tfi=_tf(s,x+5.5,y+11.5,cw-8,rowh-12); items=c['items']
        if items and items[0].get('none'):
            pp=_p(tfi,first=True); _r(pp,"미가입",13,GAP,b=True)
        else:
            first=True
            for it in items[:5]:
                pp=_p(tfi,first=first,sa=2.5); first=False
                _r(pp,f"{it['t']}  ",12.5,INK); _r(pp,it['v'],12.5,(BLUE if it.get('blue') else NAVY),b=True)
    _foot(s,rep['client'])

# ── 헬퍼: AI진단 요약 / 갱신·비갱신 박스 / 보험료 막대 ──
def _ai_summary(s,rep,y):
    _sect(s,11,y,188,"AI 진단 요약","DIAGNOSIS"); y+=10
    for idx,(hd,arr,hc,empty) in enumerate([("✓ 보유 강점",rep['strength'],GOOD,None),
        ("! 보장 공백",rep['weak'],GAP,"주요 공백 없음 — 핵심담보 균형 양호")]):
        x=11+idx*96; _rect(s,x,y,92,47,fill=(GAPBG if hc==GAP else WHITE),line=LINE,radius=True); _rect(s,x,y,92,9,fill=hc,radius=True)
        tfh=_tf(s,x+3,y-0.2,86,9,anchor=MSO_ANCHOR.MIDDLE); ph=_p(tfh,first=True); _r(ph,hd,13,WHITE,b=True)
        tf=_tf(s,x+3,y+9.5,86,36)
        if arr:
            first=True
            for a in arr[:4]:
                pp=_p(tf,first=first,sa=3); first=False
                _r(pp,f"{a['h']}  ",12.5,NAVY,b=True); _r(pp,a['d'].replace('충족률 ','').replace(' — ',' · '),11,MUT)
        else:
            pp=_p(tf,first=True); _r(pp,empty,12,GOOD,b=True)
    return y+47

def _renew_boxes(s,rep,rlist,nlist,y,bh,cont=False,more=False):
    title="갱신 / 비갱신 구조"+("  (계속)" if cont else "  (※ 이 칸에서 직접 수정)")
    _sect(s,11,y,188,title,"RENEWAL"); y+=10
    for idx,(lab,arr,hc,col,sub) in enumerate([
        (f"갱신형 {rep['renew']}건",rlist,BLUE,BLUE,"보험료 인상 가능"),
        (f"비갱신형 {rep['nonrenew']}건",nlist,DARK,NAVY,"만기까지 고정")]):
        x=11+idx*96; _rect(s,x,y,92,bh,fill=WHITE,line=LINE,radius=True); _rect(s,x,y,92,9,fill=hc,radius=True)
        tfh=_tf(s,x+3,y-0.2,86,9,anchor=MSO_ANCHOR.MIDDLE); ph=_p(tfh,first=True)
        _r(ph,lab,13,WHITE,b=True); _r(ph,f"   {sub}",9,RGBColor(0xD9,0xDE,0xE6))
        tf=_tf(s,x+3,y+10,86,bh-12); first=True
        for h in arr:
            nm=h['nm'].rstrip().rstrip('([').rstrip()
            pp=_p(tf,first=first,sa=4); first=False; _r(pp,nm,10,col,b=True); _r(pp,f"   {h['v']}",10,MUT)
        if not arr:
            pp=_p(tf,first=True); _r(pp,"—",12,MUT)
        if more and arr:
            pp=_p(tf,sa=0); _r(pp,"…다음 장 계속",10,GOLDD,it=True)
    return y+bh

def _premium_block(s,rep,bars,y,note=True):
    if note:
        tf=_tf(s,11,y-9,188,7); p=_p(tf,first=True)
        _r(p,"※ 자동분류는 100%가 아닙니다. 증권 확인 후 위 칸에서 계약을 옮기거나 글자색(파랑=갱신/검정=비갱신)을 직접 고치세요.",10.5,GAP)
    _sect(s,11,y,188,"월 보험료 구성","PREMIUM"); y+=10
    mx=max([b['amt'] for b in bars],default=1) or 1
    for b in bars:
        tf=_tf(s,11,y-0.5,42,7,anchor=MSO_ANCHOR.MIDDLE); p=_p(tf,first=True); _r(p,b['nm'][:12],11,INK)
        _rect(s,54,y+1,118,4.6,fill=RGBColor(0xEC,0xEE,0xF1)); fw=max(2,118*b['amt']/mx)
        _rect(s,54,y+1,fw,4.6,fill=(BLUE if b['renew'] else NAVY2))
        tf2=_tf(s,174,y-0.5,25,7,anchor=MSO_ANCHOR.MIDDLE); p2=_p(tf2,first=True,align=PP_ALIGN.RIGHT)
        _r(p2,f"{b['amt']:,}",11,(BLUE if b['renew'] else NAVY),b=True); y+=8

# ════════════ 2 월보험료 → 갱신/비갱신 → AI진단 (★자동 분할) ════════════
def _s2(prs,rep):
    cust=rep['client']; renew=rep['renew_list']; nonren=rep['nonrenew_list']; bars=rep['premium_bars']
    fits = len(bars)<=9 and len(renew)<=8 and len(nonren)<=8
    if fits:
        # ── 1장: 월보험료(위) → 갱신/비갱신(중) → AI진단(아래) ──
        s=_slide(prs); _head(s,cust,2,"AI 진단")
        _premium_block(s,rep,bars,31,note=False)
        tf=_tf(s,11,102,188,6); p=_p(tf,first=True)
        _r(p,"※ 자동분류는 100%가 아닙니다. 증권 확인 후 아래 칸에서 계약을 옮기거나 글자색(파랑=갱신/검정=비갱신)을 직접 고치세요.",10,GAP)
        _renew_boxes(s,rep,renew,nonren,114,bh=98,more=False)
        _ai_summary(s,rep,227)
        _foot(s,cust); return
    # ── 계약 多: 월보험료 → 갱신/비갱신 → AI진단, 각 섹션 페이지 분할 ──
    PB=28
    for pi in range(0,len(bars),PB):
        s=_slide(prs); _head(s,cust,2,"월 보험료 구성")
        _premium_block(s,rep,bars[pi:pi+PB],38,note=False); _foot(s,cust)
    CAP=15
    npg=max(math.ceil(len(renew)/CAP), math.ceil(len(nonren)/CAP)) if (renew or nonren) else 1
    for pi in range(npg):
        s=_slide(prs); _head(s,cust,2,"갱신 / 비갱신")
        more=(pi+1)*CAP<len(renew) or (pi+1)*CAP<len(nonren)
        _renew_boxes(s,rep,renew[pi*CAP:(pi+1)*CAP],nonren[pi*CAP:(pi+1)*CAP],30,bh=245,cont=(pi>0),more=more)
        _foot(s,cust)
    s=_slide(prs); _head(s,cust,2,"AI 진단"); _ai_summary(s,rep,40); _foot(s,cust)

# ════════════ 3 부위별 충족률 ════════════
def _s4(prs,rep):
    s=_slide(prs); _head(s,rep['client'],3,"부위별 충족률")
    _sect(s,11,30,188,"부위별 충족률","COVERAGE LEVEL")
    det=rep['donut_detail'][:10]; cols=5; cw=37.6; x0=11; y0=41; dsz=29
    for i,d in enumerate(det):
        col=i%cols; row=i//cols; x=x0+col*cw; yy=y0+row*40
        fn=os.path.join(_DONDIR,f"d{i}.png"); _donut(d['pct'],fn)
        s.shapes.add_picture(fn,Mm(x+(cw-dsz)/2),Mm(yy),Mm(dsz),Mm(dsz))
        tf=_tf(s,x,yy+dsz,cw,7); p=_p(tf,first=True,align=PP_ALIGN.CENTER); _r(p,d['name'],12,NAVY,b=True)
    yL=y0+(1 if len(det)<=5 else 2)*40-1
    for j,(t,c) in enumerate([("충실 70%↑",GOOD),("보강권장 40–69%",GOLDD),("취약 40%↓",GAP)]):
        xx=54+j*44; _rect(s,xx,yL+1,4,4,fill=c)
        tf=_tf(s,xx+5,yL-0.6,42,6,anchor=MSO_ANCHOR.MIDDLE); p=_p(tf,first=True); _r(p,t,10,INK)
    yt=yL+10
    _sect(s,11,yt,188,"충족률 산정 근거",f"보유 ÷ {rep['band_label']} 권장"); yt+=9.5
    heads=["영역","보유",f"권장({rep['band_label']})","충족률"]; cwid=[40,58,58,32]
    _rect(s,11,yt,188,8,fill=NAVY); xx=11
    for hh,wd in zip(heads,cwid):
        tf=_tf(s,xx,yt,wd,8,anchor=MSO_ANCHOR.MIDDLE); p=_p(tf,first=True,align=PP_ALIGN.CENTER); _r(p,hh,11,WHITE,b=True); xx+=wd
    yt+=8
    for k,r in enumerate(rep['donut_detail'][:10]):
        xx=11; vals=[r['name'],r['have'],r['rec'],f"{r['pct']}%"]; ccs=[INK,INK,MUT,_dc(r['pct'])]; bbs=[True,False,False,True]
        rb=RGBColor(0xF7,0xF9,0xFB) if k%2 else WHITE
        _rect(s,11,yt,188,7.4,fill=rb,line=LINE)
        for v,wd,cc,bb in zip(vals,cwid,ccs,bbs):
            tf=_tf(s,xx,yt,wd,7.4,anchor=MSO_ANCHOR.MIDDLE); p=_p(tf,first=True,align=PP_ALIGN.CENTER); _r(p,v,11,cc,b=bb); xx+=wd
        yt+=7.4
    yt+=2.5; _rect(s,11,yt,188,18,fill=NOTEBG,radius=True); _rect(s,11,yt,1.4,18,fill=GOLD)
    tf=_tf(s,15,yt+1.5,181,16); p=_p(tf,first=True,ls=1.4)
    _r(p,"※ 충족률 = 보유 ÷ 연령밴드 권장액 × 100 ",9.5,INK,b=True)
    _r(p,f"(상한 100%). 권장액은 업계 적정 가입금액 가이드 기준 {rep['band_label']} 표준밴드 적용. 운전자·실손·일당·응급실은 핵심담보 보유개수 기준. 개인 소득·가족력에 따라 상담으로 조정됩니다.",9.5,RGBColor(0x5C,0x53,0x40))
    if not rep['age_known']:
        _r(p," [확인] 나이·성별 미추출 → 40대 표준밴드 산정.",9.5,GAP,b=True)
    _foot(s,rep['client'])

# ════════════ 4 핵심보장 CI ════════════
def _s3(prs,rep):
    s=_slide(prs); _head(s,rep['client'],4,"핵심 보장 분석")
    y=30; ci=rep['ci']
    if ci['present']:
        _sect(s,11,y,188,"CI · 생명보험 선지급 분석","CRITICAL ILLNESS"); y+=10
        _rect(s,11,y,188,28,fill=RGBColor(0xEA,0xF2,0xFB),line=LINE,radius=True)
        tf=_tf(s,14,y+2,182,24); p=_p(tf,first=True,sa=5)
        _r(p,f"주계약 사망 {ci['samang']}",14,NAVY,b=True); _r(p,f"   ·  선지급 {ci['rate']}%   ·  잔여사망 {ci['residual']}",12,INK)
        if ci['items']:
            pp=_p(tf)
            for it in ci['items']: _r(pp,f"{it['t']} ",13,INK,b=True); _r(pp,f"{it['v']}    ",13,BLUE,b=True)
        y+=32
    _sect(s,11,y,188,"주요 치료비 정리","TREATMENT BENEFITS"); y+=10
    chiryo=rep['chiryo']; cols=2; cw=92; g=4; rowh=12
    for i,c in enumerate(chiryo):
        col=i%cols; row=i//cols; x=11+col*(cw+g); yy=y+row*rowh; none=(c['value']=='미가입')
        _rect(s,x,yy,cw,rowh-1.5,fill=(GAPBG if none else WHITE),line=LINE,radius=True)
        tf=_tf(s,x+3,yy,cw-6,rowh-1.5,anchor=MSO_ANCHOR.MIDDLE)
        p=_p(tf,first=True); _r(p,c['name'],12.5,INK,b=True); _r(p,f"   {c['value']}",12.5,(GAP if none else NAVY),b=True)
    y+=((len(chiryo)+1)//2)*rowh+5
    _sect(s,11,y,188,"보장 진단 코멘트","SUMMARY"); y+=10
    _rect(s,11,y,188,48,fill=CMTBG,line=LINE,radius=True); _rect(s,11,y,1.4,48,fill=NAVY)
    tf=_tf(s,15,y+3,181,43)
    sh=[a['h'] for a in rep['strength']]; wk=[a['h'] for a in rep['weak']]
    cm=f"총 {rep['n_contract']}건, 월 납입보험료 {rep['premium']:,}원입니다. "
    if sh: cm+=f"특히 {' · '.join(sh)} 영역은 핵심담보를 충실히 보유했습니다. "
    if wk: cm+=f"반면 {' · '.join(wk)} 영역은 충족률이 낮아 보강을 권합니다. "
    cm+="자세한 사항은 증권 및 상담을 통해 확인하시기 바랍니다."
    p=_p(tf,first=True,ls=1.7); _r(p,cm,13,INK)
    _foot(s,rep['client'])

def build_report_pptx(rep,out):
    prs=Presentation(); prs.slide_width=Mm(210); prs.slide_height=Mm(297)
    _cover(prs,rep); _s1(prs,rep); _s2(prs,rep); _s4(prs,rep); _s3(prs,rep)  # 충족률(_s4) → 핵심보장(_s3) 순
    prs.save(out); return out

if __name__=="__main__":
    import json
    rep=json.load(open('rep_kjg_real.json')); build_report_pptx(rep,'보장진단서_김진구.pptx'); print("saved")
