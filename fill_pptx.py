# -*- coding: utf-8 -*-
from pptx import Presentation
import re

def won(man):
    man=int(round(man))
    if man==0: return '미가입'
    eok=man//10000; rest=man%10000
    s=''
    if eok: s+=f'{eok}억'
    if rest:
        if rest%1000==0: s+=f'{rest//1000}천만'
        else: s+=f'{rest:,}만'
    else:
        if not eok: s='0'
    return s or '미가입'

# 폼 라벨 → (카테고리,담보) 매핑 (한장보장 슬라이드1)
LABELMAP={
 '암진단비':('암','암진단비'),'유사암':('암','유사암(갑.기.경.제)'),'전이암':('암','전이암진단비'),
 '항암치료':('암','항암치료비'),'암수술':('암','암수술'),'암주요':('암','암주요치료비'),
 '뇌혈관':('뇌혈관','뇌혈관진단비'),'뇌졸증':('뇌혈관','뇌졸증진단비'),'뇌혈관수술':('뇌혈관','뇌혈관수술비'),
 '허혈성수술':('심장','심장수술비'),'허혈성':('심장','허혈성'),'급성심근경색':('심장','급성심근경색'),'심장 수술':('심장','심장수술비'),
 '질병수술':('수술비','질병수술비'),'상해수술':('수술비','상해수술비'),'119대 수술':('수술비','119대 수술비'),
 '골절':('골절','골절진단비'),'5대골절':('골절','5대골절진단비'),'깁스':('깁스','깁스치료비'),'응급실':('응급실','응급실내원비(응급)'),
 '대인':('운전자','교통사고벌금(대인)'),'대물':('운전자','교통사고벌금(대물)'),'합의금':('운전자','교통사고처리지원금'),
 '변호사비':('운전자','변호사선임비용'),'자부상':('운전자','자동차부상위로금'),'6주미만':('운전자','6주미만처리지원금'),
 '일상배상책임':('배상','일상생활배상책임'),'상해사망':('사망','상해사망'),
 '입원':('실손','실손입원'),'통원':('실손','실손통원'),'MRI':('실손','MRI'),'도수치료':('실손','도수치료'),'비급여주사':('실손','비급여주사'),
 '치매':('요양치매','경증이상치매진단비'),'시설':('요양치매','시설급여'),
 '크라운':('화재/치아','크라운'),'임플란트':('화재/치아','임플란트'),'화재보험':('화재/치아','화재벌금/화재보험'),
}

def totals(data):
    t={}
    for r in data.get('rows',[]):
        t[(r['category'],r['name'])]=sum(r['values'].values())
    return t

def fill_pptx(data, template, path):
    prs=Presentation(template); t=totals(data); s=prs.slides[0]
    def proc(shapes):
        for sh in shapes:
            if sh.shape_type==6: proc(sh.shapes); continue
            if not sh.has_text_frame: continue
            txt=sh.text_frame.text
            if ':' not in txt: continue
            new=txt
            for label,key in LABELMAP.items():
                val=t.get(key)
                if val is None: continue
                # "label : 기존값" 패턴 치환
                pat=re.compile(re.escape(label)+r'\s*:\s*[^/\n]+')
                if pat.search(new):
                    new=pat.sub(f'{label} : {won(val)}',new,count=1)
            if new!=txt:
                # 첫 run에 통째로
                for p in sh.text_frame.paragraphs:
                    for run in p.runs: run.text=''
                sh.text_frame.paragraphs[0].text=new
    proc(s.shapes)
    prs.save(path)
    return path
