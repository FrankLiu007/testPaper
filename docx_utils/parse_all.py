import re
from docx_utils.namespaces import namespaces as docx_nsmap
import pycnnum
import   docx_utils.MyDocx as MyDocx
import  docx_utils.settings as settings
####版面分析的第2个大版本更新，以兼容各种特殊情况！
#####主要是进行版面分析，把每个题的标题、选项等部分所在的段落号，计算出来

def find_dati_row( doc, start_row, end_row):
    doc_elements = doc.elements
    mode_text=settings.dati_mode
    dati_row = []
    curr_num=1

    for i in range(start_row,end_row+1):
        text = doc_elements[i]['text'].strip()
        if '参考答案' in text:
            break
        rr=r'^'+mode_text[0]+pycnnum.num2cn(curr_num)+mode_text[-1]
        if re.match(rr, text ):
            dati_row.append((i, text, ''.join(mode_text)) )
            curr_num+=1
    return dati_row
#####给题目增加分数和题型
def add_score_and_titype(tis, dati_indexes, i, doc_elements):
    text=doc_elements[dati_indexes[i][0]]['text'].strip()
    score = re.findall(r'每[小]{0,1}题(\d{1,2})分', text)  ##，每个小题的分数
    ###题型 titpye
    # txt=re.sub(r'^[\d一二三四五六七八九]{1,2}[\s\.．、]{0,3}', '', text).strip()
    # e = txt.find('题')
    # category = txt[:e + 1]
    xx=re.findall(r'^[一二三四五六七八九][\s\.．、]{0,1}(.*)', text)[0]
    if ('(' in xx) or ('（'in xx):
        category=re.findall(r'(.*)[\(（]', xx)[0]
    else:
        category=xx

    for ti in tis:
        ti['category'] =category
        q_tpye = 'GENERAL'
        if '只有一项' in text or '的一项是' in text:
            q_tpye = 'SINGLE'
        if '单选' in text:
            q_tpye = 'SINGLE'
        if score:
            ti['total'] = int(score[0]) * len(ti['questions'])
        ss = 0
        for q in ti['questions']:

            q['number'] = re.findall(r'^(\d{1,2})[.．、]\s{0,}', doc_elements[q['stem'][0]]['text'])[0]
            if (not 'type' in q) or (q_tpye == ''):
                q['type'] = q_tpye
            if score:
                q['score'] = int(score[0])
            else:
                rr='[\(（]\D*(\d{1,2})分[\)）]'
                tt=''
                for stem in q['stem']:
                    tt=tt+doc_elements[stem]['text']
                s = re.findall(rr, tt)
                if s:
                    q['score'] = int(s[0])
                else:
                    q['score'] = 0
                ss = ss + q['score']
        if not score:
            ti['total'] = ss
    return tis
###判断1个段落是否为空
def is_blank_paragraph(element):
    tt=element.xpath('.//w:t/text()', namespaces=docx_nsmap)
    if  ''.join(tt).strip():
        return False
    if element.xpath('.//w:drawing', namespaces=docx_nsmap):
        return False
    if element.xpath('.//m:oMath', namespaces=docx_nsmap):
        return False
    if element.xpath('.//w:pict', namespaces=docx_nsmap):
        return False
    return True

#---------------------------------------------------
#####找到材料题的所有材料行
def get_title_rows(doc_elements, b_row, end_row, strict_mode=False):
    if b_row >= end_row:
        print('开始行<=结束行')
        return []
    title_rows=[]
    has_material=False
    for i in range(b_row, end_row):
        if is_blank_paragraph(doc_elements[i]['element']):
            continue
        txt=doc_elements[i]['text'].strip()
        if strict_mode:
            x = re.findall(settings.material_mode, doc_elements[i]['text'])
            if x:
                has_material = True

        xx = title_ends(txt)
        if xx==True:  ###正常结束
            break
        elif xx==False:  ####
            title_rows.append(i)
            continue
        else:           ##非正常结束,

            if '一' in xx:
                title_rows=[]
                continue
            else:
                if title_rows:
                    break
                else:
                    continue

    if strict_mode:  ##严格模式又没有材料
        if not has_material:
            return []

    return title_rows

####-----------------------------------
def parse_all(doc, start_row, end_row):
    dati_index=find_dati_row(doc, start_row, end_row)
    result=[]

    for i in range(0,len(dati_index)):
        b_row=dati_index[i][0]
        if i==len(dati_index)-1:
            e_row=end_row
        else:
            e_row = dati_index[i+1][0]
        dati=parse_one_titype(b_row, e_row, doc.elements)
        result.append(dati)
    return result
##获取题型
def get_category(txt):
    rr=settings.dati_mode[0]+'[一二三四五六七八九]'+settings.dati_mode[-1]
    xx=re.findall(rr+'(.*)', txt)[0]
    return  re.findall(r'(.*)[\(（?]', xx)[0]  ###获取题目的题型
def get_score(txt):
    xx=re.findall(r'每[小]{0,1}题(\d{1,2})分', txt)
    if xx:
        return xx[0]
    else:
        return ''
def guess_titype(txt):
    q_tpye='GENERAL'
    if '只有一项' in txt or '的一项是' in txt:
        q_tpye = 'SINGLE'
    if '单选' in txt:
        q_tpye = 'SINGLE'
    return q_tpye
######-------------------------处理1个大题---------------------------------------------
###这里的xiaoti_indexes为，当前大题包含的所有小题
def parse_one_titype(curr_row, end_row,  doc_elements):
    tis = []
    i = 0
    txt = doc_elements[curr_row]['text'].strip()
    category=get_category(txt)
    score=get_score(txt)
    q_tpye=guess_titype(txt)
    objective=False
    if '选择' in category or '单选' in category:
        objective=True
    curr_row=curr_row+1   ##略过大题这1行
    while curr_row<=end_row:
        txt=doc_elements[curr_row]['text'].strip()
        if dati_ends(txt):
            break
        ti=parse_ti(curr_row, end_row, doc_elements)
        curr_row=ti['end_row']
        ###要处理
        if not 'category' in ti:
            ti['category']=category
          ###add score
        if ti['questions']==[]: ###没有小问的题目的处理

            tis.append(ti)
            continue
        for q in ti['questions']:
            if score:  ##比较2种方法的分数的差异
                if 'score' in q:
                    if q['score']!=score:
                        print('2种方法得到的小题分数的不一致！')
                else:
                    q['score'] = score
            if q_tpye:  ##2种方法的题型：
                q['type'] = q_tpye
            if not 'objective' in q:
                q['objective']=objective   ####是否客观题

        tis.append(ti)

    return tis

def dati_ends(txt):
    rr=settings.dati_mode[0]+'([一二三四五六七八九])'+settings.dati_mode[-1]
    yy = re.findall(rr, txt)  ###某部分遇到了
    if yy :
        return True
    return False

def parse_ti(curr_row, end_row, doc_elements, strict_mode=False):
    ti={}
    ti['title']=[]
    ti['questions']=[]

    title_rows = get_title_rows(doc_elements, curr_row, end_row, strict_mode=strict_mode)
    if title_rows:
        curr_row = title_rows[-1] + 1

    while curr_row<=end_row:
        if is_blank_paragraph(doc_elements[curr_row]['element']):
            curr_row+=1
            continue
        txt=doc_elements[curr_row]['text'].strip()
        if ti_ends(txt):
            break
        yy = re.findall(r'^(\d{1,2})[\s\.．、]', txt)
        if yy :
            if int(yy[0])>settings.curr_ti_number:   ###刚好是下一个小题？
                q=parse_question(curr_row, end_row, doc_elements)
                ti['questions'].append(q)
            if title_rows:
                curr_row=q['end_row']
            else:
                curr_row = q['end_row']
                break
        else:
            curr_row+=1
            break
    ti['end_row']=curr_row
    ti['title']=title_rows
    return ti


def parse_question(curr_row, end_row, doc_elements):
    q={}
    q['stem']=[curr_row]
    q['options']=[]

    txt = doc_elements[curr_row]['text'].strip()
    yy = re.findall(r'^(\d{1,2})[\s\.．、]', txt)
    if yy:
        q['number']=yy[0]
        settings.curr_ti_number=int(yy[0])
    if '只有一项' in txt or '的一项是' in txt:
        q['type'] = 'SINGLE'

    yy = re.findall(r'[\(（](\d{1,2})分[\)）]', txt)
    if yy:
        q['score'] = yy[0]

    curr_row+=1
    while curr_row <= end_row:
        if is_blank_paragraph(doc_elements[curr_row]['element']):
            curr_row+=1
            continue
        txt=doc_elements[curr_row]['text'].strip()
        if question_ends(txt):
            break
        yy=re.findall(r'^[A-G][\s\.．、]', txt) #某个选项遇到了
        if yy:
            q['options'].append(curr_row)
            curr_row+=1
        else:  ##不是选项
            if q['options']:  ###不是选项，可是已经有选项了,说明到下一个题了
                break
            else:
                q['stem'].append(curr_row)
                curr_row+=1
    q['end_row'] = curr_row
    if q['options']:
        q['objective']=True

    return q

def title_ends(txt):
    rr=settings.dati_mode[0]+'[一二三四五六七八九]'+settings.dati_mode[-1]
    yy = re.findall(rr, txt)  ###某部分遇到了
    if yy :
        print('出错了，标题非正常结束，遇到大题')
        return yy[0]
    if settings.jie_mode:
        rr = settings.jie_mode[0] + '[一二三四五六七八九]' + settings.jie_mode[-1]
        yy = re.findall(rr, txt)  # 某个小节遇到了
        if yy:
            return yy[0]

    yy = re.findall(settings.xiaoti_mode, txt)  ###这个小题可能结束了
    if yy:
        if int(yy[0])>=settings.curr_ti_number:
            return True
    return False

def ti_ends(txt):
    rr=settings.dati_mode[0]+'[一二三四五六七八九]'+settings.dati_mode[-1]
    yy = re.findall(rr, txt)  ###某部分遇到了
    if yy :
        print('出错了，标题非正常结束，遇到大题')
        return yy[0]
    if settings.jie_mode:
        rr = settings.jie_mode[0] + '[一二三四五六七八九]' + settings.jie_mode[-1]
        yy = re.findall(rr, txt)  # 某个小节遇到了
        if yy:
            return yy[0]

    ####
    yy = re.findall(settings.mode_text, txt)  ###这个小题结束了
    if yy:
        return True
    return False

def question_ends(txt):
    rr=settings.dati_mode[0]+'[一二三四五六七八九]'+settings.dati_mode[-1]
    yy = re.findall(rr, txt)  ###某部分遇到了
    if yy :
        print('出错了，标题非正常结束，遇到大题')
        return yy[0]
    if settings.jie_mode:
        rr = settings.jie_mode[0] + '[一二三四五六七八九]' + settings.jie_mode[-1]
        yy = re.findall(rr, txt)  # 某个小节遇到了
        if yy:
            return yy[0]

    yy = re.findall(settings.xiaoti_mode, txt)  ###这个小题可能结束了
    if yy:
        if int(yy[0])>=settings.curr_ti_number:
            return True
    yy = re.findall(settings.xiaoti_mode, txt)  ###这个小题结束了
    if yy:
        return True

    ####
    yy = re.findall(settings.mode_text, txt)  ###这个小题结束了
    if yy:
        return True

    return False
###
def get_answer_start_row(doc):
    row=-1
    for i in range(0,len(doc.elements)) :
        if '参考答案' in doc.elements[i]['text']:
            row= i
            break
    return row-1
