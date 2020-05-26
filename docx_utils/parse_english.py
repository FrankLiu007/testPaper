import re
from docx_utils.namespaces import namespaces as docx_nsmap
import pycnnum
import docx_utils.MyDocx as MyDocx
from . import settings
from lxml import etree
#####主要是进行版面分析，把每个题的标题、选项等部分所在的段落号，计算出来

#####找到材料题的所有材料行
def get_title_rows(doc_elements, b_row, curr_row, mode_text):
    has_material=False
    mode_text=settings.mode_text
    subject=settings.subject
    if b_row >= curr_row:
        print('开始行<=结束行')
        return []

    for i in range(b_row, curr_row):
        x = re.findall(mode_text, doc_elements[i]['text'])
        if x:
            has_material=True
            break
    if has_material:
        return list(range(b_row, curr_row))
    else:
        return []

def parse_jie(curr_row, end_row,doc, curr_status):
    b=curr_row

    while curr_row < end_row:

        txt = doc.elements[curr_row]['text'].strip()
        xx = re.findall(r'^第[一二三四五六七八九][小]{0,1}节', txt)
        if xx:

            ti=parse_ti(curr_row+1, doc, curr_status)
            return ti
        xx = re.findall(settings.mode_text, txt)  ###材料题开始了
        if xx:
            curr_status.append(xx[0])
            ti = parse_ti(curr_row + 1, doc, curr_status)
            return  ti

        curr_row += 1

    return

####一个小题结束了
def question_ends(txt):
    yy = re.findall(r'第[一二三四五六七八九]部分', txt)  ###某部分遇到了
    if yy:
        return True
    yy = re.findall(r'第[一二三四五六七八九][小]{0,1}节', txt)  # 某个小节遇到了
    if yy:
        return True
    yy = re.findall(r'[一二三四五六七八九][\s\.．、]', txt)  # 某个大题遇到了
    if yy:
        return True
    yy = re.findall(settings.mode_text, txt)  ###这个小题结束了
    if yy:
        return True
    yy = re.findall(r'^\d{1,2}[\s\.．、]', txt)  ###这个小题结束了
    if yy:
        return True
    return False

def parse_ti(curr_row,doc, curr_status):
    ti={}

    txt = doc.elements[curr_row]['text'].strip()
    xx = re.findall(r'^第[一二三四五六七八九][小]{0,1}节', txt)

    b_num, e_num=get_question_num(doc.elements,curr_row,settings.mode_text)
    curr_num=b_num
    for num in range(b_num, e_num+1):
        txt=doc.elements[num]['text'].strip()
        xx=re.findall(r'^'+str(num)+'[\s\.、．]', txt)
        if xx:
            q=parse_question(xx[0], num, doc)

def parse_question(curr_num, curr_row, doc):
    q={}
    txt = doc.elements[curr_row]['text'].strip()
    xx = re.findall(r'^' + str(curr_num) + '[\s\.、．]', txt)
    if not xx:
        print('不应出的错误: 未获得小题题号!')
    q['number']=xx[0]
    q['stem']=[curr_row]
    q['options']=[]
    while curr_row < len(doc.elements) :
        curr_row+=1
        txt=doc.elements[curr_row]['text'].strip()
        if question_ends(txt):
            q['end_row']=curr_row-1
            return q
        yy=re.findall(r'^[A-G][\s\.．、]', txt) #某个选项遇到了
        if yy:
            q['options'].append(curr_row)
        else:
            q['stem'].append(curr_row)
    return q
def parse_all(doc,start_row, end_row):

    curr_row=start_row-1
    tree=etree.Element('root')
    ti=''
    question=''
    jie=''
    dati=''
    while curr_row<=end_row:
        curr_row+=1
        txt=doc.elements[curr_row]['text'].strip()

        yy = re.findall(r'第[一二三四五六七八九]部分', txt)  ###某部分遇到了
        zz = re.findall(r'[一二三四五六七八九][\s\.．、]', txt)  # 某个大题遇到了
        if yy or zz:
            if dati:
                tree.append(dati)
            dati=etree.Element('dati')
            dati.text=str(curr_row)

        yy = re.findall(r'第[一二三四五六七八九][小]{0,1}节', txt)  # 某个小节遇到了
        if yy:
            if jie:
                dati.append(jie)
            jie = etree.Element('jie')
            jie.text = str(curr_row)

        yy = re.findall(settings.mode_text, txt)  ###这个小题结束了
        if yy:
            if jie:
                jie.append(ti)
            elif dati:
                dati.append(ti)
            ti=etree.Element('ti')
            ti.text=str(curr_row)

        yy = re.findall(r'^\d{1,2}[\s\.．、]', txt)  ###这个小题结束了
        if yy:
            if question:
                ti.append(question)
            question=etree.Element('question')

        yy = re.findall(r'^A-G[\s\.．、]', txt)  ###遇到选项了
        if yy:
            option=etree.Element('option')
            option.text=str(curr_row)
            question.append(option)
    return tree



def get_question_num(text):
    r = re.findall(settings.mode_text, text)
    b_number, e_number = r[0]
    return (int(b_number), int(e_number))

def parse_english(doc,start_row, end_row ):
    doc_elements = doc.elements
    curr_status=[]
    curr_row=start_row
    tis=[]
    curr_num=1
    tree = etree.Element('root')
    while(curr_row<=end_row):
        txt=doc_elements[curr_row]['text'].strip()
        if txt=='':
            curr_row+=1
            continue
        xx = re.findall(r'^第'+pycnnum.num2cn(curr_num)+'部分', txt)

        if xx:
            dati = etree.Element('dati')
            dati.text=str(curr_row)

            ti=parse_jie(curr_row+1, doc, curr_status)
            tis.append( ti.copy() )

            curr_num+=1
    return tis

