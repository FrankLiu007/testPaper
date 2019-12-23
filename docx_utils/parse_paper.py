import re
from lxml import etree
import uuid
from docx_utils.namespaces import namespaces as docx_nsmap
import pycnnum
import docx_utils.MyDocx as MyDocx
#####主要是进行版面分析，把每个题的标题、选项等部分所在的段落号，计算出来

# 判断是否为标题
def isNumber(char):
    num1 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    num2 = ['一', '二', '三', '四', '五', '六', '七', '八', '九']
    if char in num1 or char in num2:
        return True
    return False


# ###计算模式字符串
# def get_mode_string(text):
#     tt = text[0:3]
#     k = 0
#     mode_string = ''
#     for i in range(0, 3):
#         if isNumber(text[i]):
#             k = i
#             break
#     if k == 0:
#         if text[1] in ['\t', ' ']:
#             mode_string = ''
#     return mode_string

#分析试卷的版面
# 大题模式为中文数字一 ， 小题模式为 阿拉伯数字 1
# def analys_layout(doc):
#     '''
#     :param doc:
#     :return: tree
#     格式：[ (row,paragraphs[row].text, mode_text),(...)]
#     '''
#     pars = []  ###可能
#     paragraphs = doc.paragraphs
#     row = 0
#
#     ##找出所有带数字的行和行号(  中文数字 一， 阿拉伯数字 1 )
#     while (row < len(paragraphs)):
#         text = paragraphs[row].text.strip()
#         for n in range(0, 4):
#             if n >= len(text):
#                 continue
#             if isNumber(text[n]):
#                 pars.append((n, row, text))  ##n--数字出现的位置，i--段落号
#                 break
#         row = row + 1
#
#     ##找所有模式，行号为 1 或者 一
#     modes = []
#     for n, row, text in pars:
#         if n == 0:
#             if text[n] in ['1', '一'] and (not isNumber(text[n + 1])):
#                 modes.append((n, row, text, text[n]))
#         else:
#             if text[n] in ['1', '一'] and (not isNumber(text[n - 1])) and (not isNumber(text[n + 1])):
#                 modes.append((n, row, text, text[n]))
#
#     tree = []
#     for i in range(0, len(modes)):  ##每个模式
#         n, begin_row, text, mode_text = modes[i]  ##n为行号
#         tmp_mode = []
#         tmp_mode.append((begin_row, text, mode_text))
#
#         next_number = get_next_number(mode_text)
#         if text[n+1]=='．':
#             text=text[:n+1]+'.'+text[n+2:] ###全角半角“.”的转换，防止用户没有正确使用“.”号
#
#         for row in range(begin_row + 1, len(paragraphs)):  ###查找某个模式的所有的值
#             tt=paragraphs[row].text.strip()
#             if len(tt)<n+2:
#                 continue
#             x=tt.find('．')
#             if x!=-1 and x <=n+2:
#                 tt=tt[:x]+'.'+tt[x+1:]   ###全角半角"."的转换，防止用户没有正确使用"."号
#
#             if tt.startswith(text[:n] + mode_text + text[n + 1]):  ##发现了相同的模式，退出
#                 break
#             if tt.startswith(text[:n] + next_number + text[n + 1]):
#                 # print('next_string:', text[:n] + next_number + text[n + 1], paragraphs[row].text)
#                 next_number = get_next_number(next_number)
#                 tmp_mode.append((row, tt, text[:n] + mode_text + text[n + 1]))
#         if len(tmp_mode) > 1:
#             tree.append(tmp_mode.copy())
#
#     return tree


##获取1个选项，[A-G]. 形式的
def get_option(text):
    text = text.strip()
    indexes = []
    options = []
    for item in re.finditer(r'[A-G]\s{0,2}[\.．]', text):
        indexes.append((item.group(), item.span()))
    print('in get_option,text=', text)
    if indexes[0][1] != (0, 2):  ###校检结果，保证
        print('获取选择题选项出错，请检查试题格式')
        print('text=', text)

    i = 0
    while (i < len(indexes)):
        b = indexes[i][1][0]
        if i == len(indexes) - 1:
            options.append((option_text[0], text[b:].strip()))
            break
        e = indexes[i + 1][1][0]
        option_text = indexes[i][0]
        options.append((option_text[0], text[b:e].strip()))
        b = e
        i = i + 1
    return options

###判断1个段落是否为空
def is_blank_paragraph(element):

    tt=element.xpath('.//w:t/text()')
    if  ''.join(tt).strip():
        return False
    if element.xpath('.//w:drawing', namespaces=docx_nsmap):
        return False
    if element.xpath('.//m:oMath', namespaces=docx_nsmap):
        return False
    if element.xpath('.//w:pict', namespaces=docx_nsmap):
        return False
    return True


def parse_tis(dati_indexes,xiaoti_indexes, doc_elements, mode_text):  ##处理1个大题，例如 “一、选择题”
    tis={}        ###ti={ 'title':'title' , questions:[]}
    i=0

    curr_dati_num=0
    for  j in range(0, len( xiaoti_indexes)):
        if dati_indexes[curr_dati_num][0] < xiaoti_indexes[j][0]:
            if dati_indexes[curr_dati_num+1][0] > xiaoti_indexes[j][0]:
                ###该小题属于当前大题
                pass
            else:
                ###该小题不属于当前大题
                curr_dati_num+=1

        next=xiaoti_indexes[j+1]
#-------------------------------------------------------
##获取材料题，到底有多少小问（题）
def get_question_quantity(doc_elements, title_rows, mode_text):
    text=''
    for row in title_rows:
        text=text+''.join( doc_elements[row]['element'].xpath('.//w:t/text()', namespaces=docx_nsmap) )
    r = re.findall(mode_text, text)
    start_ti_number, stop_ti_number = r[0]
    start_ti_number = int(start_ti_number)
    stop_ti_number = int(stop_ti_number)
    return stop_ti_number - start_ti_number + 1

#---------------------------------------------------
#####找到材料题的所有材料行
def get_title_rows(doc_elements, b_row, curr_row, mode_text):
    has_material=False

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
######-------------------------处理1个大题---------------------------------------------
###这里的xiaoti_indexes为，当前大题包含的所有小题
def parse_one_titype(curr_row, end_row, xiaoti_indexes, doc_elements, mode_text):
    tis = []
    i = 0
    curr_row+=1

    while i < len(xiaoti_indexes):
        if i==len(xiaoti_indexes)-1:   ###最后一个小题
            question = parse_question(xiaoti_indexes[i],  end_row, doc_elements)
            tis.append({'questions':[question], 'title':[]})
            i=i+1
            continue

        if xiaoti_indexes[i][0]>curr_row:   ###该小题可能是材料题的第X问
            title_rows=get_title_rows(doc_elements,curr_row, xiaoti_indexes[i][0], mode_text)

            if title_rows:  ###处理多问的小题
                ti = {'questions':[], 'title':[]}
                n=get_question_quantity(doc_elements, title_rows  ,mode_text)
                for j in range(0,n):
                    if i+j==len(xiaoti_indexes)-1:
                        question=parse_question(xiaoti_indexes[i+j], end_row, doc_elements)
                    else:
                        question = parse_question(xiaoti_indexes[i + j], xiaoti_indexes[i + j+1][0]-1, doc_elements)
                    ti['questions'].append(question.copy())
                curr_row=question['end_row']+1
                ti['title'] = title_rows
                tis.append(ti.copy())
                i=i+n
            else: ####不可能是最后一个小题，最后1个小题，已经在循环开始时，就被处理了！
                question = parse_question(xiaoti_indexes[i],  xiaoti_indexes[i + 1][0] - 1, doc_elements)
                curr_row = question['end_row'] + 1
                tis.append({'questions':[question], 'title':[]})
                i = i + 1
        elif xiaoti_indexes[i][0]==curr_row:    ###不是材料题
            question = parse_question(xiaoti_indexes[i],  xiaoti_indexes[i+1][0]-1, doc_elements)
            curr_row=question['end_row']+1
            tis.append({'questions':[question], 'title':[]})
            i=i+1

    return tis


####处理1个题
def isObjective(curr_row, next_row, children):
    # print('next_row=',next_row)
    for i in range(curr_row, next_row+1):
        text = children[i]['text'].strip()
        if re.match(r'[A-G][．\.]', text):
            return (True, i)
    return (False, -1)


####解析1道题
def parse_question(xiaoti, end_row, doc_elements):
    # curr_row=xiaoti_indexes
    mode_text = r'(\d{1,2})[～\-~](\d{1,2})[小]{0,1}题'
    start_row=xiaoti[0]
    objective, index = isObjective(start_row, end_row, doc_elements)
    question = {}
    question['end_row']=0
    question['objective']=objective
    question['stem'] = []
    if objective:
        options = []
        question['stem']=list(range(start_row, index))
        for j in range(index, end_row+1):
            if re.match(r'[A-G][．\.]', doc_elements[j]['text']):
                options.append(j)
                question['end_row']=j

        question['options'] = options
    else:
        question['stem']=list(range(start_row, end_row))
        question['end_row'] = end_row

    return question

def check_run(child):
    i = 0
    # print('child=', child)
    run = child.__copy__()
    rPr = run.xpath('.//w:rPr', namespaces=run.nsmap)  ##删除run的属性

    if len(rPr) > 1:
        print('docx格式出错了，len(w:rPr)!=1')
    elif len(rPr) == 1:
        run.remove(rPr[0])
    # else: 没有rPr,啥都不用做

    wt = run.xpath('.//w:t', namespaces=run.nsmap)
    wdrawing = run.xpath('.//w:drawing', namespaces=run.nsmap)
    moMath = run.xpath('.//m:oMath', namespaces=docx_nsmap)
    i = len(wt) + len(wdrawing) + len(moMath)
    return i


###删除空白行
def remove_blank_paragraph(doc):
    is_blank = True

    i = len(doc.paragraphs) - 1
    nn = 0  ##number of blank paragraphs
    while (i >= 0):
        text = ''
        paragraph = doc.paragraphs[i]
        for run in paragraph.runs:
            wdrawing = run.xpath('.//w:drawing', namespaces=run.nsmap)
            if len(wdrawing) > 0:
                is_blank = False
                break
            moMath = run.xpath('.//m:oMath', namespaces=docx_nsmap)
            if len(moMath) > 0:
                is_blank = False
                break

        if paragraph.text.strip() == '':
            is_blank = True
        if is_blank:
            doc.paragraphs.remove(paragraph)

            nn = nn + 1
        i = i - 1
    print('空白段落总数：', nn)

    return 0

def find_xiaoti_row(doc, start_row, end_row):
    xiaoti_row=[]
    doc_elements = doc.elements

    curr_num=1
    for n in range(start_row,end_row+1):
        text = doc_elements[n]['text'].strip()
        if '参考答案' in text:
            break
        rr = r'^(' + str(curr_num) + r')[\s\.、．]'
        tt=re.findall(rr, text)
        if tt:
            xiaoti_row.append((n, text, '1.',tt[0]))
            curr_num += 1

    return xiaoti_row

####找出大题的行
def find_dati_row( doc, start_row, end_row):
    doc_elements = doc.elements
    dati_row = []
    curr_num=1
    for i in range(start_row,end_row+1):
        text = doc_elements[i]['text'].strip()
        if '参考答案' in text:
            break
        rr=r'^'+pycnnum.num2cn(curr_num)+r'[\s\.、．]'
        if re.match(rr, text ):
            dati_row.append((i, text, '一、'))
            curr_num+=1
    return dati_row

###获取某个大题包含的小题（一包含哪几个1，2，3）
def get_dati_children(dati_indexes, i, xiaoti_indexes):
    xiaoti_list = []
    curr_row=dati_indexes[i][0]

    if i==len(dati_indexes)-1:   ###最后一个大题的情况
        next_row=xiaoti_indexes[-1][0]+1
    else:
        next_row = dati_indexes[i + 1][0]
    for jj in range(0, len(xiaoti_indexes)):
        if xiaoti_indexes[jj][0] > curr_row and xiaoti_indexes[jj][0] < next_row:
            xiaoti_list.append(xiaoti_indexes[jj])
    return xiaoti_list

def AnalysQuestion(doc,start_row, end_row,mode_text ):
    doc_elements = doc.elements
    dati_indexes=find_dati_row( doc, start_row, end_row)
    xiaoti_indexes=find_xiaoti_row( doc, dati_indexes[0][0]+1, end_row)
    if (len(dati_indexes)==0) or (len(xiaoti_indexes)==0):
        return ()

    ####获取所有大题的  小题
    tis = []
    curr_row, text, mode_tt = dati_indexes[0]
    i = 0
    all_ti = []
    while i < len(dati_indexes):  ##处理1种题型

        if i==len(dati_indexes) - 1:  ##如果是最后一个大题
            next_row=end_row
            xiaotis=get_dati_children(dati_indexes, i, xiaoti_indexes)
        else:
            next_row, next_text, mode_tt = dati_indexes[i + 1]
            xiaotis=get_dati_children(dati_indexes, i, xiaoti_indexes)
        tis = parse_one_titype(curr_row, next_row, xiaotis, doc_elements, mode_text)  ##处理1种题型的所有题目
        all_ti.append(tis.copy())
        i = i + 1
        curr_row = next_row

    return all_ti
###判读某行是否为大题行
def is_dati_row(dati_index,row):
    for index in dati_index:
        if index[0]==row:
            return True
    return False


###分析答案的结构
def AnalysAnswer(doc,start_row, end_row ):
    doc_elements = doc.elements
    dati_indexes=find_dati_row( doc, start_row, end_row)
    xiaoti_indexes=find_xiaoti_row( doc, start_row, end_row)
    if (len(dati_indexes)==0) or (len(xiaoti_indexes)==0):
        return ()
    all_ti = []
    ###分析参考答案的结构
    for i in range(0,len(xiaoti_indexes)):
        if i == len(xiaoti_indexes) - 1:
            b_row = xiaoti_indexes[i][0]
            e_row=end_row
        else:
            e_row=xiaoti_indexes[i+1][0]
            b_row=xiaoti_indexes[i][0]

        ti_num=re.findall( r'^(\d{1,2})[\s\.、．]', doc_elements[b_row]['text'].strip())[0]
        ti_index={'answer':[],'explain':[],'num':ti_num}
        curr_status='答案'     ###第一次，默认是答案

        for j in range(b_row+1, e_row ):
            tt=re.findall(r'【(.{2,5})】', doc_elements[j]['text'].strip())
            if dati_indexes:             ###如果大题存在，且该行是大题行，跳出
                if is_dati_row(dati_indexes,j):
                    break
            if tt:
                if tt[0].strip()=='解析':
                    curr_status='解析'
                    continue
                elif tt[0].strip()=='答案':
                    curr_status='答案'
                    continue
                else :   ####
                    curr_status='未识别'
                    continue

            if curr_status=='解析':
                ti_index['explain'].append(j)
            elif curr_status=='答案':
                ti_index['answer'].append(j)
        all_ti.append( ti_index.copy() )
    return all_ti

####删除括号及其里面的内容
def remove_brackets(sentence):
    result = re.findall(r'(^.*)[（\(].*[\)）](.*)', sentence)
    return ''.join(result)


'''
试卷的格式，我们认为只有2级，
1.大题（一、填空题）
2. 小题（19.）
'''

if __name__ == "__main__":
    path = 'data/2019年全国II卷文科综合高考真题.docx'
    doc = MyDocx.Document(path)
    all_ti_index = processPaper2(doc)
