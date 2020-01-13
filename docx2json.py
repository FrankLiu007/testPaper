from docx_utils.parse_paper import  parse_paper, parse_answer
from docx_utils.ti2html import get_ti_content, paragraphs2htmls
import re
import json
import sys
import os
from docx_utils import settings
from docx_utils import namespaces as docx_nsmap
from lxml import etree
import docx_utils.MyDocx as MyDocx
'''
###获取答案
#答案和试卷不一样，答案一定是先有题号，然后跟着答案。
# 试卷则不同，试卷的选择题里面有材料题（看一段材料，做几个选择题，请参照文综试卷的选择题部分）
'''

def merge_answer(tis, answers):

    for ti in tis:
        reference = ''
        for q in ti['questions']:
            if q['objective']:
                tt=answers[q['number']]['answer']
                # tt=tt.replace('<p>','').replace('</p>','').strip().split()
                tt = re.findall(r'<p[\s]{0,}.*?>(.*?)</p>', tt)

                q['solution'] = ''.join(tt).strip()
                if len(q['solution'])==1:
                    q['type']='SINGLE'
                else:
                    q['type'] = 'MULTIPLE'
            # else:
            #     q['solution'] = answers[q['number']]['answer']
            q.pop('objective')
            reference = reference + q['number'] + '. 【答案】' + answers[q['number']]['answer']+"<p>【解析】</p>"+answers[q['number']]['explain']
        ti['reference'] = reference
    return 0

###---------------------------------------------------
def get_tis(doc, all_ti_index):
    all_tis=[]
    for dati in all_ti_index:
        for ti_index in dati:
            ti=get_ti_content(doc, ti_index)
            questions=ti['questions']
            ti['category']=ti_index['category']
            # ti['total'] = ti_index['total']
            for i in range(0, len(questions)):
                questions[i]['objective']=ti_index['questions'][i]['objective']  ##暂时objective
                questions[i]['number']=ti_index['questions'][i]['number']  ##加入题号
                questions[i]['type'] = ti_index['questions'][i]['type']  ##加入题号
                questions[i]['score'] = ti_index['questions'][i]['score']  ##加入分数
            all_tis.append(ti)
    return all_tis

###-----------------------
def get_answer_start_row(doc):
    row=-1
    for i in range(0,len(doc.elements)) :
        if '参考答案' in doc.elements[i]['text']:
            row= i+1
            break
    return row
##------------------利用答案确定题目的单选、多选属性-----------------------------
def check_ti_type(tis):

    for ti in tis:
        for q in ti['questions']:
            pass

# -----------------------------------------
def docx_paper2json(pars):
    # mode_text = r'(\d{1,2})[～\-~](\d{1,2})[小]{0,1}题'
    mode_text=settings.mode_text
    data_dir = pars['working_dir']
    paper_path = pars['question_docx']
    doc = MyDocx.Document(os.path.join(data_dir, paper_path))
    answer_start_row=get_answer_start_row(doc)
    if answer_start_row==-1:
        end_row=len(doc.elements)-1
    else:
        end_row=answer_start_row-2
    all_ti_index = parse_paper(doc, 0, end_row, mode_text)

    ###处理试卷
    print('开始处理试卷...')
    tis = get_tis(doc, all_ti_index)
    ####处理答案
    if answer_start_row!=-1:
        print('开始处理答案...')
        answer_indexes=parse_answer(doc,answer_start_row,len(doc.elements)-1)
        if not answer_indexes:
            print('answers = get_answer(doc, all_answer_index)')
            print('获取答案内容出错！')
            return tis
        answers={}
        for answer_index in answer_indexes:
            answer={'answer':'','explain':'','num':answer_index['num']}
            if answer_index['answer']:
                html=paragraphs2htmls(doc,answer_index['answer'], wt_style_ignore=True)  ##答案的文字部分不需要样式
                answer['answer']=clean_answer(html)
            if answer_index['explain']:
                html=paragraphs2htmls(doc,answer_index['explain'], wt_style_ignore=True) ##答案的文字部分不需要样式
                answer['explain']=clean_answer(html)
            answers[answer['num']]=answer.copy()
        print('开始 合并试题和答案...')
        merge_answer(tis, answers)
        check_result(tis)

    return tis
###删除html两端的空格(&nbsp;)
def remove_nbsp(html):
    html=html.strip()
    e=len(html)
    while True:
        l=html.rfind('&nbsp;',0, e)
        if l==0:
            html=html[6:]
        elif l==-1:
            return html
        elif l==len(html)-6:
            html=html[:l]
        else:
            e=l

def clean_answer(html):
    paragraphs=re.findall(r'(<p[\s]{0,}.*?>.*?</p>)', html)
    result=[]
    for paragraph in paragraphs:
        b=paragraph.find('>')
        e=paragraph.find('<', b)
        txt=paragraph[b+1:e]
        txt=re.sub(r'\d{1,2}[\s\.、．]','',txt).replace('【答案】', '').replace('【解析】', '')
        txt=remove_nbsp(txt)
        if txt=='':
            continue
        result.append(paragraph[:b+1]+txt+paragraph[e:])
    return ''.join(result)

    # txt.re
    # xx=re.match(r'<p[\s]{0,}.*?>.*?</p>', html)   ###第一个段落（p标签）
    #
    # text=html[xx.start():xx.end()]
    # text=remove_nbsp(text)
    #
    # b = html.find('>')
    # e = html.find('<', b)
    # tt=html[b:e]
    # num=re.findall(r'(\d{1,2}[\s\.、．])', tt)
    # if num :
    #     tt=tt.replace(num[0],'')
    # tt = tt.replace('【答案】', '').replace('【解析】', '')
    # return html[:b]+tt+html[e+1:]



####重新检查结果
def check_result(tis):
    check_objective(tis)
    pass
###解决没有选项的选择题
def check_objective(tis):

    pass

###处理命令行参数
def parse_commandline(argv):
    i = 1
    pars = {}
    while (i < len(argv)):
        if argv[i] == '-working_dir':
            pars['working_dir'] = argv[i + 1]
            i = i + 1
        elif argv[i] == '-help':
            print_ussage()
            exit(0)
        elif argv[i] == '-subject':
            pars['subject'] = argv[i + 1]
            i = i + 1
        elif argv[i] == '-question_docx':
            pars['question_docx'] = argv[i + 1]
            i = i + 1

        elif argv[i] == '-img_dir':
            pars['img_dir'] = argv[i + 1]
            i = i + 1
        elif argv[i] == '-http_head':
            pars['http_head'] = argv[i + 1]
            i = i + 1
        elif argv[i] == '-out_json':
            pars['out_json'] = argv[i + 1]
            i = i + 1

        i = i + 1
    return pars

def print_ussage():
    print('docx2json使用说明')
    print('python docx2json options')
    print('-help    打印脚本使用说明')
    print('-working_dir  dir  设置工作目录')
    print('-img_dir  dir  设置图片目录')
    print('-subject  数学  设置学科')
    print('-question_docx  试卷docx文件')
    print('-http_head  http_head  设置http头')


def check_pars(pars):
    img_dir = pars['img_dir']
    working_dir = pars['working_dir']
    http_head = pars['http_head']
    ##--检查img_dir参数检查-----------
    if img_dir == '':
        print('还未设置img_dir！')
        exit(0)
    else:
        if not os.path.exists(os.path.join(working_dir, img_dir)):
            print('img_dir:'+img_dir+' 不存在，开始创建...')
            os.makedirs(os.path.join(working_dir, img_dir))
    #####检查http_head是否设置
    if pars['http_head'] == '':
        print('还未设置http_head！')
        exit(0)
    else:
        if http_head[-1] != '/':
            pars['http_head'] = http_head + '/'
    ###检查临时文件目录
    tmp_dir=os.path.join(pars['working_dir'], 'tmp')
    if not os.path.exists(tmp_dir):
        print('临时文件夹 '+tmp_dir+' 不存在, 开始创建..')
        os.makedirs(tmp_dir)
    settings.tmp_dir=tmp_dir
##检测环境
def check_env():

    if os.name == 'nt':
        print('警告！ 尽量不要在windows下，使用本脚本，可能会出现一些错误！')
    try:
        pass
    except:
        print('ruby 或者 mathtype_to_mathml 未正确安装！')
        exit(0)
    try:
        pass
    except:
        print('未找到wmf2svg模块')
        exit(0)

def save_html(tis, fname):
    zz='''
<!DOCTYPE html>
<html>
<head>
  <script type="text/javascript"
     src="http://cdn.mathjax.org/mathjax/latest/MathJax.js?config=TeX-AMS-MML_HTMLorMML">
  </script>
</head>
<body>
'''
    outf=open(fname, 'w')

    print(zz, file=outf)
    for ti in tis:
        print(ti['category'], file=outf)
        print(ti['title'] , file=outf)
        for q in ti['questions']:
            tt=q['number'] +'、(' + str(q['score'])+'分)  题型：'+ q['type'] + q['stem']
            if 'options' in q:
                for option in q['options']:
                    tt=tt+option['label']+option['content'] +'<br/>'

            print(tt +'<br/>', file=outf)
        print(ti['reference'] , file=outf)
    zz='''
</body>
</html>'''
    print(zz, file=outf)
    outf.close()

if __name__ == "__main__":
    ###run 本脚本的例子：
    ## python docx2json.py  -working_dir data  -subject 文综  -question_docx 2019年全国II卷文科综合高考真题.docx  -answer_docx 2019年全国II卷文科综合高考真题-答案.docx -img_dir img -http_head https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/  -out_json 文综.json
    # python docx2json.py  -working_dir data  -subject 数学  -question_docx  2019年全国I卷理科数学高考真题.docx  -answer_docx 2019年全国I卷理科数学高考真题答案.docx -img_dir img -http_head https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/  -out_json 文综.json

    settings.init()
    pars = {}

    if len(sys.argv)<5:  ###跑例子用的默认参数,保证在ipython下面也可以直接跑
        print('参数错误，正确用法： docx2json.py 真题.docx 答案.docx')
        pars['working_dir'] = 'data'
        pars['subject'] = '语文'
        pars['question_docx'] = '标准测试-语文周练（6月2日）.docx'
        # pars['answer_docx'] = '2019年全国I卷理科数学高考真题答案.docx'
        # 标准测试-数学理试题12-22.docx  标准测试-高二地理周考卷9.docx  标准测试-语文周练（6月2日）.docx

        pars['img_dir'] = 'img'
        pars['http_head'] = ' https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/'
        pars['out_json'] = '地理.json'

    else:
        pars = parse_commandline(sys.argv)

    print('开始检查输入参数...')
    check_pars(pars)

    print('开始检查运行环境...')
    check_env()

    settings.img_dir = os.path.join(pars['working_dir'], pars['img_dir'])
    settings.http_head = pars['http_head']
    settings.mathtype_convert_to="mathml"  ###mathml or png
    settings.subject = pars['subject']

    tis = docx_paper2json(pars)

    with  open(os.path.join(pars['working_dir'], pars['out_json']), 'w', encoding='utf-8') as fp:
        json.dump(tis, fp, ensure_ascii=False, indent=4, separators=(',', ': '))
