import docx
from docx_utils.parse_paper import  processPaper2
from docx_utils.ti2html import get_ti_content,  get_element_text
import re
import json
import sys
import os
from docx_utils import settings
from docx_utils import namespaces as docx_nsmap
from lxml import etree
'''
###获取答案
#答案和试卷不一样，答案一定是先有题号，然后跟着答案。
# 试卷则不同，试卷的选择题里面有材料题（看一段材料，做几个选择题，请参照文综试卷的选择题部分）
'''

def get_answer(doc, all_indexes):
    all_answer = {}
    for dati in all_indexes:
        curr_dati_row = dati[0]
        xiaoti_indexes = dati[1]
        curr_index = 0
        while curr_index < len(xiaoti_indexes):
            curr_index, ans = get_ti_content(doc, xiaoti_indexes, curr_index, curr_dati_row, '')
            if len(ans['questions']) != 1:
                print('答案格式错误！')
            content = re.sub(r'^<p>\d{1,2}[.．]\s{0,}', '', ans['questions'][0]['stem'])

            all_answer[ans['questions'][0]['number']] = content
    return all_answer


def merge_answer(tis, answer_list):
    for ti in tis:
        reference = ''
        for q in ti['questions']:

            q['solution'] = answer_list[q['number']]
            reference = reference + q['number'] + '. ' + q['solution']
            if 'options' in q:
                q['solution'] = q['solution'].replace('<p>', '').replace('</p>', '').strip()  ####选择题的答案不能是html
            else:
                q.pop('solution')

        ti['reference'] = reference
    return 0


#####给题目增加分数和题型
def add_score_and_titype(ti, text):
    score = re.findall(r'每[小]{0,1}题(\d{1,2})分', text)  ##，每个小题的分数

    ###题型 titpye
    xx = re.findall(r'(.{1,8}题)', text)[0]
    b = text.find('、')
    e = text.find('题')
    if b != -1 and e != -1:
        type_str = xx[b + 1:e + 1]
        if b > 3:
            print('题型识别可能出错：', xx[b + 1:e + 1])
    ti['category'] = type_str

    q_tpye = 'GENERAL'
    if '只有一项' in text or '的一项是' in text:
        q_tpye = 'SINGLE'

    if score:
        ti['total'] = int(score[0]) * len(ti['questions'])
    ss = 0
    for q in ti['questions']:
        if (not 'type' in q) or (q_tpye == ''):
            q['type'] = q_tpye
        if score:
            q['score'] = int(score[0])
        else:
            s = re.findall(r'["(","（"](\d{1,2})分[")","）"]', q['stem'])
            if s:
                q['score'] = int(s[0])
            else:
                q['score'] = 0
            ss = ss + q['score']
    if not score:
        ti['total'] = ss
    return 0

def get_tis(doc, all_ti_index):
    mode_text = r'完成\d{1,2}～\d{1,2}题'  ##模式字符串

    tis = []
    # while(i<len(all_ti_index)):
    for dati in all_ti_index:
        curr_dati_row = dati[0]

        xiaoti_indexes = dati[1]  ####题目集合index
        curr_row = xiaoti_indexes[0]  # 题型段落index
        curr_index = 0
        while (curr_index < len(xiaoti_indexes)):
            curr_index, ti = get_ti_content(doc, xiaoti_indexes, curr_index, curr_dati_row, mode_text)
            add_score_and_titype(ti, get_element_text( doc,curr_dati_row) )
            tis.append(ti)
    return tis


# -----------------------------------------
def docx_paper2json(pars):
    data_dir = pars['working_dir']
    paper_path = pars['question_docx']
    doc = docx.Document(os.path.join(data_dir, paper_path))
    all_ti_index = processPaper2(doc)

    ###处理试卷
    i = 0
    print('开始处理试卷...')
    tis = get_tis(doc, all_ti_index)
    ####处理答案
    if 'answer_docx' in pars:
        answer_path = pars['answer_docx']
        print('开始处理答案...')
        answer_fullpath=os.path.join(data_dir, answer_path)
        if not os.path.exists(answer_fullpath):
            print('答案文件不存在！忽略答案文件！')
            return tis
        doc = docx.Document(answer_fullpath)
        all_answer_index = processPaper2(doc)
        if not all_answer_index:
            print('all_answer_index = processPaper2(doc)')
            print('分析答案版面出错！')
            return tis
        answers = get_answer(doc, all_answer_index)
        if not answers:
            print('answers = get_answer(doc, all_answer_index)')
            print('获取答案内容出错！')
            return tis
        print('开始 合并试题和答案...')
        merge_answer(tis, answers)

    return tis


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
        elif argv[i] == '-answer_docx':
            pars['answer_docx'] = argv[i + 1]
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
    print('-answer_docx  答案docx文件')
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


if __name__ == "__main__":
    ###run 本脚本的例子：
    ## python docx2json.py  -working_dir data  -subject 文综  -question_docx 2019年全国II卷文科综合高考真题.docx  -answer_docx 2019年全国II卷文科综合高考真题-答案.docx -img_dir img -http_head https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/  -out_json 文综.json
    # python docx2json.py  -working_dir data  -subject 数学  -question_docx  2019年全国I卷理科数学高考真题.docx  -answer_docx 2019年全国I卷理科数学高考真题答案.docx -img_dir img -http_head https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/  -out_json 文综.json

    settings.init()
    pars = {}

    if len(sys.argv)<5:  ###跑例子用的默认参数,保证在ipython下面也可以直接跑
        print('参数错误，正确用法： docx2json.py 真题.docx 答案.docx')
        pars['working_dir'] = 'data'
        pars['subject'] = '化学'
        pars['question_docx'] = '2019年咸宁高中高一年级10月25日化学周练.docx'
        # pars['answer_docx'] = '2019年全国I卷理科数学高考真题答案.docx'
        pars['img_dir'] = 'img'
        pars['http_head'] = ' https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/'
        pars['out_json'] = '化学.json'

    else:
        pars = parse_commandline(sys.argv)

    print('开始检查输入参数...')
    check_pars(pars)



    print('开始检查运行环境...')
    check_env()

    settings.img_dir = os.path.join(pars['working_dir'], pars['img_dir'])
    settings.http_head = pars['http_head']
    subject = pars['subject']

    if subject == '语文':
        # in ['数学','物理','化学', '历史', '地理','生物']:
        pass
    elif subject == '英语':
        pass
    else:

        tis = docx_paper2json(pars)

    with  open(os.path.join(pars['working_dir'], pars['out_json']), 'w', encoding='utf-8') as fp:
        json.dump(tis, fp, ensure_ascii=False, indent=4, separators=(',', ': '))

