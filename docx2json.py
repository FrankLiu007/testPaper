import docx
from docx_utils.parse_paper import  processPaper, processPaper2
from docx_utils.ti2html import  get_ti_content
import  re
import json
import sys
import os
from docx_utils import  settings
'''
###获取答案
#答案和试卷不一样，答案一定是先有题号，然后跟着答案。
# 试卷则不同，试卷的选择题里面有材料题（看一段材料，做几个选择题，请参照文综试卷的选择题部分）
'''
def get_answer(doc , all_indexes):
    all_answer = {}
    for dati in all_indexes:
        curr_dati_row = dati[0]
        xiaoti_indexes = dati[1]
        curr_index = 0
        while curr_index<len(xiaoti_indexes):
            curr_index , ans = get_ti_content(doc , xiaoti_indexes , curr_index ,curr_dati_row, '')
            if len(ans['questions'])!=1:
                print('答案格式错误！')
            content = re.sub(r'\d{1,2}[.．]\s{0,}', '', ans['questions'][0]['stem'])

            all_answer[ans['questions'][0]['number']]=content
    return all_answer

def merge_answer(tis , answer_list):
    for ti in tis:
        reference=''
        for q in ti['questions']:

            q['solution']=answer_list[q['number']]
            reference = reference + q['number'] + '. ' + q['solution']
            if 'options' in q:
                q['solution']=q['solution'].replace('<p>', '').replace('</p>', '').strip()  ####选择题的答案不能是html
            else:
                q.pop('solution')

        ti['reference']=reference
    return 0

#####给题目增加分数和题型
def add_score_and_titype(ti, text):
    score = re.findall(r'每[小]{0,1}题(\d{1,2})分', text)     ##，每个小题的分数

    ###题型 titpye
    xx=re.findall(r'(.{1,8}题)', text)[0]
    b=text.find('、')
    e=text.find('题')
    if b!=-1 and e!=-1:
        type_str=xx[b+1:e+1]
        if b > 3:
            print('题型识别可能出错：', xx[b+1:e+1])
    ti['category'] = type_str

    q_tpye='GENERAL'
    if '只有一项' in text  or '的一项是' in text :
        q_tpye='SINGLE'

    if score:
        ti['total'] = int(score[0]) * len(ti['questions'])
    ss=0
    for q in ti['questions']:
        if (not 'type' in q) or  (q_tpye==''):
            q['type']=q_tpye
        if score:
            q['score'] = int(score[0])
        else:
            s = re.findall(r'["(","（"](\d{1,2})分[")","）"]', q['stem'])
            if s:
                q['score'] = int(s[0])
            else:
                q['score']=0
            ss = ss + q['score']
    if not score:
        ti['total'] = ss
    return 0

##-----------------------------
def get_tis(doc, all_ti_index):
    mode_text = r'完成\d{1,2}～\d{1,2}题'  ##模式字符串
    paragraphs=doc.paragraphs
    tis = []
    # while(i<len(all_ti_index)):
    for dati in all_ti_index:
        curr_dati_row = dati[0]

        xiaoti_indexes = dati[1]          ####题目集合index
        curr_row = xiaoti_indexes[0]      # 题型段落index
        curr_index = 0
        while (curr_index < len(xiaoti_indexes)):

            curr_index, ti = get_ti_content(doc, xiaoti_indexes, curr_index, curr_dati_row, mode_text)
            add_score_and_titype(ti, paragraphs[curr_dati_row].text)
            tis.append(ti)
    return  tis
#-----------------------------------------
def docx_paper2json(data_dir, paper_path, answer_path):

    doc = docx.Document(os.path.join(data_dir, paper_path))
    all_ti_index = processPaper2(doc)
    paragraphs = doc.paragraphs
###处理试卷
    i = 0
    tis=get_tis(doc, all_ti_index)
####处理答案
    doc = docx.Document(os.path.join(data_dir, answer_path))
    all_answer_index = processPaper2(doc)
    answers=get_answer(doc,all_answer_index )
    merge_answer(tis, answers)

    return tis

if __name__ == "__main__":
###run 本脚本的例子：
## python docx2json.py  src  文综  2019年全国II卷文科综合高考真题.docx 2019年全国II卷文科综合高考真题-答案.docx img https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/ 文综.json
# python docx2json.py  src  数学 2019年全国I卷理科数学高考真题.docx 2019年全国I卷理科数学高考真题答案.docx img https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/ 数学.json
    settings.init()

    if len(sys.argv)==8:
        data_dir=sys.argv[1]
        subject=sys.argv[2]
        paper_path=sys.argv[3]
        answer_path=sys.argv[4]
        img_dir=sys.argv[5]
        http_head=sys.argv[6]
        out_path=sys.argv[7]

        settings.img_dir=os.path.join(data_dir, img_dir)
        settings.http_head=http_head
    else:    ###跑例子用的默认参数,保证在ipython下面也可以直接跑
        print('参数错误，正确用法： process_3x_paper.py 真题.docx 答案.docx')
        data_dir='src'
        subject='文综'
        paper_path='化学试卷.docx'
        answer_path='化学答案.docx'
        img_dir='img'
        http_head=' https://ehomework.oss-cn-hangzhou.aliyuncs.com/item/'
        out_path='化学.json'
        settings.img_dir=os.path.join(data_dir, img_dir)
        settings.http_head=http_head


    if subject =='语文':
        # in ['数学','物理','化学', '历史', '地理','生物']:
        pass
    elif subject=='英语':
        pass
    else :

        tis=docx_paper2json(data_dir, paper_path, answer_path )

    with  open(os.path.join(data_dir, out_path), 'w', encoding='utf-8') as fp:
        json.dump(tis, fp, ensure_ascii=False,indent = 4, separators=(',', ': '))
