import docx
from docx_utils.parse_paper import  processPaper
from docx_utils.ti2html import  get_ti_content
import  re
import json
def get_answer(doc , all_indexes):
    all_ans = []
    for dati in all_indexes:
        curr_dati_row = dati[0]
        xiaoti_indexes = dati[1]
        curr_index = 0
        while curr_index<len(xiaoti_indexes):
            curr_index , ans = get_ti_content(doc , xiaoti_indexes , curr_index ,curr_dati_row, '')
            content = re.sub(r'\d{1,2}[.．]\s{0,}', '', ans['title'])
            all_ans.append(content)

    return all_ans

def merge_answer(subjects , ans_list):
    for i in range(0 , len(subjects)-1):
        sub = subjects[i]
        sub['reference'] = ans_list[i].strip()
        if len(sub['questions']) == 1:
            sub['questions'][0]['solution'] = ans_list[i].strip()

    return subjects

if __name__ == "__main__":

    path = 'src/2019年全国II卷文科综合高考真题.docx'
    doc = docx.Document(path)
    all_ti_index = processPaper(doc)
    paragraphs = doc.paragraphs

###处理试卷
    i = 0
    mode_text = r'完成\d{1,2}～\d{1,2}题'  ##模式字符串
    tis = []
    # while(i<len(all_ti_index)):
    for dati in all_ti_index:
        curr_dati_row = dati[0]

        xiaoti_indexes = dati[1]
        curr_row = xiaoti_indexes[0]
        curr_index = 0
        while (curr_index < len(xiaoti_indexes)):
            curr_index, ti = get_ti_content(doc, xiaoti_indexes, curr_index, curr_dati_row, mode_text)
            tis.append(ti)

####处理答案
    path = 'src/2019年全国II卷文科综合高考真题-答案.docx'
    doc = docx.Document(path)
    all_answer_index = processPaper(doc)
    answers=get_answer(doc,all_answer_index )



    # with  open('math_data.json', 'w', encoding='utf-8') as fp:
    #     json.dump(tis, fp, ensure_ascii=False,indent = 4, separators=(',', ': '))
