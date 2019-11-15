import docx
from docx_utils.parse_paper import processPaper
import re
from lxml import etree
from dwml import omml
from docx_utils.namespaces import namespaces as docx_nsmap
import uuid
import os
from . import settings

'''
试卷的格式，我们认为只有2级，
1.大题（一、填空题）
2. 小题（19.）

'''
'''关于docx中的各种元素（数学公式、图片、表格、）
如果用代码生成docx: 1个run里面,可以包含多个文本(w:t)、图片（w:drawing）、公式等，
但是实际上，只要一编辑，1个run里面变只会有1个元素，然后在run里面设置这个元素的属性
数学公式可以和run同级，为oMathPara，可以为run中的元素，为oMath
图片必须为run中的元素（包裹在一个run里面），且该run只有这个图片元素，没有test等



'''
###尝试猜测题目类型
'''
题目类型：单选题，多选题，问答题，材料题（包含多个小题，他们可能是单选题、多选题、问答题），
'''


def guess_titype():
    pass


##获取材料题，到底有多少小问（题）
def get_question_quantity(paragraph, mode_text):
    text = paragraph.text
    r = re.findall(r'据此完成(\d{1,2})～(\d{1,2})题', text)
    start_ti_number, stop_ti_number = r[0]
    start_ti_number = int(start_ti_number)
    stop_ti_number = int(stop_ti_number)
    return stop_ti_number - start_ti_number + 1


####--------------------------------------
## 获取材料题的位置
def find_title_row(doc, b_row, curr_row, mode_text):
    if b_row > curr_row:
        print('开始行<结束行')
        return 'error'
    elif b_row == curr_row:
        return -1

    for i in range(b_row, curr_row):
        x = re.findall(mode_text, doc.paragraphs[i].text)
        if len(x) != 0:
            return i

    return -1

## 获取各个选项
def split_options(option_html):
    ops = []
    last_index = 0
    for i in range(ord('B') , ord('Z')):
        item = chr(i)
        index1 = option_html.find(item+'．')
        index2 = option_html.find(item + '.')
        index = index1 if index1>-1 else index2
        option=option_html[last_index:last_index+1]
        htmls = option_html[last_index+2:index]
        if len(option) == 0:
            return ops
        last_index = index
        ops.append({'label':option, 'content':htmls.strip() })

    return ops


### 获取标签
def get_tag(child):
    '''

    :param child: etree.Element,一般为一个run
    :return: 返回该run里面的内容的类型的字符串
    '''

    if child.tag.split('}')[-1] == 'pPr':
        return 'w:pPr'

    if child.tag.split('}')[-1] == 'oMath' or child.tag.split('}')[-1] == 'oMathPara':
        return 'm:oMath'

    wt = child.xpath('.//w:t', namespaces=child.nsmap)
    if len(wt) == 1:
        return 'w:t'

    wdrawing = child.xpath('.//w:drawing', namespaces=child.nsmap)
    if len(wdrawing) == 1:
        return 'w:drawing'

##处理w:t元素
def w_t2html(child):
    t = child.xpath('.//w:t/text()', namespaces=child.nsmap)[0]
    text = t.replace('<', '&lt;').replace('>', '&gt;')
    html = text
    return html

##处理w:drawing元素
'''
需要处理inline模式和float模式的2种图片，

'''
def w_drawing2html(doc, child):

    img_dir=settings.img_dir
    http_head=settings.http_head

    ##--检查img_dir参数检查-----------
    if img_dir == '':
        print('还未设置img_dir！')
        exit(0)
    else:
        if not os.path.exists(img_dir):
            os.mkdir(img_dir)
    #####检查http_head是否设置
    if http_head == '':
        print('还未设置http_head！')
        exit(0)
    if http_head[-1] != '/':
        http_head = http_head + '/'
    ##-----------------------------

    pics = child.xpath('.//w:drawing', namespaces=docx_nsmap)
    if len(pics) != 1:
        print("docx格式可能错误，w:drawing可能包含多张图片！")
        return 0


    pic = pics[0]
    mode=''
    if pic.xpath('.//wp:inline', namespaces=docx_nsmap):
        mode='inline'
    if pic.xpath('.//wp:anchor', namespaces=docx_nsmap):
        mode='anchor'

    one_mes = dict()
    size_ele = pic.xpath('.//wp:extent ', namespaces=docx_nsmap)[0]
    width = int(size_ele.attrib['cx']) / (360000 * 0.0264583)
    height = int(size_ele.attrib['cy']) / (360000 * 0.0264583)
    one_mes['width'] = width
    one_mes['height'] = height
    # element = pic.xpath('.//wp:'+mode, namespaces=pic.nsmap)[0]
    # a_graphic = pic.xpath('.//wp:'+mode+'//a:graphic', namespaces=docx_nsmap)
    blip = pic.xpath('.//a:blip ', namespaces=docx_nsmap)[0]
    blip_attr = blip.attrib
    for attr in blip_attr:
        if 'embed' in attr:
            one_mes['rId'] = blip_attr[attr]
            break
    pic_name = one_mes['rId']
    img = doc.part.rels[pic_name].target_ref
    ext = os.path.splitext(img)[-1]
    path = str(uuid.uuid1()).replace('-', '') + ext

    html = '<img src="' + http_head + path + '" width=' + "{:.4f}".format(one_mes["width"]) + \
           ' height=' + "{:.4f}".format(one_mes["height"]) + '>'
    img_part = doc.part.related_parts[pic_name]
    with open(os.path.join(img_dir, path), 'wb') as f:
        f.write(img_part._blob)

    return {'html':html, 'mode':mode}

##处理m:oMath元素
def o_math2html(child):
    tag = child.tag.split('}')[-1]

    if tag == 'oMath':
        tt = etree.Element('oMathPara')
        tt.append(child.__copy__())  ####使用copy， 不影响原来的结构
    else:
        tt = child

    mm = etree.tostring(tt).decode('utf-8')
    for math in omml.load_string(mm):
        text = math.latex
        break

    text = text.replace('<', '&lt;').replace('>', '&gt;')   ##处理大于号、小于号
    html = '\(' + text + '\)'
    return html

###处理title，得到title的htmls
def get_title_htmls(doc, titles_index):

    htmls = []
    images=[]
    for index in titles_index:
        result = paragraph2html(doc, index)
        htmls.append(result['htmls'].copy())
        images.append(result['images'])
    if len(images)>1:
        print('Warning! 该题目的标题部分包含', len(images), '张图片！')


    return ''.join(htmls)+'</br>'.join(images)


###检测1个run里面是否有多种内容，这种情况是无效的！！
def check_run(child):
    tag = get_tag(child)
    if tag == 'm:oMath' or tag == 'm:oMathPara' or tag=='w:tab':
        return  1

    i = 0
    # print('child=', child)
    run = child.__copy__()
    rPr = run.xpath('.//w:rPr', namespaces=run.nsmap)  ##删除run的属性

    if len(rPr) > 1:
        print(child , '184')
        print('docx格式出错了，len(w:rPr)!=1')
    elif len(rPr) == 1:
        run.remove(rPr[0])
    # else: 没有rPr,啥都不用做

    wt = run.xpath('./w:t', namespaces=run.nsmap)
    wdrawing = run.xpath('.//w:drawing', namespaces=run.nsmap)
    i = len(wt) + len(wdrawing)

    return i

def merge_wt(tree):  ###一个段落
    children = tree.getchildren()
    i = 0
    result = etree.Element('{' + docx_nsmap['w'] + '}p', nsmap=docx_nsmap)
    last_is_wt = False
    last_wt = ''
    while (i < len(children)):
        child = children[i]
        tag = get_tag(child)
        if tag=='w:pPr':
            i=i+1
            continue
        if tag == 'w:t':
            if last_is_wt:
                text = child.xpath('.//w:t/text()', namespaces=docx_nsmap)[0]
                last_wt_text=last_wt.xpath('.//w:t/text()', namespaces=docx_nsmap)[0]
                last_wt_text= last_wt_text + text
            else:
                last_wt = child
                last_is_wt = True

        else:
            if last_wt != '':
                result.append(last_wt)
                last_wt = ''
                last_is_wt = False
            result.append(children[i])
        i = i + 1

    return result

def paragraph2html(doc, index):
    tree = etree.fromstring(doc.paragraphs[index]._element.xml)

    children = tree.getchildren()
    htmls = []

    for child in children:
        tag = get_tag(child)

        if tag=='w:pPr':
            continue
        vv = check_run(child)

        if vv > 1:
            print('run中包含了多个类型 ')
            exit()
            continue
        elif vv == 0:
            # print('run里面没有找到合适元素！')
            # print('run=', child)
            continue
        html = ''

        if tag == 'w:t':  ##处理文本
            html = w_t2html(child)
            htmls.append(html)
        elif tag == 'w:drawing':  ##处理图片
            result = w_drawing2html(doc, child)
            if result['mode']=='inline':  ##不处理浮动图片，留到别处统一处理
                htmls.append(result['html'])

        elif tag == 'm:oMath' or tag == 'm:oMathPara':  ##处理数学公式
            html = o_math2html(child)
            htmls.append(html)
        elif tag == 'table':  ##处理表格
            pass

        # if html != ' ':   空格也应该原样输出，因为有时候存在特意的空格


    return ''.join(htmls)

def check_options(options):
    for i in range(1, len(options)):
        if ord(options[i]['label']) - ord(options[i - 1]['label']) != 1:
            print('获取options错误，请检查')
            return False
    return True

###获取选项的文本 + 特殊格式
###默认认为选项的字体等信息是不重要的！！！！
def options2html(doc, row):
    # result=[]
    text = ''
    images=''
    paragraph = doc.paragraphs[row]
    tree = etree.fromstring(paragraph._element.xml)
    children = tree.getchildren()
    for child in children:
        tag = get_tag(child)

        if tag=='w:pPr':
            continue
        vv = check_run(child)
        if vv > 1:
            print('run中包含了多个类型 ')
            print('run=', child)
            exit()
        elif vv == 0:
            print('options2html: run里面没有找到合适元素！')
            continue

        if tag == 'w:t':
            text = text + child.xpath('./w:t/text()', namespaces=docx_nsmap)[0]
        elif tag == 'w:drawing':
            result = w_drawing2html(doc, child)
            if result['mode']=='inline':  ##不处理浮动图片，留到后面一起处理
                text = text + result['html']
        elif tag == 'm:oMath':
            text = text + o_math2html(child)
        elif tag == '':
            pass

    return text

###处理options，得到options的html
def get_option_htmls(doc, options_indexes):

    option_html = ''

    for index in options_indexes:
        option_html += options2html(doc, index)   ###所有options组合起来的html
    options_htmls = split_options(option_html)   ###把每个optiion拆分出来

    if not check_options(options_htmls):
        print('选项识别错误')
        print('options=', options_htmls)

    return options_htmls

##for titel, 不包含选项的段落，可以直接转换
def paragraphs2htmls(doc, title_indexes):
    paragraphs = doc.paragraphs
    htmls = []

    for index in title_indexes:
        html = paragraph2html(doc, index)

        # reg = r'\d{1,2}[.．]\s{0,1}(\(|\（)\d{1,2}.{1}(\)|\）)'
        # text = re.sub(reg , '' , text)
        if re.match(r'^\（[\u4e00-\u9fa5]+\）' , html) is None:   ##不包含中文字符
            pass
        else:
            html = ''
        htmls.extend(html)

    return ''.join(htmls)
###单独处理浮动的图片
def get_float_image( doc, xiaoti_indexes, curr_xiaoti_index):
    paragraphs=doc.paragraphs
    htmls=[]
    indexes=[]
    indexes.extend(xiaoti_indexes[curr_xiaoti_index]['title'])
    if 'options' in xiaoti_indexes[curr_xiaoti_index]:
        indexes.extend(xiaoti_indexes[curr_xiaoti_index]['options'])

    for index in indexes:
        tree=etree.fromstring(doc.paragraphs[index]._element.xml)
        x=tree.xpath('.//w:drawing/wp:anchor', namespaces=docx_nsmap)
        if x:
            result=w_drawing2html(doc, x[0].getparent().getparent())
            if result['mode']=='anchor':
                htmls.append(result['html'])
    return ''.join(htmls)

def get_ti_content(doc, xiaoti_indexes, curr_xiaoti_index, curr_dati_row, mode_text):
    paragraphs = doc.paragraphs
    curr_row = xiaoti_indexes[curr_xiaoti_index]['title'][0]
    #####上一个题目的结尾的行号+1
    if curr_xiaoti_index == 0:
        last_row = curr_dati_row
    else:
        if 'options' in xiaoti_indexes[curr_xiaoti_index - 1]:
            # print('xiaoti_indexes[{0}]'.format(curr_xiaoti_index - 1), xiaoti_indexes[curr_xiaoti_index - 1])
            last_row = xiaoti_indexes[curr_xiaoti_index - 1]['options'][-1]
        else:
            last_row = xiaoti_indexes[curr_xiaoti_index - 1]['title'][-1]

    title_start_row = find_title_row(doc, last_row + 1, curr_row, mode_text)
    if title_start_row == -1:  ###不是大题包含小题模式（先有材料，然后跟几个题）
        ti = get_xiaoti_content(doc, xiaoti_indexes, curr_xiaoti_index)
        image_html=get_float_image(doc, xiaoti_indexes, curr_xiaoti_index)
        ti['stem']=ti['stem']+'</br>' + image_html
        return (curr_xiaoti_index + 1, {'title':'', 'questions':[ti] })

    ti = {}     ####开始处理大题包含小题的模式（材料题）
    lst = list(range(title_start_row, curr_row))
    ti['title'] = paragraphs2htmls(doc, lst)

    i = 0
    n = get_question_quantity(paragraphs[title_start_row], mode_text)
    questions = []
    while (i < n):
        question = get_xiaoti_content(doc, xiaoti_indexes, curr_xiaoti_index + i)
        image_html = get_float_image(doc, xiaoti_indexes, curr_xiaoti_index+i)
        question['stem'] = question['stem'] + '</br>' + image_html
        questions.append(question)
        i = i + 1
    ti['questions'] = questions

    return (curr_xiaoti_index + n, ti)


##处理1个小题
def get_xiaoti_content(doc, xiaoti_indexes, curr_index):
    q = {}
    title_indexes = xiaoti_indexes[curr_index]['title']

    xx = paragraphs2htmls(doc, title_indexes)
    q['stem']=re.sub(r'\d{1,2}[.．]\s{0,}', '', xx)   ###去除题号

    q['index'] = re.findall(r'^(\d{1,2})[.．、]\s{0,}', xx)[0]   ###获取题号

    if 'options' in xiaoti_indexes[curr_index]:
        option_indexes = xiaoti_indexes[curr_index]['options']
        q['options'] = get_option_htmls(doc, option_indexes)
        if '一项' in doc.paragraphs[title_indexes[0]].text:
            q['type']='single'

    return q

if __name__ == "__main__":
    path = '../src/2019年全国II卷文科综合高考真题.docx'
    doc = docx.Document(path)
    all_ti_index = processPaper(doc)
    paragraphs = doc.paragraphs

    i = 0
    mode_text = r'完成\d～\d题'  ##模式字符串

    tis = []
    # while(i<len(all_ti_index)):
    for dati in all_ti_index:

        curr_dati_row = dati[0]
        # guess_titype(paragraphs[curr_dati_row].text)

        xiaoti_indexes = dati[1]
        curr_row = xiaoti_indexes[0]
        curr_index = 0
        while (curr_index < len(xiaoti_indexes)):
            curr_index, ti = get_ti_content(doc, xiaoti_indexes, curr_index, curr_dati_row, mode_text)
            tis.append(ti)
