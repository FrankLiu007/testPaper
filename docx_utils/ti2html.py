import docx
from docx_utils.parse_paper import processPaper
import re
from lxml import etree
from dwml import omml
from docx_utils.namespaces import namespaces as docx_nsmap
import uuid
import os
import subprocess
from . import settings
from PIL import Image
import io
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
        if index1==-1 and index2==-1:
            break
        index = index1 if index1>-1 else index2
        option=option_html[last_index:last_index+1]
        htmls = option_html[last_index+2:index]
        if len(option) == 0:
            return ops
        last_index = index
        ops.append({'label':option, 'content':htmls.strip() })
    ops.append({'label':chr(i-1), 'content':option_html[last_index+2:] })
    return ops


### 获取标签
def get_tag(child):
    '''

    :param child: etree.Element,一般为一个run
    :return: 返回该run里面的内容的类型的字符串
    '''
    w_pPr='{'+docx_nsmap['w']+'}pPr'
    m_oMath='{'+docx_nsmap['m']+'}oMath'
    m_oMathPara='{'+docx_nsmap['m']+'}oMathPara'

    if w_pPr==child.tag:
        return 'w:pPr'

    if child.tag== m_oMath or child.tag == m_oMathPara:
        return 'm:oMath'
    #w:t
    wt = child.xpath('.//w:t', namespaces=child.nsmap)
    if len(wt) == 1:
        return 'w:t'
    #w:drawing
    wdrawing = child.xpath('.//w:drawing', namespaces=child.nsmap)
    if len(wdrawing) == 1:
        return 'w:drawing'
    #w:pict
    wpict = child.xpath('.//w:pict', namespaces=child.nsmap)
    if len(wpict) == 1:
        return 'w:pict'
    wobject = child.xpath('.//w:object', namespaces=child.nsmap)
    if len(wobject) == 1:
        return 'w:object'

##处理w:t元素
def w_t2html(child):
    t = child.xpath('.//w:t/text()', namespaces=child.nsmap)[0]
    t=t.replace('<', '&lt;').replace('>', '&gt;')

    kk = '{' + docx_nsmap['w'] + '}val'
    ##1.处理上下标
    ee=child.xpath('./w:rPr/w:vertAlign', namespaces=docx_nsmap)
    if ee:
        val = ee[0].attrib[kk]
        if val=="subscript":
            t='<sub>'+t+'</sub>'
        elif val=='superscript':
            t = '<sup>' + t + '</sup>'

    ##2.处理下划线
    ee = child.xpath('./w:rPr/w:u', namespaces=docx_nsmap)
    if ee:
        # if ee[0] == "": ##下划线
        #     t = '<>' + t + '</u>'
        # elif ee['']:  ##双下划线
        t = '<u>' + t + '</u>'
    ##3.处理加粗
    ee = child.xpath('./w:rPr/w:b', namespaces=docx_nsmap)
    if ee:
        t = '<b>' + t + '</b>'
    ##4.处理倾斜
    ee = child.xpath('./w:rPr/w:i', namespaces=docx_nsmap)
    if ee:
        t = '<i>' + t + '</i>'

    return t

###wmf2svg -----convert by wmf2svg-0.9.8.jar--------------
## wmf2gd -t png -o abc.png test.wmf   --maxwidth=130 --maxpect
##测试过 imagemagick的convert、wmf2gd和pillow的save来处理wmf，效果不好
def wmf2svg(blob, svg_path):
    fname=uuid.uuid1().hex+'.wmf'
    with open(fname, 'wb') as f:
        f.write(blob)
    cmd='java -jar docx_utils/wmf2svg-0.9.8.jar  '+fname+ '  ' + svg_path
    status, output = subprocess.getstatusoutput(cmd)
    if status:
        print('Warning!\r\n ' + cmd + '\r\n wmf转png失败！')
        return 0
    return 1
##mathtype2mml---
def math_type2mml(ole_blob, mml):
    '''
    mathtype嵌入的数学公式,转成mathml是ruby的mathtype_to_mathml做的，(https://github.com/jure/mathtype_to_mathml)
    1. intall mathtype_to_mathml:
    gem install mathtype_to_mathml pry
    2. usage:
    require "mathtype_to_mathml"
    MathTypeToMathML::Converter.new('oleObject1.bin').convert
    exit
    '''
    fname=uuid.uuid1()+'.bin'
    with open(fname, 'wb') as f:
        f.write(ole_blob)
    output = subprocess.check_output(['ruby', '-w', 'docx_utils/mathtype_ole2mathml.rb', fname])
    mathml = output.decode('utf-8').replace('<?xml version="1.0"?>', '').replace('block','inline')  ###暂时用inline，一般试卷中的公式都是inline
    return mathml

###处理w:object
def w_object2html(doc, child):
    img_dir = settings.img_dir
    http_head = settings.http_head

    ole={}
    ole['styles']={}

    styles=child.xpath('.//v:shape', namespaces=docx_nsmap)[0].attrib['style']
    for style in styles.split(';'):
        a,b=style.split(':')
        ole['styles'][a]=b
    ole['width']=ole['styles']['width'].replace('pt', '')
    ole['height'] = ole['styles']['height'].replace('pt', '')

    ole_object = child.xpath('.//o:OLEObject', namespaces=docx_nsmap)[0]

    if ole_object.attrib['ProgID']=="Equation.DSMT4":  #是mathtype嵌入的数学公式,转成html
        rId = ole_object.attrib['{' + docx_nsmap['r'] + '}id']
        ole['object_path'] = doc.part.rels[ole_object['rId']].target_ref
        ole['ole_part'] = doc.part.rels[ole_object['rId']].target_part
        mml=math_type2mml(ole['ole_part'].blob)
        return {'html':mml}

    else:    ##read v:imagedata only, usually .wmf file， 暂时转换成svg
        ole['rId'] = child.xpath('.//v:imagedata', namespaces=docx_nsmap).attrib[docx_nsmap['r'] + 'id']
        ole['img_path'] = doc.part.rels[ole['rId']].target_ref
        ole['img_part'] = doc.part.rels[ole['rId']].target_part
        ext = os.path.splitext(ole['path'])[-1]
        if ext=='.wmf':
            out_img_path = str(uuid.uuid1()).replace('-', '') + '.svg'
            wmf2svg(ole['img_part'].blob , out_img_path)
            html = '<img src="' + http_head + out_img_path + '" width=' + "{:.4f}".format(ole["width"]) + \
                   ' height=' + "{:.4f}".format(ole["height"]) + '>'

            return {'html': html, 'mode': 'inline'}

##------另外一种格式的图片，只有inline 模式-------------
def w_pict2html(doc, child):
    print('call w_pict2html')
    img_dir=settings.img_dir
    http_head=settings.http_head

    pics = child.xpath('.//w:pict', namespaces=docx_nsmap)
    if len(pics) != 1:
        print("docx格式可能错误，w:drawing可能包含多张图片！")
        return 0
    pic = pics[0]

    fig = dict()
    fig['styles']={}
    styles=pic.xpath('.//v:shape', namespaces=docx_nsmap)[0].attrib['style']
    for style in styles.split(';'):
        a,b=style.split(':')
        fig['styles'][a]=b
    fig['rId']=pic.xpath('.//v:imagedata', namespaces=docx_nsmap)[0].attrib['{'+docx_nsmap['r']+'}id']
    fig['path'] = doc.part.rels[fig['rId']].target_ref
    fig['img_part']=doc.part.rels[fig['rId']].target_part
    fig['width']=fig['styles']['width'].replace('pt', '')
    fig['height'] = fig['styles']['height'].replace('pt', '')

    ext = os.path.splitext(fig['path'] )[-1]
    fname = uuid.uuid1().hex   ###convert all image to .png
    out_img_path=''
    if ext=='.wmf':  ##扩展名为.wmf
        ## wmf2gd, imagemagick的convert效果都不太好，暂时不用
        ##暂时不用imagemagick的convert
        with open( 'out.wmf', 'wb') as f :
            f.write(fig['img_part'].blob)
        out_img_path=fname+'.svg'
        status, output=subprocess.getstatusoutput('java -jar docx_utils/wmf2svg-0.9.8.jar  out.wmf ' + os.path.join(img_dir, out_img_path)  )
        if  status:
            print('Warning! '+fig['path'] +'转png失败！')
    else:
        image=Image.open(io.BytesIO(fig['img_part'].blob))
        out_img_path=fname+'.png'
        image.save(os.path.join(img_dir, out_img_path))          ###convert all image to .png by PIL

    html = '<img src="' + http_head + out_img_path + '" width=' + fig["width"] + \
           ' height=' + fig["height"] + '>'


    return {'html':html, 'mode':'inline'}

##处理w:drawing元素
'''
需要处理inline模式和float模式的2种图片，

'''
def w_drawing2html(doc, child):
    img_dir=settings.img_dir
    http_head=settings.http_head

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

    ####直接取出rId
    rId=pic.xpath('.//a:blip ', namespaces=docx_nsmap)[0].attrib['{'+docx_nsmap['r']+'}embed']

    img = doc.part.rels[rId].target_ref
    ext=os.path.splitext(img)[-1]
    img_part = doc.part.rels[rId].target_part

    fname = str(uuid.uuid1()).replace('-', '')   ###convert all image to .png
    if ext=='.wmf':  ###
        out_img=  fname+'.svg'
        x=wmf2svg(img_part.blob,  os.path.join(img_dir, out_img))
    else:
        out_img=fname+'.png'
        image=Image.open(io.BytesIO(img_part._blob))
        image.save(os.path.join(img_dir, out_img))          ###convert all image to .png by PIL

    html = '<img src="' + http_head + out_img + '" width=' + "{:.4f}".format(one_mes["width"]) + \
           ' height=' + "{:.4f}".format(one_mes["height"]) + '>'

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

###检测1个run里面是否有多种内容，这种情况是无效的！！
def check_run(child):
    tag = get_tag(child)
    if tag == 'm:oMath' or tag == 'm:oMathPara' or tag=='w:tab':
        return  1
    i = 0

    run = child.__copy__()
    rPr = run.xpath('.//w:rPr', namespaces=run.nsmap)  ##删除run的属性

    if len(rPr) > 1:
        print('docx格式出错了，len(w:rPr)!=1')
    elif len(rPr) == 1:
        run.remove(rPr[0])
    # else: 没有rPr,啥都不用做

    wt = run.xpath('./w:t', namespaces=run.nsmap)
    wdrawing = run.xpath('.//w:drawing', namespaces=run.nsmap)
    wpict=run.xpath('.//w:pict', namespaces=run.nsmap)
    wobject = run.xpath('.//w:object', namespaces=run.nsmap)
    i = len(wt) + len(wdrawing)+len(wpict)+len(wobject)

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
            # print('run里面没有找到合适元素！,可能有未识别的')
            pass

        if tag == 'w:t':  ##处理文本
            html = w_t2html(child)
            htmls.append(html)
        elif tag == 'w:drawing':  ##处理图片
            result = w_drawing2html(doc, child)
            if result['mode']=='inline':  ##不处理浮动图片，留到别处统一处理
                htmls.append(result['html'])
        elif tag == 'w:pict':
            print('found w:pict')
            html=w_pict2html(doc, child)['html']
            htmls.append(html)
        elif tag=='w:object':  ##处理可能的ole对象
            pass
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
            # print('options2html: run里面没有找到合适元素！')
            continue

        if tag == 'w:t':
            text = text + child.xpath('./w:t/text()', namespaces=docx_nsmap)[0]
        elif tag == 'w:drawing':
            result = w_drawing2html(doc, child)
            if result['mode']=='inline':  ##不处理浮动图片，留到后面一起处理
                text = text + result['html']
        elif tag == 'm:oMath':
            text = text + o_math2html(child)
        elif tag=='w:pict':
            html = w_pict2html(doc, child)['html']
            text=text+html
        elif tag == 'w:object':
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

        if html.strip()!='':
            htmls.append(html)
    result=''
    for html in htmls:
        result=result+ '<p>'+html+'</p>'

    return result

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
        ti['stem']=ti['stem'] + image_html
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
        question['stem'] = question['stem']  + image_html
        questions.append(question)
        i = i + 1
    ti['questions'] = questions

    return (curr_xiaoti_index + n, ti)


##处理1个小题
def get_xiaoti_content(doc, xiaoti_indexes, curr_index):
    q = {}
    title_indexes = xiaoti_indexes[curr_index]['title']

    xx = paragraphs2htmls(doc, title_indexes)
    q['stem']=re.sub(r'^<p>\d{1,2}[.．]\s{0,}', '<p>', xx)   ###去除题号

    q['number'] = re.findall(r'^<p>(\d{1,2})[.．、]\s{0,}', xx)[0]   ###获取题号

    if 'options' in xiaoti_indexes[curr_index]:
        option_indexes = xiaoti_indexes[curr_index]['options']
        q['options'] = get_option_htmls(doc, option_indexes)
        if '一项' in doc.paragraphs[title_indexes[0]].text:
            q['type']='SINGLE'
    else:
        q['type'] = 'GENERAL'
    return q

if __name__ == "__main__":
    path = '../data/2019年全国II卷文科综合高考真题.docx'
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
