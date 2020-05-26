import docx_utils.MyDocx as MyDocx
import re
from lxml import etree
# from dwml import omml   这个包维护比较少，还有部分bug，换包
from docx_utils.namespaces import namespaces as docx_nsmap
import docx_utils.MyDocx as MyDocx
import uuid
import os
import subprocess
from . import settings
from PIL import Image
import io

'''关于docx中的各种元素（数学公式、图片、表格、）
如果用代码生成docx: 1个run里面,可以包含多个文本(w:t)、图片（w:drawing）、公式等，
但是实际上，只要一编辑，1个run里面变只会有1个元素，然后在run里面设置这个元素的属性
数学公式可以和run同级，为oMathPara，可以为run中的元素，为oMath
图片必须为run中的元素（包裹在一个run里面），且该run只有这个图片元素，没有test等

'''

##获取材料题，到底有多少小问（题）
def get_full_tag(str0):
    n,s=str0.split(':')
    return '{'+docx_nsmap[n] + '}'+ s

def get_question_quantity(element, mode_text):
    text = ''.join( element.xpath('.//w:t/text()', namespaces=docx_nsmap) )
    r = re.findall(mode_text, text)
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
        x = re.findall(mode_text, doc.elements[i]['text'])
        if len(x) != 0:
            return i
    return -1

##emu2px
#  1 inch = 914400 EMU
#  1 inch = 72*4/3
def emu2px(emu):
    px=int(emu)/914400*72*4/3
    return px
###1inch =
def inch_pt2px(str0):
    if 'in' in str0:
        ##inch 转px
        return  "{:.1f}".format(float(str0[:-2])*72*4/3)+'px'
    elif 'pt' in str0:
        ##pt 转 px
       return  "{:.1f}".format(float(str0[:-2])*4/3)+'px'
    elif 'px' in str0:
        return str0[:-2]+'px'

## 获取各个选项
def split_options(option_html):
    option_html=''.join( re.findall(r'<p[\s]{0,}.*?>(.*?)</p>', option_html))
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

    w_tblPr='{'+docx_nsmap['w']+'}tblPr'
    if w_pPr==child.tag:
        pass

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
def w_t2html(child, wt_style_ignore=False):

    t = child.xpath('.//w:t/text()', namespaces=child.nsmap)[0]
    t=t.replace('<', '&lt;').replace('>', '&gt;').replace(' ', '&nbsp;')

    kk = '{' + docx_nsmap['w'] + '}val'
    ##1.处理上下标
    ee=child.xpath('./w:rPr/w:vertAlign', namespaces=docx_nsmap)
    if ee:
        val = ee[0].attrib[kk]
        if val=="subscript":
            t='<sub>'+t+'</sub>'
        elif val=='superscript':
            t = '<sup>' + t + '</sup>'
    if wt_style_ignore:   ####处理选项等时候，不需要下面的属性
        return t
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
    fullpath=os.path.join(settings.tmp_dir,  fname)
    with open(fullpath, 'wb') as f:
        f.write(blob)
    cmd='java -jar docx_utils/wmf2svg-0.9.8.jar  '+ fullpath + '  ' + svg_path
    status, output = subprocess.getstatusoutput(cmd)
    if status:
        print('Warning!\r\n ' + cmd + '\r\n wmf转png失败！')
        return 0
    return 1
##mathtype2mml---
def math_type2mml(ole_blob):
    '''
    mathtype嵌入的数学公式,转成mathml是ruby的mathtype_to_mathml做的，(https://github.com/jure/mathtype_to_mathml)
    1. intall mathtype_to_mathml:
    gem install mathtype_to_mathml pry
    2. usage:
    require "mathtype_to_mathml"
    MathTypeToMathML::Converter.new('oleObject1.bin').convert
    exit
    '''
    fname=uuid.uuid1().hex+'.bin'
    if not os.path.exists('tmp'):
        os.mkdir('tmp')
    with open( os.path.join('tmp',fname), 'wb') as f:
        f.write(ole_blob)

    output = subprocess.check_output(['ruby', '-w', 'docx_utils/mathtype_ole2mathml.rb', os.path.join('tmp',fname)])
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
        if ':' in style:
            a,b=style.split(':')
            ole['styles'][a]=b

    ole['height']=inch_pt2px( ole['styles']['height'])
    ole['width']=inch_pt2px(ole['styles']['width'])

    ole_object = child.xpath('.//o:OLEObject', namespaces=docx_nsmap)[0]
    ole['rId'] = ole_object.attrib[get_full_tag('r:id')]
    ole['object_path'] = doc.rIds[ole['rId']]['path']
    ole['blob'] = doc.rIds[ole['rId']]['blob']
    ole['style'] = "vertical-align:middle"

    ProgID=ole_object.attrib['ProgID']
    print('mathtype',ole_object.attrib['ProgID'])
    if ProgID=="Equation.DSMT4":  #是mathtype嵌入的数学公式,转成html
        fname = uuid.uuid1().hex
        if settings.mathtype_convert_to=="mathml":
            mml=math_type2mml(ole['blob'])
            return {'html':mml, 'mode': 'inline'}
        elif settings.mathtype_convert_to=="png":
            img_rId= child.xpath('.//v:imagedata', namespaces=docx_nsmap)[0].attrib[get_full_tag('r:id')]
            ole['img_path'] = doc.rIds[img_rId]['path']
            ext = os.path.splitext(ole['img_path'])[-1]
            if ext == '.wmf':
                out_img_path = fname + '.svg'
                ole['src'] = http_head + fname + '.svg'
                wmf2svg(doc.rIds[img_rId]['blob'], os.path.join(img_dir, out_img_path))

            elif ext == '.png':
                out_img_path = fname + '.png'
                with open(os.path.join(img_dir, out_img_path), 'wb') as f:
                    f.write(doc.rIds[img_rId]['blob'])
                    f.close()
                ole['src'] = http_head + out_img_path

            img = create_img_tag(ole)
            html = etree.tostring(img).decode('utf-8')
            return {'html': html, 'mode': 'inline'}
    else:    ##read v:imagedata only, usually .wmf file， 暂时转换成svg
        pass

    return {'html': '', 'mode': 'inline'}
###创建img标签
def create_img_tag(fig):
    img=etree.Element('img')
    img.attrib['style']=fig['style']
    img.attrib['src']=fig['src']
    img.attrib['height']=fig["height"]
    img.attrib['width'] = fig["width"]
    return img

##------另外一种格式的图片，只有inline 模式-------------
def w_pict2html(doc, child):

    img_dir=settings.img_dir
    http_head=settings.http_head

    pics = child.xpath('.//w:pict', namespaces=docx_nsmap)
    if len(pics) != 1:
        print("docx格式可能错误，w:drawing可能包含多张图片！")
        return 0
    pic = pics[0]

    fig = {}
    fig['styles']={}
    styles=pic.xpath('.//v:shape', namespaces=docx_nsmap)[0].attrib['style']
    for style in styles.split(';'):
        a,b=style.split(':')
        fig['styles'][a]=b
    fig['rId']=pic.xpath('.//v:imagedata', namespaces=docx_nsmap)[0].attrib[get_full_tag('r:id')]
    fig['path'] = doc.rIds[fig['rId']]['path']
    fig['blob']=doc.rIds[fig['rId']]['blob']
    fig['width']=inch_pt2px(fig['styles']['width'])
    fig['height'] =inch_pt2px(fig['styles']['height'])

    ext = os.path.splitext(fig['path'] )[-1]
    fname = uuid.uuid1().hex   ###convert all image to .png
    out_img_path=''
    if ext=='.wmf':  ##扩展名为.wmf
        ## wmf2gd, imagemagick的convert效果都不太好，暂时不用
        ##暂时不用imagemagick的convert
        fullpath=os.path.join(settings.tmp_dir,'out.wmf' )
        with open( fullpath, 'wb') as f :
            f.write(fig['blob'])
        out_img_path=fname+'.svg'
        status, output=subprocess.getstatusoutput('java -jar docx_utils/wmf2svg-0.9.8.jar '+fullpath+ '  ' + os.path.join(img_dir, out_img_path) )
        if  status:
            print('Warning! '+fig['path'] +'转png失败！')
    else:
        image=Image.open(io.BytesIO(fig['blob']))
        out_img_path=fname+'.png'
        image.save(os.path.join(img_dir, out_img_path))          ###convert all image to .png by PIL

    fig['style'] = "vertical-align:middle"
    fig['src'] = http_head + out_img_path
    img=create_img_tag(fig)

    return {'html':etree.tostring(img).decode('utf8'), 'mode':'inline'}

###获取图片的裁剪的区域
def get_crop_box(pic, size0):  ##担心有一天size被系统用了
    width, height=size0
    srcRect = pic.xpath('.//a:srcRect', namespaces=docx_nsmap)

    N=100000    ##docx中的图片为百分比，需要除10万

    if srcRect:
        if not srcRect[0].attrib:  ###没有任何参数，不需要裁剪
            return ()

        left=0
        top=0
        right=width
        bottom=height
        for item in srcRect[0].attrib:
            tt=int(srcRect[0].attrib[item])
            if item=='l':
                left=tt/N*width
            elif item=='t':
                top=tt/N*height
            elif item=='r':
                right=(1- tt/N)*width
            elif item=='b':
                bottom = (1- tt/N)*height

        return (left, top, right, bottom)

    else:
        return ()

###是否是水印图，最明显的是学科网的
##暂时只删除学科网的小黄点，只主要靠ocr
def is_watermark(fig):
    if float(fig['height'])<5  and float(fig['width'])<5 :
        return True
    else:
        return False


##处理w:drawing元素
'''
需要处理inline模式和float模式的2种图片，

'''
def w_drawing2html(doc, child):
    ##fig saved all information of the image
    img_dir=settings.img_dir
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

    fig = dict()
    size_ele = pic.xpath('.//wp:extent ', namespaces=docx_nsmap)[0]

    width = emu2px(size_ele.attrib['cx'])
    height = emu2px(size_ele.attrib['cy'])
    fig['width'] ="{:.0f}".format( width)
    fig['height'] ="{:.0f}".format(height)

    ####直接取出rId
    rId=pic.xpath('.//a:blip ', namespaces=docx_nsmap)[0].attrib['{'+docx_nsmap['r']+'}embed']

    fig['fullpath'] = doc.rIds[rId]['path']
    fig['ext']=os.path.splitext(fig['fullpath'])[-1]

    fig['blob'] = doc.rIds[rId]['blob']
    if is_watermark(fig):
        return {'html':'', 'mode':mode}
    fname = uuid.uuid1().hex   ###convert all image to .png
    if fig['ext']=='.wmf':  ###
        out_img=  fname+'.svg'
        x=wmf2svg(fig['blob'],  os.path.join(img_dir, out_img))
    elif fig['ext']=='.emf':    ###暂时不处理emf图片
        print('found emf image')
        return {'html':'', 'mode':mode}

    else:###
        out_img=fname+'.png'
        image=Image.open(io.BytesIO(fig['blob']) )
        crop_box=get_crop_box(pic, image.size)
        if crop_box:
            image.crop(crop_box).save(os.path.join(img_dir, out_img))
        else:
            image.save(os.path.join(img_dir, out_img))          ###convert all image to .png by PIL
    fig['style']="vertical-align:middle"
    fig['src']=settings.http_head + out_img
    img=create_img_tag(fig)

    return {'html':etree.tostring(img).decode('utf8'), 'mode':mode}

##处理m:oMath元素
def o_math2html(child):
    tag = child.tag.split('}')[-1]
    if tag == 'oMath':
        tt = etree.Element('oMathPara')
        tt.append(child.__copy__())  ####使用copy， 不影响原来的结构
    else:
        tt = child
    ##早期用的dwml转换公式为latex
    # mm = etree.tostring(tt).decode('utf-8')
    # for math in omml.load_string(mm):
    #     text = math.latex
    #     break
    # text = text.replace('<', '&lt;').replace('>', '&gt;')   ##处理大于号、小于号
    # html = '\(' + text + '\)'
    xml=MyDocx.Document.omml2mml_transform(tt)   ###使用微软提供的xsl来转换
    return   etree.tostring( xml).decode('utf8').replace('mml:math','math')

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

##处理表格
def w_tbl2html(doc, child):
    tbl={}
    rows=child.xpath('./w:tblGrid/w:gridCol', namespaces=docx_nsmap)
    tbl['row_number']=len(rows)
    tbl['rows']=[]
    for row in rows:
        width=row.attrib['{'+docx_nsmap['w']+'}w']
        tbl['rows'].append({'width':width} )
    result=''
    row_elements=child.xpath('./w:tr', namespaces = docx_nsmap)

    for i in range(0, len(row_elements)) :
        column_elements=row_elements[i].xpath('./w:tc', namespaces=docx_nsmap)
        for j in range(0,len(column_elements))  :
            html=paragraph2html(doc, column_elements[j].xpath('./w:p', namespaces = docx_nsmap)[0])
            result= result+ '<td>'+html+'</td>'
        result ='<tr>'+ result+'</tr>'
    result= '<table border="1" cellspacing="0">'+ result+ '</table>'

###border="0" cellpadding="3" cellspacing="1" bgcolor="black"

    return result

def calculate_indents(child):
    x=0    ###缩进x个字符，方便web显示
    w_firstline=get_full_tag('w:firstLine')
    w_hanging=get_full_tag('w:hanging')
    w_left=get_full_tag('w:left')
    w_right = get_full_tag('w:right')
    ind=child.xpath('.//w:ind', namespaces=docx_nsmap)
    indents={}
    if ind:
        for key, value in ind[0].attrib.items():
            indents[key]=float(value)/20
        if w_firstline in indents:
            x=float(indents[w_firstline])/10  ####twips to em
        elif w_hanging in indents:
            x=float(indents[w_hanging])/10   ####twips to em
        if w_left in indents:
            x=x+float(indents[w_left]/10)
        return x
    return 0
###
def set_paragraph_property(child):
    pPr=child.xpath('.//w:pPr', namespaces=docx_nsmap)
    p_element = etree.Element('p')
    if not pPr: ###没有找到w:pPr元素
        return p_element
    pPr=pPr[0]
    # style = "text-indent: 2em;"
    ind=calculate_indents(child)
    if ind:
        p_element.attrib['style']="text-indent: "+ str(ind)  +"em;"

    aline= pPr.xpath('.//w:jc', namespaces=docx_nsmap)
    if aline:  ###左对齐和两端对齐，在docx里面没有任何的显示
        xx=aline[0].attrib[get_full_tag('w:val')]
        if xx=='right':
            p_element.attrib['align']='right'
        elif xx == 'center':
            p_element.attrib['align'] = 'center'
        elif xx == 'distribute':
            p_element.attrib['align'] = 'justify'

    return p_element
####段落转html-----------------
def paragraph2html(doc, parent_element, wt_style_ignore=False):

###表格，它本身就是一个段落，处理后返回
    if parent_element.tag=='{'+docx_nsmap['w']+'}tbl':
        return  w_tbl2html(doc, parent_element)

    htmls = []
###处理word里面的行编号-----不准确，暂时放一下！！！
    numPr = parent_element.xpath('.//w:pPr/w:numPr', namespaces=docx_nsmap)
    if numPr:
        paraId = parent_element.attrib[get_full_tag('w:rsidP')]
        htmls.append( doc.numPr[paraId] )
###不是表格的情况

    p_element=set_paragraph_property(parent_element)

    for child in parent_element.getchildren():

        tag = get_tag(child)
        if tag=='w:pPr':
            continue
        vv = check_run(child)
        if vv > 1:
            print('警告！！！run中包含了多个类型，将只处理第一个类型 ')
            # continue
        elif vv == 0:
            # print('run里面没有找到合适元素！,可能有未识别的')
            pass

        if tag == 'w:t':  ##处理文本
            html = w_t2html(child, wt_style_ignore=wt_style_ignore)
            htmls.append(html)
        elif tag == 'w:drawing':  ##处理图片
            result = w_drawing2html(doc, child)
            if result['html']:
                if result['mode']=='inline':  ##不处理浮动图片，留到别处统一处理
                    htmls.append(result['html'])
        elif tag == 'w:pict':
            print('found w:pict')
            html=w_pict2html(doc, child)['html']
            htmls.append(html)
        elif tag=='w:object':  ##处理可能的ole对象
            html = w_object2html(doc, child)['html']
            htmls.append(html)
        elif tag == 'm:oMath' or tag == 'm:oMathPara':  ##处理数学公式
            html = o_math2html( child)
            htmls.append(html)
        elif tag == 'table':  ##处理表格,表格比较特殊，不会出现这种情况
            pass
        # if html != ' ':   空格也应该原样输出，因为有时候存在特意的空格
    xx=etree.tostring(p_element).decode('utf8').replace('/>','>')+ ''.join(htmls) +'</p>'
    return  xx

def check_options(options):
    for i in range(1, len(options)):
        if ord(options[i]['label']) - ord(options[i - 1]['label']) != 1:
            print('获取options错误，请检查')
            return False
    return True


###处理options，得到options的html
def get_option_htmls(doc, options_indexes):
    option_html = ''
    for index in options_indexes:
        option_html += paragraph2html(doc, doc.elements[index]['element'], wt_style_ignore=True)   ###所有options组合起来的html
    options_htmls = split_options(option_html)   ###把每个optiion拆分出来

    if not check_options(options_htmls):
        print('选项识别错误')
        print('options=', options_htmls)

    return options_htmls

##for titel, 不包含选项的段落，可以直接转换
def paragraphs2htmls(doc, title_indexes, wt_style_ignore=False):
    htmls = []
    for index in title_indexes:
        element=doc.elements[index]['element']

        html = paragraph2html(doc, element, wt_style_ignore=wt_style_ignore)
        # print('处理段落：', doc.elements[index]['text'])
        if html.strip()!='':
            htmls.append(html)
    result=''

    for html in htmls:
        result=result+ html
    image_html = get_float_image(doc, title_indexes)
    return result+image_html

###单独处理浮动的图片
def get_float_image( doc, row_list):
    htmls=[]
    for row in row_list:
        element=doc.elements[row]['element']
        x=element.xpath('.//w:drawing/wp:anchor', namespaces=docx_nsmap)
        if x:
            result=w_drawing2html(doc, x[0].getparent().getparent())
            if result['html']:
                if result['mode']=='anchor':
                    htmls.append(result['html'])
    return ''.join(htmls)

####获取每个题的html文本
def get_ti_content(doc, ti_index):

    #####上一个题目的结尾的行号+1
    ti={'questions':[], 'title':''}
    if ti_index['title']:
        ti['title']=paragraphs2htmls(doc, ti_index['title'])
    for question in ti_index['questions']:
        q=get_xiaoti_content(doc,question)
        ti['questions'].append(q)
    return ti


def split_ti_and_number(html):
    # 先找到所有段落，第一个段落里面的题号要删除
    pp=re.findall(r'(<p[\s]{0,}.*?>.*?</p>)',html)
    b = pp[0].find('>')
    txt=re.sub(r'\d{1,2}[\s\.．]','', pp[0][b+1:])
    pp[0]=pp[0][:b+1]+txt
    tt=''.join(pp)

    num=re.findall(r'^<p[\s]{0,}.*?>.*?(\d{1,2}[\.．]\s{0,})', html)[0]
    # if  texts.strip():
    #     xx=re.findall(r'^(\d{1,2}[.．]\s{0,})', texts.strip())
    #     if xx:
    #         num=re.findall(r'^(\d{1,2})[.．]\s{0,}', xx[0])[0]
    #         return ( num, html0.replace(xx[0], '') )
    return (num, tt)
##处理1个小题
def get_xiaoti_content(doc, question):
    q = {}
    title_indexes = question['stem']
    xx = paragraphs2htmls(doc, title_indexes)
    q['number'], q['stem'] = split_ti_and_number(xx)  ##更加安全，可靠的方式，避免题号和选项有加粗的问题


    if 'options' in question:
        q['options'] = get_option_htmls(doc, question['options'])
        if '一项' in   doc.elements[title_indexes[0]]['text'] :
            q['type']='SINGLE'
        else:
            q['type'] = 'MULTIPLE'
    else:
        q['type'] = 'GENERAL'
    return q

if __name__ == "__main__":
    path = '../data/2019年全国II卷文科综合高考真题.docx'
    doc = MyDocx.Document(path)
    # all_ti_index = AnalysQuestion(doc,0,len(doc.elements))
    i = 0
    mode_text = r'\d{1,2}[～-~]\d{1,2}[小]{0,1}题'   ##模式字符串

