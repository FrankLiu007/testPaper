from lxml import etree
from docx_utils.namespaces import namespaces as docx_nsmap
import docx
### 获取标签
def get_tag(child):
    '''
    :param child: etree.Element,一般为一个run
    :return: 返回该run里面的内容的类型的字符串
    '''
    if child.tag.split('}')[-1] == 'pPr':
        return 'w:pPr'

    wt = child.xpath('.//w:t', namespaces=child.nsmap)
    if len(wt) == 1:
        return 'w:t'
    wdrawing = child.xpath('.//w:drawing', namespaces=child.nsmap)
    if len(wdrawing) == 1:
        return 'w:drawing'
    moMath = child.xpath('.//m:oMath', namespaces=child.nsmap)
    if len(moMath) == 1:
        return 'm:oMath'

def merge_wt(tree):  ###一个段落
    children = tree.getchildren()
    i = 0
    result = etree.Element('{' + docx_nsmap['w'] + '}p', nsmap=docx_nsmap)
    last_is_wt = False
    last_wt = ''
    last_wt_run=''
    last_rPr=''
    while (i < len(children)):
        child = children[i]
        tag = get_tag(child)
        if tag=='w:pPr':
            i=i+1
            continue
        if tag=='w:pPr':
            i=i+1
            continue

        if tag == 'w:t':  ###当前run是w:t
            if last_is_wt:  ##上一个run也是w:t
                if  check_rPr( last_wt_run, child): #上一个run的w:t和这次的格式一样，可以合并
                    text = child.xpath('.//w:t/text()', namespaces=docx_nsmap)[0]
                    print('last_wt=', last_wt, 'i=', i)
                    last_wt=last_wt_run.xpath('.//w:t', namespaces=docx_nsmap)[0]

                    print('text=',text)
                    # last_wt=last_wt[0]
                    print('last_wt_text=', last_wt.text)
                    last_wt.text= last_wt.text + text
                else: #上一个run的w:t和这次的格式不一样，
                    result.append(last_wt_run.__copy__())
                    last_wt_run = child
                    last_is_wt = True
            else:###上一个run不是w:t
                last_wt_run = child
                last_is_wt = True

        else:
            if last_is_wt :
                print('append result',last_wt_run.text)
                result.append(last_wt_run.__copy__())
                last_wt_run = ''
                last_is_wt = False

            result.append(child.__copy__())

        i = i + 1
    ##-----------------while结束

    if last_wt_run!='' and last_is_wt:   ###如果最后一个run是w:t，需要处理
        print('append result', last_wt_run.text)
        result.append(last_wt_run.__copy__())

    return result
def check_rPr(last_run, run):  ##检查run的property是否相同
    #加粗
    w_b=last_run.xpath('./w:rPr/w:b', namespaces=last_run.nsmap)
    bb = run.xpath('./w:rPr/w:b', namespaces=run.nsmap)
    if len(w_b)!=len(bb):
        return False

    #倾斜
    w_i=last_run.xpath('./w:rPr/w:i', namespaces=last_run.nsmap)
    bb = run.xpath('./w:rPr/w:i', namespaces=run.nsmap)
    if len(w_i)!=len(bb):
        return False

    #（双）下划线
    kk = '{' + docx_nsmap['w'] + '}val'

    w_u=last_run.xpath('./w:rPr/w:u', namespaces=last_run.nsmap)
    bb = run.xpath('./w:rPr/w:u', namespaces=run.nsmap)
    if len(w_u)!=len(bb):
        return False
    else:
        if len(w_u) == 1:
            val = w_u[0].attrib[kk]
            if bb[0].attrib[kk]!=val:
                return False
    #上下标
    w_vertAlign=last_run.xpath('./w:rPr/w:vertAlign', namespaces=last_run.nsmap)
    bb = run.xpath('./w:rPr/w:vertAlign', namespaces=run.nsmap)
    if len(w_vertAlign)!=len(bb):
        return False
    else:
        if len(w_vertAlign)==1:
            val = w_vertAlign[0].attrib[kk]
            if bb[0].attrib[kk]!=val:
                return False
    return True

if __name__ == "__main__":
    path='src/test.docx'
    doc=docx.Document(path)
    tree=etree.fromstring(doc.paragraphs[0]._element.xml)
    result=merge_wt(tree)

    for item in result:
        print(item.xpath('./w:t/text()', namespaces=docx_nsmap))