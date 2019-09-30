import docx
from lxml import etree
import uuid
import os
doc=docx.Document('ImgQuestion/test.docx')
if not os.path.exists('img'):
    os.mkdir('img')

def get_tag(el):
    tag = el.tag
    e_index = tag.find('}')
    tag = tag[e_index + 1:]
    tag = tag.split('\s{1}')[0]
    return tag

def git_pic(tree):
    pics = tree.xpath('.//w:drawing' , namespaces = tree.nsmap)
    pic_mes = []
    for pic in pics:
        one_mes = dict()
        size_ele = pic.xpath('.//wp:extent ' , namespaces = pic.nsmap)[0]
        width = int(size_ele.attrib['cx'])/(360000*0.0264583)
        height = int(size_ele.attrib['cy'])/(360000*0.0264583)
        one_mes['width'] = width
        one_mes['height'] = height
        inline_ele = pic.xpath('.//wp:inline' , namespaces = pic.nsmap)[0]
        a_graphic = inline_ele.getchildren()[len(inline_ele)-1]
        blip = pic.xpath('.//a:blip ' , namespaces = a_graphic.nsmap)[0]
        blip_attr = blip.attrib
        for attr in blip_attr:
            value = blip_attr[attr]
            if 'embed' in attr:
                one_mes['rId'] = value
        pic_mes.append(one_mes)

    return pic_mes

def git_reall_pic(doc , pic_list):
    for pic in pic_list:
        pic_name = pic['rId']
        img = doc.part.rels[pic_name].target_ref
        img_part = doc.part.related_parts[pic_name]
        path = str(uuid.uuid1()).replace('-' , '')
        path ='img/'+path+ '.jpeg'
        pic['path'] = path

        tag = '<img src="'+path+'" width='+str(pic["width"])+' height='+str(pic["height"])+'>'
        pic['tag'] = tag

        with open(path,'wb') as f:
            f.write(img_part._blob)

    return pic_list

index = 0

all_pic_list = []
paragraphs = doc.paragraphs

while(1):
    tree = etree.fromstring(paragraphs[index]._element.xml)
    result = git_pic(tree)
    pic_list = git_reall_pic(doc , result)
    all_pic_list.extend(pic_list)
    # 对象里面的tag就是要显示的标签
    index+=1
    if index == len(paragraphs):
        break


