import zipfile
import os
import lxml.etree as ET
import uuid
import io
from PIL import Image
###docx_nsmap  命名空间
docx_nsmap={
    'w14':'http://schemas.microsoft.com/office/word/2010/wordml',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
    'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
    'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
    'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
    'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
    'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
    'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
    'am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp':
        'http://schemas.openxmlformats.org/drawingml/2006/'
        'wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

#--------读取所有rId
def read_rIds( zip_handle):
    path = 'word/_rels/document.xml.rels'
    rIds = {}
    tt=zip_handle.open(path, 'r').read()
    tree = ET.fromstring(tt.decode('utf8').splitlines()[1])
    for relation in tree.xpath('.//*[local-name()="Relationship"]', namespaces=docx_nsmap):
        id = relation.attrib['Id']
        path = relation.attrib['Target']

        if path.startswith('..'):
            path = path.replace('../', '')
        else:
            path = 'word/' + path
        if 'TargetMode' in relation.attrib:  ####外部的资源文件，不要读取
            if relation.attrib['TargetMode'] == "External":
                continue
        rIds[id] = path

    return rIds
###
# def wmf2png(wmf_bolb):
#     image = Image.open(io.BytesIO(wmf_bolb)).save(fname+'.png')

path='test.docx'
out_path='tmp'
doc={}
zip_handle=zipfile.ZipFile(path, 'r')
rIds=read_rIds(zip_handle)
xx=zip_handle.open('word/document.xml').read()
xml=ET.fromstring(xx)

for ole in xml.xpath('.//w:object', namespaces=docx_nsmap):
    tt=ole.xpath('.//o:OLEObject', namespaces=docx_nsmap)
    if tt[0].attrib['ProgID']=='Equation.DSMT4':
        fname=uuid.uuid1().hex
        r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        eq_rId=tt[0].attrib[r]
        ff=zip_handle.open(rIds[eq_rId])
        eq_obj_blob=ff.read()
        with open(os.path.join(out_path, fname+'.bin'),'wb') as f:
            f.write(eq_obj_blob)
            f.close()

        img_rId=ole.xpath('.//v:shape/v:imagedata', namespaces=docx_nsmap)[0].attrib[r]
        zz=zip_handle.open(rIds[img_rId])
        img_obj_blob=zz.read()
        Image.open(io.BytesIO(img_obj_blob)).save(os.path.join(out_path,fname + '.png'))   ###wmf2png
