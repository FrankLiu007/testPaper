import zipfile
from lxml import etree
import os
import re
import roman
import pycnnum
from docx_utils.namespaces import namespaces as docx_nsmap
class Document:
    ##公开属性
    rIds={}
    files = []
    elements=[]
    numbering={}
    zip_handle=None
    w_val = '{' + docx_nsmap['w'] + '}val'

    def numPr2text(self, child):
        numId=child.xpath('.//w:pPr/w:numPr/w:numId', namespaces=docx_nsmap)[0].attrib[self.w_val]
        ilvl=child.xpath('.//w:pPr/w:numPr/w:ilvl', namespaces=docx_nsmap)[0].attrib[self.w_val]
        if not ( numId in self.numbering):
            return ''
        numFmt=self.numbering[numId][ilvl]['numFmt']
        lvlText=self.numbering[numId][ilvl]['lvlText']

        curr_num = self.numbering[numId][ilvl]['current']
        tt=''
        if numFmt=="chineseCountingThousand":
            tt=re.sub(r"%\d{1,2}", pycnnum.num2cn(curr_num) ,lvlText)
            self.numbering[numId][ilvl]['current'] += 1
        elif numFmt=="lowerLetter":
            tt = re.sub(r"%\d{1,2}", chr(ord('a')+curr_num), lvlText)
            self.numbering[numId][ilvl]['current'] += 1
        elif numFmt=="upperLetter":
            tt = re.sub(r"%\d{1,2}", chr(ord('A')+curr_num), lvlText)
            self.numbering[numId][ilvl]['current'] += 1
        elif numFmt=="lowerRoman":
            tt = re.sub(r"%\d{1,2}", roman.toRoman(curr_num), lvlText)
            self.numbering[numId][ilvl]['current']+=1
        elif numFmt=="upperRoman":
            tt = re.sub(r"%\d{1,2}", roman.toRoman(curr_num), lvlText)
            self.numbering[numId][ilvl]['current']+=1
        elif numFmt=="decimal":
            tt = re.sub(r"%\d{1,2}", str(curr_num), lvlText)
            self.numbering[numId][ilvl]['current'] += 1
        else:  ###其它情况一概用中文的一（居然遇到japaneseCounting）
            tt=re.sub(r"%\d{1,2}", pycnnum.num2cn(curr_num) ,lvlText)
            self.numbering[numId][ilvl]['current'] += 1
        return tt

    def get_text(self, child):
        numPr=child.xpath('.//w:pPr/w:numPr', namespaces=docx_nsmap)
        tt=''
        if numPr:
            tt=self.numPr2text(child)
        text=child.xpath('.//w:t/text()', namespaces= docx_nsmap)
        return tt + ''.join(text)
###
    def get_elements(self):
        children=self.doc_root.xpath('.//w:body', namespaces=docx_nsmap)[0].getchildren()
        elements=[]
        for child in children:
            txt=self.get_text(child)
            elements.append({'text':txt,'element':child})
        return elements

###获取numbering
    def get_numbering(self, zip_handle):
        path = 'word/numbering.xml'
        numbering={}
        numIds={}
        abstractNums={}
        abstractNums={}
        if path in self.files:
            f = zip_handle.open(path, 'r')
            tree = etree.fromstring(f.readlines()[1])
            f.close()
            w_val = '{' + docx_nsmap['w'] + '}val'
            w_numId='{' + docx_nsmap['w'] + '}numId'
            w_abstractNumId='{' + docx_nsmap['w'] + '}abstractNumId'
            w_ilvl='{' + docx_nsmap['w'] + '}ilvl'

            for abstractNum  in tree.xpath('.//w:abstractNum', namespaces=docx_nsmap):
                id=abstractNum.attrib[w_abstractNumId]
                lvls={}
                for lvl in abstractNum.xpath('./w:lvl', namespaces=docx_nsmap):
                    lvl_info={}
                    ilvl=lvl.attrib[w_ilvl]
                    lvl_info['start']=lvl.xpath('./w:start', namespaces=docx_nsmap)[0].attrib[w_val]
                    lvl_info['numFmt']=lvl.xpath('./w:numFmt', namespaces=docx_nsmap)[0].attrib[w_val]
                    lvl_info['lvlText']=lvl.xpath('./w:lvlText', namespaces=docx_nsmap)[0].attrib[w_val]
                    lvl_info['current']=int(lvl_info['start'])
                    lvls[ilvl]= lvl_info.copy()

                abstractNums[id]=lvls.copy()
            for num in tree.xpath('.//w:num', namespaces=docx_nsmap):

                abstractNumId=num.xpath('./w:abstractNumId', namespaces=docx_nsmap)[0].attrib[w_val]
                numIds[num.attrib[w_numId]]=abstractNums[abstractNumId]
            return numIds

    ###获取rId列表
    def process_rIds(self, zip_handle):
        path='word/_rels/document.xml.rels'
        rIds={}
        if path in self.files:
            f=zip_handle.open(path, 'r')
            tree=etree.fromstring(f.readlines()[1])
            f.close()
            for relation in tree.xpath('.//*[local-name()="Relationship"]', namespaces=docx_nsmap):
                resource={}
                id=relation.attrib['Id']
                resource['id']=id
                path=relation.attrib['Target']
                if path.startswith('..'):
                    path=path.replace('../', '')
                else:
                    path='word/'+path

                f2=zip_handle.open(path, 'r')
                resource['blob']=f2.read()
                f2.close()
                resource['path']=path
                rIds[id]=resource.copy()

            return rIds

    def get_file_list(self, zip_handle):
        files=[]
        for f in zip_handle.filelist:
            files.append(f.filename)
        return files

    def read_document(self, zip_handle):
        self.inital_read(zip_handle)    ####做一些预处理工作

        f = zip_handle.open('word/document.xml','r')
        self.doc_xml = f.readlines()[1]
        f.close()
        self.doc_root=etree.fromstring(self.doc_xml)
        self.elements=self.get_elements()

    ####初始读一些参数，保证文档能正常
    def inital_read(self, zip_handle):
        self.files=self.get_file_list(zip_handle)
        self.rIds=self.process_rIds(zip_handle)
        self.numbering=self.get_numbering(zip_handle)
        pass
    def __init__(self, path):

        self.zip_handle=zipfile.ZipFile(path, 'r')
        self.read_document(self.zip_handle)
        self.zip_handle.close()

if __name__ == "__main__":
    path='d:/test/崇阳一中2020届高三理科数学测试卷.zip'
    doc=Document(path)