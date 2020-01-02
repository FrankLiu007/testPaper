import zipfile
from lxml import etree
import os
import re
import roman
import pycnnum
from docx_utils.namespaces import namespaces as docx_nsmap
import subprocess
import uuid
class Document:
    ##静态属性
    w_val = '{' + docx_nsmap['w'] + '}val'
    omml2mml_transform = etree.XSLT(etree.parse('docx_utils/omml2mml.xsl'))
    def wmf_emf2png(self):
        img_lst=[]
        for rId in self.rIds:
            rId=self.rIds[rId]
            fname, ext=os.path.splitext(rId['path'])
            if ext== '.wmf' or ext== '.emf':
                img_lst.append(rId)

        if not os.path.exists('tmp'):
            os.mkdir('tmp')

        for rId in img_lst:
            fname=os.path.split(rId['path'])[-1]
            path=os.path.join('tmp', fname)
            with open(path, 'wb') as f:
                f.write(rId['blob'])
                f.close()

        with open('flist.txt', 'w') as f:
            for rId in img_lst:
                pp=os.path.split(rId['path'])[-1]
                fin=os.path.join('tmp', pp)
                fout=fin[:-4]+'.png'
                print(fin+'  '+fout+'  0   0', file=f)
            f.close()
        status,output=subprocess.getstatusoutput( os.path.join('docx_utils', 'WmfEmf2png.exe')+' -l flist.txt')

        ###update file_blobs

        for rId in img_lst:
            # fname, ext=
            pp = os.path.split(rId['path'])[-1]
            path=rId['path']
            self.file_blobs.pop(rId['path'])  ###删除emf文件
            with open(os.path.join('tmp',pp[:-4]+'.png'), 'rb') as f:
                self.file_blobs[path[:-4]+'.png']=f.read()
                f.close()
            old_path = rId['path'][5:]  ###remove "word/"
            new_path=old_path[:-4]+'.png'
            rId['blob']=self.file_blobs[path[:-4]+'.png']   ##这个要带word/
            rId['path']=path[:-4]+'.png'
            ###update 'word/_rels/document.xml.rels'
            print('old_path:',old_path, 'new_path', new_path)
            self.file_blobs['word/_rels/document.xml.rels']=self.file_blobs['word/_rels/document.xml.rels'].replace(old_path.encode(), new_path.encode())


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
        if path in self.file_blobs:
            tree = etree.fromstring(self.file_blobs[path].decode('utf8').splitlines()[1])
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
        if path in self.file_blobs:
            tree=etree.fromstring(self.file_blobs[path].decode('utf8').splitlines()[1])
            for relation in tree.xpath('.//*[local-name()="Relationship"]', namespaces=docx_nsmap):
                resource={}
                id=relation.attrib['Id']
                resource['id']=id
                path=relation.attrib['Target']

                if path.startswith('..'):
                    path=path.replace('../', '')
                else:
                    path='word/'+path
                if 'TargetMode'in relation.attrib:   ####外部的资源文件，不要读取
                    if relation.attrib['TargetMode']=="External":
                        resource['path'] = path
                        rIds[id] = resource.copy()
                        continue
                resource['blob']=self.file_blobs[path]

                resource['path']=path
                rIds[id]=resource.copy()

            return rIds
    ##一次读取所有文件
    def read_all_files(self, zip_handle):
        file_blobs={}
        for f in zip_handle.filelist:
            hh=zip_handle.open(f.filename, 'r')
            file_blobs[f.filename]=hh.read()
            hh.close()
        return file_blobs

    def read_document(self, zip_handle):
        self.inital_read(zip_handle)    ####做一些预处理工作
        path='word/document.xml'
        self.doc_xml=self.file_blobs[path].decode('utf8').splitlines()[1]  ###

        self.doc_root=etree.fromstring(self.doc_xml)
        self.elements=self.get_elements()

    ####初始读一些参数，保证文档能正常
    def inital_read(self, zip_handle):
        self.file_blobs=self.read_all_files(zip_handle)
        self.rIds=self.process_rIds(zip_handle)
        self.numbering=self.get_numbering(zip_handle)
        pass
    def __init__(self, path):
        self.rIds={}
        self.fname=path
        self.file_blobs = {}
        self.elements=[]
        self.numbering={}
        zip_handle=zipfile.ZipFile(path, 'r')
        self.read_document(  zip_handle )
        zip_handle.close()
    ###保存zip文件
    def save(self, outf=''):
        if not outf:
            outf=self.fname
        zip_f=zipfile.ZipFile(outf,'w')
        for f in self.file_blobs:
            blob=self.file_blobs[f]
            if blob:
                zip_f.writestr(f, blob)
        zip_f.close()
    def __del__(self):
        pass
if __name__ == "__main__":
    path='d:/test/崇阳一中2020届高三理科数学测试卷.zip'
    doc=Document(path)