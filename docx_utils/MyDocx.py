import zipfile
from lxml import etree
class MyDocx():

    def get_text(self):
        pass
    def __init__(self, path):
        zip_handle=zipfile.ZipFile(path, 'r')
        f=zip_handle.open('word/document.xml')
        self.xml=f.readlines()[1]    ####
        f.close()
        self.tree=etree.fromstring(self.xml)