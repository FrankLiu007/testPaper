import zipfile
from lxml import etree
from .namespaces import namespaces as docx_nsmap
from . import MyDocx
path=''
doc=MyDocx.Document(path)
