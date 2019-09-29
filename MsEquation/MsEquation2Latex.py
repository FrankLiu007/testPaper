from dwml import omml
from lxml import etree
import docx
import uuid
import mammoth
def get_latex(ml):
    ###only have one equation,

    if 'oMathPara' in ml.getparent().tag:
        mm=etree.tostring(ml.getparent()).decode('utf-8')
    else:
        tt=etree.Element('oMathPara')
        tt.append(ml.__copy__())  ####使用copy， 不影响原来的结构
        mm=etree.tostring(tt).decode('utf-8')

    for math in  omml.load_string(mm):
        # print(math.latex)
        return  math.latex

def MSequation2latex(doc):
    tree = doc._element
    math_elements = tree.xpath('.//m:oMath')
    math_text = '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>$$ {0}$$</w:t></w:r>'

    for ml in math_elements:
        latex = get_latex(ml)
        txt = '<w:r><w:t>$$ ' + latex + '$$</w:t></w:r>'
        t1 = etree.fromstring(math_text.replace('{0}', latex))
        p = ml.getparent()
        print(p.tag)
        if 'oMathPara' in p.tag:
            print('find oMathPara')
            p.getparent().replace(p, t1)
            # p.getparent().addprevious(t1)
            # p.getparent().remove(p)
        else:
            print('not find oMathPara')
            p.replace(ml, t1)
            # p.remove(ml)
    return doc
def get_html(str1):
    document = "<!DOCTYPE html><html lang='zh_CN'><head><meta charset='UTF-8'><title>Document</title><style>table,table td,table th{border:1px solid;border-collapse: collapse;}</style></head><body>";
    document = document + str1
    document = document + "<script type='text/x-mathjax-config'>MathJax.Hub.Config({tex2jax: {inlineMath: [['$','$'], ['@@','@@']]}});</script>"
    document = document + "<script type='text/javascript' async src='https://cdn.mathjax.org/mathjax/latest/MathJax.js?config=TeX-AMS_CHTML'></script></body></html>"
    return  document
if __name__ == "__main__":
    path='MsEquation/equation1.docx'
    doc=docx.Document(path)
    MSequation2latex(doc)
    doc.save('MsEquation/out.docx')
    with open("MsEquation/out.docx", "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html =get_html(result.value)
        with open('MsEquation/abc.html', 'w', encoding='utf-8') as f:
            f.write(html)
