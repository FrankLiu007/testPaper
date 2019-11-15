import lxml.etree as ET
import docx
path='MsEquation/omml2mml.xsl'
tree = docx.Document('MsEquation/equation1.docx')._element
math_elements = tree.xpath('.//m:oMath')

xslt = ET.parse(path)
transform = ET.XSLT(xslt)

newdom = transform(math_elements[0])
print(ET.tostring(newdom))    ###打印出来的mml的xml，复制粘贴到word直接变成公式


###mathml(mml) 在html里面的用法
##公式内容放<math>标签里面即可