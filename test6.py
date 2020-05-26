
import   docx_utils.MyDocx as MyDocx
import  docx_utils.settings as settings
from docx_utils.parse_all import get_answer_start_row, parse_all




if __name__ == "__main__":
    path='data/标准测试-2018年全国英语真题.docx'
    settings.init()
    settings.strict_mode=False
    settings.dati_mode=['第', 'X', '部分']
    doc=MyDocx.Document(path)
    row = get_answer_start_row(doc)
    all_ti=parse_all(doc,1, row )




# from lxml import etree
# from docx_utils import settings
# import re
# from docx_utils import MyDocx
# def parse_all(doc,start_row, end_row):
#
#     curr_row=start_row-1
#     tree=etree.Element('root')
#     ti=''
#     question=''
#     jie=''
#     dati=''
#     while curr_row<=end_row:
#         curr_row+=1
#         txt=doc.elements[curr_row]['text'].strip()
#
#         yy = re.findall(r'^第[一二三四五六七八九]部分', txt)  ###某部分遇到了
#         zz = re.findall(r'^[一二三四五六七八九][\s\.．、]', txt)  # 某个大题遇到了
#         if yy or zz:
#             if type(dati)==etree._Element:
#                 tree.append(dati)
#             dati=etree.Element('dati')
#             dati.text=str(curr_row)
#             continue
#
#         yy = re.findall(r'第[一二三四五六七八九][小]{0,1}节', txt)  # 某个小节遇到了
#         if yy:
#             if type(jie)==etree._Element:
#                 dati.append(jie)
#             jie = etree.Element('jie')
#             jie.text = str(curr_row)
#             continue
#
#         yy = re.findall(settings.mode_text, txt)  ###材料题出现了？
#         if yy:
#             if type(ti)==etree._Element:
#                 if type(jie)==etree._Element:
#                     jie.append(ti)
#                 else:
#                     dati.append(ti)
#
#             ti=etree.Element('ti')
#             ti.text=str(curr_row)
#             continue
#
#         yy = re.findall(r'^(\d{1,2})[\s\.．、]', txt)  ###这个小题出现了
#         if yy:
#             if type(question)==etree._Element:
#                 if ti:##材料题的情况
#                     ti.append(question)
#                 else:###不是材料题的情况
#                     ti=etree.Element('ti')
#                     ti.append(question)
#                     dati.append(ti)
#                     ti=''
#
#             question=etree.Element('question')
#             question.attrib['number']=yy[0]
#             continue
#
#         yy = re.findall(r'^[A-G][\s\.．、]', txt)  ###遇到选项了
#         if yy:
#             option=etree.Element('option')
#             option.text=str(curr_row)
#             question.append(option)
#             continue
#     return tree
# def get_answer_start_row(doc):
#     row=-1
#     for i in range(0,len(doc.elements)) :
#         if '参考答案' in doc.elements[i]['text']:
#             row= i+1
#             break
#     return row
# if __name__ == "__main__":
#     path="d:/test.docx"
#     settings.init()
#     doc=MyDocx.Document(path)
#     row=get_answer_start_row(doc)
#     tree=parse_all(doc, 1, row-1)
#     print(etree.tostring(tree,encoding='utf8', pretty_print=True).decode('utf8'))
#









# from lxml import etree
# import docx
# from docx_utils import settings
# from docx_utils.ti2html import paragraph2html
# from docx_utils.namespaces import namespaces as docx_nsmap
# settings.init()
# path = 'd:/test/表格.docx'
# doc = docx.Document(path)
# htmls = []
# tree=etree.fromstring(doc.element.xml)
# body=tree.xpath('.//w:body', namespaces= docx_nsmap)[0]
# for child in body.getchildren():
#
#     htmls.append(paragraph2html(doc, child))
# import sys
# import openpyxl
#
# def find_start_row(worksheet):
#     ##找到标题行
#     rows=worksheet.max_row
#     for i in range(1, rows + 1):
#         x = ''
#         for j in range(0, len(worksheet[i])):
#             x = x + str(worksheet[i][j].value)
#         print('x=', x)
#         if '姓名' in x:
#             return i + 1
#
# path=sys.argv[1]
# out_xlsx=sys.argv[2]
#
# workbook=openpyxl.load_workbook(path)
# worksheet=workbook.worksheets[0]
# columns=worksheet.max_column
# rows=worksheet.max_row
#
# ##找到标题行
# start_row=find_start_row(worksheet)
#
# row_info=[]
# ##读取标题的数据
# for j in range(0,len(worksheet[start_row-1])):
#     row_info.append(worksheet[start_row-1][j].value)
#
# ###找出所有班级
# jj=row_info.index('班级')
# classes=set()
# for i in range(start_row, rows):
#     cl=worksheet[i][jj].value
#     if cl:
#         classes.add(cl)
#
# persons=[]
# for i in range(start_row, rows + 1):
#     p={}
#     for j in range(0, len(worksheet[i])):
#         p[row_info[j]]=worksheet[i][j].value
#     persons.append(p.copy())
#
#
# total={}
#
# for class0 in classes:
#     total[class0]=[0]*13   ###13个元素。
#
#
# for person in persons:
#     if person['政治总'] and person['历史总'] and person['地理总']:
#         person['总分']= person['地理总']+person['政治总']+person['历史总']
#     else:
#         continue
#     class0=person['班级']
#
#
#     if person['总分']>= total[class0][12]:
#         total[class0][12]=person['总分']
#
#     if person['总分']>=160:
#         total[class0][0]=total[class0][0]+1
#         if person['总分']>=180:
#             total[class0][1] = total[class0][1] + 1
#             if   person['总分'] >= 200:
#                 total[class0][2] = total[class0][2] + 1
#
#     if person['政治总']>=50:
#         total[class0][3]=total[class0][3]+1
#         if person['政治总']>=60:
#             total[class0][4] = total[class0][4] + 1
#             if   person['政治总'] >= 70:
#                 total[class0][5] = total[class0][5] + 1
#
#     if person['历史总']>=50:
#         total[class0][6] = total[class0][6] + 1
#         if person['历史总']>=60:
#             total[class0][7] = total[class0][7] + 1
#             if   person['历史总'] >= 70:
#                 total[class0][8] = total[class0][8] + 1
#
#     if person['地理总']>=50:
#         total[class0][9]=total[class0][9]+1
#         if person['地理总']>=60:
#             total[class0][10] = total[class0][10] + 1
#             if   person['地理总'] >= 70:
#                 total[class0][11] = total[class0][11] + 1
#
# out_sheet=workbook.create_sheet("统计结果")
# out_sheet.insert_cols(1,50)
# out_sheet.insert_cols(1,50)
# ss=['总分>=160', '总分>=180','总分>=200','政治总分>=50','政治总分>=60','政治总分>=70',\
#     '历史总分>=50','历史总分>=60','历史总分>=70','地理总分>=50', '地理总分>=60', '地理总分>=70', '总分最高分']
#
# out_sheet[1][0].value="sadf"
# for i in range(0,13):
#     out_sheet[1][i+1].value=ss[i]
#
# row=2
# for key in total:
#     out_sheet[row][0].value=key
#     for i in range(0,13):
#         out_sheet[row][i+1].value=total[key][i]
#     row=row+1
# workbook.save(out_xlsx)
#
#
#测试亲情话机
# import socket
# ip='47.110.42.14'
# ip2='47.110.139.213'
# data="0010051412"   ###心跳包
# port=7070
# sock = socket.socket()
# sock.connect((ip2, port))
# sock.sendall(data.encode())
# recv=sock.recv(1024)
# print(recv.decode())
#
#
# while(1):
#     i=i+1
#     conn,addr=server.accept()
#     data=conn.recv(1024)
#     print(data)
#     conn.send('get data:'+data)
#     conn.close()
#
#
# import socket
# server=socket.socket()
# ip='0.0.0.0'
# port=22
# server.bind((ip, port))
# server.listen()
# s.connect((ip, port))
# data=input("请输入一些文字：")
# s.send(data.encode())
# recvData=s.recv(1024)
# print('received:', recvData)
# s.close()
#
#
#
# import socket
# ip='0.0.0.0'
# port=22
# server = socket.socket()
# server.bind((ip, port))
# server.listen()
# print('开始接收消息!')
# i=0
# while(1):
#     i=i+1
#     conn,addr=server.accept()
#     data=conn.recv(1024)
#     print(data)
#     conn.send(data)
#     conn.close()