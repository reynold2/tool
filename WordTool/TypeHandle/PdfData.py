import pdfplumber
import pandas as pd
import xlwt
from xpinyin import Pinyin

def pdftest(path="data/系统监控工作站需求规格说明-【内部】.pdf"):
    p = Pinyin()
    writebook = xlwt.Workbook()  # 打开一个excel

    sheet = writebook.add_sheet("test1")
    table_index = 0


    with pdfplumber.open(path) as pdf:

        pages = pdf.pages[1:15]  # 第一页的信息

        # text = page.extract_text()
        # print(text)
        for pageindex,page in enumerate(pages):

            table = page.extract_tables()
            for t in table:
                # print(len(table))
                for text_index, text in enumerate(t):
                    if len(text)==4:
                        continue
                    if table_index!=0:
                        if text[0]=="序号":
                            continue
                    if text[0]=="":
                        continue
                    if text[0]=="单元参数设置命令":
                        continue
                    if text[0]=="分机状态信息上报":
                        continue
                    if text[0]=="过程控制命令":
                        continue
                    if text[0]=="控制结果上报":
                        continue
                    if text[0]=="控制命令响应":
                        continue
                    sheet.write(table_index,3, table_index)
                    dxie = ""
                    try:
                        dxie=p.get_initials(text[1], u'')

                    except:
                        print("当前页数", pageindex,"出现异常",table_index)
                    finally:
                        sheet.write(table_index, 4, dxie)

                    # for context_index,context in enumerate(text):
                    #     # if context_index == 1:
                    #     #     if context[context_index] != "":
                    #     #         print(context[context_index])
                    #     #         zifu=""
                    #     #         for abcdfg in context[context_index]:
                    #     #             temp=p.get_initials(abcdfg, u'')
                    #     #             zifu.append(temp)
                    #     #         sheet.write(table_index, 4, zifu)
                    #     print("rowx:",table_index)
                    #     print("colx:", context_index)
                    #     sheet.write(table_index, context_index,context)

                    table_index = table_index + 1

    writebook.save('data/3.xls')
if __name__=="__main__":
    pdftest()
