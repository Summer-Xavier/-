import re
import json
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from os import listdir



def find_table_after_title(doc, target_title):
    under_target_title = False

    for element in doc.element.body.iterchildren():
        if isinstance(element, CT_P):
            paragraph = Paragraph(element, doc)
            text = paragraph.text
            istitle = (True if re.match(r'\d+\.\d+.+?', text, re.DOTALL) else False)
            isin = (target_title in text)
            if istitle and isin:
                under_target_title = True
            elif istitle:
                under_target_title = False
            # elif under_target_title and istitle:  为什么不需要按这段被注释的代码，加一个under_target_title的判定条件
            #     under_target_title = False
            # else:
            #     under_target_title = False   为什么不用注释段？这个问题很典

        elif isinstance(element, CT_Tbl) and under_target_title:
            table = Table(element, doc)
            return table

    return None


def get_type1_table_content(table):
    lst = []
    latest_cell = None
    for row in table.rows:
        for cell in row.cells:
            temp_latest_cell = cell.text
            if temp_latest_cell != latest_cell:
                latest_cell = temp_latest_cell
                lst.append(latest_cell)
            else:
                continue
    dic = {lst[i]: lst[i+1] for i in range(len(lst)) if i % 2 == 0}
    return dic


def get_type2_table_content(table):
    if table:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        json_list = []
        for row in table.rows[1:]:
            row_data = [cell.text.strip() for cell in row.cells]
            data_dict = dict(zip(headers, row_data))
            json_list.append(data_dict)

        return json_list
    else:
        return None



if __name__ == '__main__':
    ## 文件集合信息
    path = r'D:\My Researches\数据集团\企查查报告'
    files = [path + '\\' + f for f in listdir(path) if f.endswith('.docx') and '企业信用报告专业版' in f]

    ## 获取的信息
    type1_table_names = ['工商信息']
    type2_table_names = ['股东信息',
                         '主要人员',
                         '对外投资',
                         '间接持股企业',
                         '控制企业',
                         '招投标',
                         '招聘',
                         '供应商',
                         '客户',
                         '上榜榜单',
                         '商标信息',
                         '专利信息',
                         '作品著作权',
                         '软件著作权',
                         '资质证书'
                         ]

    total_data = []
    for file in files:
        doc = Document(file)
        file_dic = {}
        for type1 in type1_table_names:
            file_dic.update({
                type1: get_type1_table_content(find_table_after_title(doc, type1))
            })
        for type2 in type2_table_names:
            file_dic.update({
                type2: get_type2_table_content(find_table_after_title(doc,type2))
            })
        total_data.append(file_dic.copy())

    print(total_data)