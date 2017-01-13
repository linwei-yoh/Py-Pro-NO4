#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from openpyxl import Workbook, load_workbook
import re
import pickle

tar_excel_path = '../xlsx/target.xlsx'
output_path = '../xlsx/output.xlsx'
pick_path = '../save.txt'


# 1.逐个读取各个专利 获取专利名称(唯一） 和 专利所有人
# 2.查询当前专利所有人是否已经在所有人字典中存在 否，则创建并添加 是,则添加专利名称到其专利集合中
# 3.遍历完成，获得所有人字典 记载了全部所有人，及其所拥有的专利
# 4.逐个读取各个专利 获取专利名称(唯一） 和 引用专利(转化为唯一专利）
# 5.查询引用专利是否在专利字典中存在 否，则创建并添加 是,则添加专利名称到引用专利的被引用专利集合中
# 6.遍历完成，获得专利被引用字典 记载了全部被引用专利，及其所引用该专利的专利集合
# 7.逐个遍历所有人字典，对单个所有人的全部专利的被引用集合作并集处理，获得对此所有人的引用
# 8.获得该并集与有该所有人拥有的专利集合的差集(集合相减)，获得有效的专利引用集合，获得中心度
# 9.增加所有人的集合体 公司，专利对象增加 引用集合属性

class GroupItem(object):
    """用于包含公司的集团对象"""

    def __init__(self, name):
        self.name = name
        self.company_set = set()

    def add_company(self, name):
        self.company_set.add(name)


class CompanyItem(object):
    """所有人对象 包含所有人名称 和 所有专利集合"""

    def __init__(self, name):
        self.name = name
        self.patent_set = set()
        self.centre_degrees = 0

    def add_patent(self, name):
        self.patent_set.add(name)


class PatentItem(object):
    """专利对象 包含专利名称 和 引用专利集合"""

    def __init__(self, name):
        self.name = name
        self.cited_set = set()
        self.cite_set = set()

    # 该专利被引用专利集合
    def add_cited_set(self, name):
        self.cited_set.add(name)

    # 该专利引用集合
    def add_cite_set(self, name):
        if isinstance(name, str):
            self.cite_set.add(name)
        elif isinstance(name, set):
            self.cite_set.update(name)


group_dict = {}  # 集团字典
company_dict = {}  # 所有人字典
patent_dict = {}  # 专利字典
unique_name_dict = {}  # 专利唯一名称字典


def add_company_to_group(com_name, group_name):
    """建立集团-公司字典"""
    global group_dict
    if group_name not in group_dict:
        new_com = GroupItem(group_name)
        group_dict[group_name] = new_com
    group_dict[group_name].add_company(com_name)


def add_patent_to_company(name_patent, company_set):
    """建立公司-专利字典"""
    global company_dict
    for company_name in company_set:
        if company_name not in company_dict:
            new_company = CompanyItem(company_name)
            company_dict[company_name] = new_company
        company_dict[company_name].add_patent(name_patent)


def add_patent_and_cite(name_patent, cite_ver):
    """建立专利-引用关系"""
    global patent_dict

    # 添加引用专利对象到字典中
    if name_patent not in patent_dict:
        new_patent = PatentItem(name_patent)
        patent_dict[name_patent] = new_patent
    patent_dict[name_patent].add_cite_set(cite_ver)


def add_patent_and_cited(name_patent, cite_ver):
    """ 建立专利-被引用关系"""
    global patent_dict

    for cite_name in cite_ver:

        # 添加被引用专利对象到字典中 并增加其被引用属性
        if cite_name not in patent_dict:
            new_patent = PatentItem(cite_name)
            patent_dict[cite_name] = new_patent

        patent_dict[cite_name].add_cited_set(name_patent)


def set_unique_name(name_set, unique_name):
    """建立专利号-唯一专利号字典"""
    global unique_name_dict
    for name_item in name_set:
        if name_item not in unique_name_dict:
            unique_name_dict[name_item] = unique_name


def get_unique_name(names):
    """通过专利号获得对应的唯一专利号"""
    if isinstance(names, str):
        if names in unique_name_dict:
            return unique_name_dict[names]
        else:
            return names
    elif isinstance(names, set) or isinstance(names, list):
        result = set()
        for item in names:
            if item in unique_name_dict:
                result.add(unique_name_dict[item])
            else:
                result.add(item)
        return result
    else:
        return None


def cal_centre_degree():
    """中心度计算"""
    global company_dict
    global patent_dict

    print("开始计算中心度")
    # 遍历所有人
    for company_name, company_obj in company_dict.items():
        valid_set = set()
        # 遍历该所有人的全部专利
        for patent in company_obj.patent_set:
            # 且该专利在专利字典中存在
            if patent in patent_dict:
                valid_set = valid_set | patent_dict[patent].cited_set
        valid_set = valid_set - company_obj.patent_set
        # 获得中心度
        company_obj.centre_degrees = len(valid_set)
        # print("所有人: %-60s 中心度为: %-5s" % (company_name, company_obj.centredegrees))


def init_data_from_excel(excel_path):
    """读取excel数据，并建立数据关联用的字典"""
    print('正在读取excel文件.................')
    wb = load_workbook(excel_path)
    sheet_names = wb.get_sheet_names()
    ws = wb.get_sheet_by_name(sheet_names[0])
    print('读取完成,开始数据处理.............')

    # 获得行数
    num_rows = ws.max_row

    # 获得标签行
    tag_row = [item.value for item in ws[1]]
    tag_row = [item.strip() for item in tag_row]

    # 获得关键数据列号
    patent_index = tag_row.index('GA') + 1  # 专利唯一名称
    company_index = tag_row.index('AE') + 1  # 所有人名称
    ref_index = tag_row.index('CP') + 1  # 引用专利
    subname_index = tag_row.index('PN') + 1  # 专利多个名称

    pattern_com = re.compile(r'\((.+?)\)')
    # 建立 所有人-唯一专利名字典 专利名-唯一专利名字典
    #      集团-所有人
    for i in range(2, num_rows):
        # 公司set
        company_str = str(ws.cell(row=i, column=company_index).value)
        company_set = set([item.strip() for item in company_str.split(';')])
        # 专利唯一名称
        patent_name = str(ws.cell(row=i, column=patent_index).value)
        # 专利子名称set
        subname_str = str(ws.cell(row=i, column=subname_index).value)
        subname_set = set([item.strip() for item in subname_str.split(';')])

        # 建立专利名-唯一专利名字典
        set_unique_name(subname_set, patent_name)

        # 建立所有人-唯一专利名字典
        add_patent_to_company(patent_name, company_set)

        # 建立集团-所有人字典
        for company in company_set:
            add_company_to_group(company, pattern_com.findall(company)[0])

        print('名称一致化处理已经完成%d/%d' % (i, num_rows))

    pattern = re.compile(r'([a-zA-Z0-9]+?-[A-Za-z]\d?)\s')
    # 建立专利-引用关系
    for j in range(2, num_rows):
        # 专利名
        patent_name = str(ws.cell(row=j, column=patent_index).value)
        # 引用专利名
        cite_set = set()
        cite_str = str(ws.cell(row=j, column=ref_index).value)
        if cite_str is None or cite_str.strip() == '':
            pass
        else:
            # 获取引用正则匹配列表
            cite_set = set(pattern.findall(cite_str)[1:])
            cite_set = get_unique_name(cite_set)

        add_patent_and_cite(patent_name, cite_set)
        add_patent_and_cited(patent_name, cite_set)

        print('引用关联已经完成%d/%d' % (i, num_rows))

    cal_centre_degree()


def pickle_save_dicts(path):
    global group_dict
    global company_dict
    global patent_dict

    print("准备保存关系字典")
    with open(path, 'wb') as save_dicts:
        pickle.dump(group_dict, save_dicts)
        pickle.dump(company_dict, save_dicts)
        pickle.dump(patent_dict, save_dicts)
    print("保存关系字典完成")


def pickle_read_dicts(path):
    global group_dict
    global company_dict
    global patent_dict

    print("准备读取关系字典")
    with open(path, 'rb') as read_dict:
        group_dict = pickle.load(read_dict)
        company_dict = pickle.load(read_dict)
        patent_dict = pickle.load(read_dict)
    print("读取关系字典完成")


def save_to_excel(path):
    print("开始创建excel表格")
    wb = Workbook()

    # 表1
    ws1 = wb.create_sheet("公司中心度")

    group_name_index = 1
    company_name_index = 2
    degree_name_index = 3
    sheet_row_index = 2

    ws1.cell(row=1, column=group_name_index, value='集团名称')
    ws1.cell(row=1, column=company_name_index, value='公司名称')
    ws1.cell(row=1, column=degree_name_index, value='中心度')

    for group_name, group_item in group_dict.items():
        for company_name in group_item.company_set:
            ws1.cell(row=sheet_row_index, column=group_name_index, value=group_name)
            ws1.cell(row=sheet_row_index, column=company_name_index, value=company_name)
            ws1.cell(row=sheet_row_index, column=degree_name_index, value=company_dict[company_name].centre_degrees)
            sheet_row_index += 1
    print("完成中心度表")

    # 表2
    ws2 = wb.create_sheet("专利引用")
    patent_index = 1
    patent_cited_index = 2
    sheet_row_index = 2

    ws2.cell(row=1, column=patent_index, value='唯一专利号')
    ws2.cell(row=1, column=patent_cited_index, value='引用专利号')

    for patent_name, patent_item in patent_dict.items():
        if len(patent_item.cite_set) == 0:
            ws2.cell(row=sheet_row_index, column=patent_index, value=patent_name)
            ws2.cell(row=sheet_row_index, column=patent_cited_index, value='')
        else:
            for item_cite in patent_item.cite_set:
                ws2.cell(row=sheet_row_index, column=patent_index, value=patent_name)
                ws2.cell(row=sheet_row_index, column=patent_cited_index, value=item_cite)
    print("完成专利引用表")

    # 表3
    ws3 = wb.create_sheet("公司引用")
    # 获得该公司各个专利的引用集合的并集 - 该公司专利 获得的差集
    for patent_name, patent_item in patent_dict.items():
        pass

    print("完成公司引用表")

    wb.save(path)
    print('excel 创建完成')


if __name__ == '__main__':
    pickle_read_dicts(pick_path)
    save_to_excel(output_path)
