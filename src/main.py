#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from openpyxl import Workbook, load_workbook
import re

tar_excel_path = '../xlsx/target.xlsx'
output_path = '../xlsx/output.xlsx'


# 所有人对象 包含所有人名称 和 所有专利集合
class OwnerItem(object):
    def __init__(self, name):
        self.name = name
        self.patentset = set()
        self.centredegrees = 0

    def addToPatentset(self, item):
        self.patentset.add(item)


# 专利对象 包含专利名称 和 引用专利集合
class PatentItem(object):
    def __init__(self, name):
        self.name = name
        self.refset = set()

    def addToRefset(self, item):
        self.refset.add(item)


owner_dict = {}  # 所有人字典
patent_dict = {}  # 专利字典
unique_name_dict = {}  # 专利唯一名称字典

# 1.逐个读取各个专利 获取专利名称(唯一） 和 专利所有人
# 2.查询当前专利所有人是否已经在所有人字典中存在 否，则创建并添加 是,则添加专利名称到其专利集合中
# 3.遍历完成，获得所有人字典 记载了全部所有人，及其所拥有的专利
# 4.逐个读取各个专利 获取专利名称(唯一） 和 引用专利(转化为唯一专利）
# 5.查询引用专利是否在专利字典中存在 否，则创建并添加 是,则添加专利名称到引用专利的被引用专利集合中
# 6.遍历完成，获得专利被引用字典 记载了全部被引用专利，及其所引用该专利的专利集合
# 7.逐个遍历所有人字典，对单个所有人的全部专利的被引用集合作并集处理，获得对此所有人的引用
# 8.获得该并集与有该所有人拥有的专利集合的差集(集合相减)，获得有效的专利引用集合，获得中心度

"""name_patent 专利名
   name_owner  专利所有人名"""


def addPatentToOwner(ownerset, name_patent):
    global owner_dict
    for owner_name in ownerset:
        if owner_name not in owner_dict:
            new_owner = OwnerItem(owner_name)
            owner_dict[owner_name] = new_owner
        owner_dict[owner_name].addToPatentset(name_patent)


"""name_patent 引用该专利的专利"""


def addpatenttodict(name_patent):
    global patent_dict

    # 添加引用专利对象到字典中
    if name_patent not in patent_dict:
        new_patent = PatentItem(name_patent)
        patent_dict[name_patent] = new_patent


"""name_ref 被引用的专利
   name_patent 引用该专利的专利"""


def addRefToPatent(name_patent, ref_set):
    global patent_dict

    for name_ref in ref_set:
        # 先将被引用专利号转为唯一专利号
        if name_ref in unique_name_dict:
            unique_id = unique_name_dict[name_ref]
        else:
            # 未找到对应的唯一专利号
            unique_id = name_ref

        # 添加被引用专利对象到字典中 并增加其被引用属性
        if unique_id not in patent_dict:
            new_patent = PatentItem(unique_id)
            patent_dict[unique_id] = new_patent

        patent_dict[unique_id].addToRefset(name_patent)


def get_ref_unique_name(nameset, unique_name):
    global unique_name_dict
    for name_item in nameset:
        if name_item not in unique_name_dict:
            unique_name_dict[name_item] = unique_name


def getCentreDegree():
    global owner_dict
    global patent_dict

    wb = Workbook()
    ws = wb.active
    row_index = 1
    # 遍历所有人
    for owner_name, owner_obj in owner_dict.items():
        validset = set()
        # 遍历该所有人的全部专利
        for patent in owner_obj.patentset:
            # 且该专利在专利字典中存在 即有引用过其他专利
            if patent in patent_dict:
                validset = validset | patent_dict[patent].refset
        validset = validset - owner_obj.patentset
        # 获得中心度
        owner_obj.centredegrees = len(validset)
        ws.cell(row=row_index, column=1, value=owner_name)
        ws.cell(row=row_index, column=2, value=owner_obj.centredegrees)
        row_index += 1
        print("所有人: %-60s 中心度为: %-5s" % (owner_name, owner_obj.centredegrees))
    wb.save(output_path)


def calOwnerCentreDegree():
    print('正在读取excel文件.................')
    wb = load_workbook(tar_excel_path)
    sheetnames = wb.get_sheet_names()
    ws = wb.get_sheet_by_name(sheetnames[0])
    print('读取完成,开始数据处理.............')

    # 获得行数
    num_rows = ws.max_row

    # 获得标签行
    tag_row = [item.value for item in ws[1]]
    tag_row = [item.strip() for item in tag_row]

    # 获得关键数据列号
    patent_index = tag_row.index('GA') + 1  # 专利唯一名称
    owner_index = tag_row.index('AE') + 1  # 所有人名称
    ref_index = tag_row.index('CP') + 1  # 引用专利
    subname_index = tag_row.index('PN') + 1  # 专利多个名称

    pattern = re.compile(r'([a-zA-Z0-9]+?-[A-Za-z]\d?)\s')

    # 建立两个字典 所有人-唯一专利名字典 专利名-唯一专利名字典
    for i in range(2, num_rows):
        # 所有人set
        owner_str = str(ws.cell(row=i, column=owner_index).value)
        owner_set = set([item.strip() for item in owner_str.split(';')])
        # 专利唯一名称
        patent_name = str(ws.cell(row=i, column=patent_index).value)
        # 专利子名称set
        subname_str = str(ws.cell(row=i, column=subname_index).value)
        subname_set = set([item.strip() for item in subname_str.split(';')])

        # 建立专利名-唯一专利名字典
        get_ref_unique_name(subname_set, patent_name)

        # 建立所有人字典
        addPatentToOwner(owner_set, patent_name)
        print('名称一致化处理已经完成%d/%d' % (i, num_rows))

    # 建立被引用-引用字典
    for j in range(2, num_rows):
        # 建立专利对象
        patent_name = str(ws.cell(row=j, column=patent_index).value)
        addpatenttodict(patent_name)

        # 建立专利引用字典 需要将被引用的专利名称改为唯一专利名
        ref_str = str(ws.cell(row=j, column=ref_index).value)

        if ref_str is None or ref_str.strip() == '':
            pass
        else:
            # 获取引用正则匹配列表
            ref_list = pattern.findall(ref_str)
            # 不包含第一个对自己的引用
            addRefToPatent(patent_name, set(ref_list[1:]))

        print('引用关联已经完成%d/%d' % (i, num_rows))

    print('开始计算中心度:')
    # 获得中心度
    getCentreDegree()


if __name__ == '__main__':
    calOwnerCentreDegree()
