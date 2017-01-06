#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd

tar_excel_path = '../output.xls'


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


"""name_ref 被引用的专利
   name_patent 引用该专利的专利"""


def addRefToPatent(name_patent, name_ref):
    global patent_dict

    # 添加引用专利对象到字典中
    if name_patent not in patent_dict:
        new_patent = PatentItem(name_patent)
        patent_dict[name_patent] = new_patent

    if name_ref == "":
        return

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
        print(owner_name, "\t\t\t中心度为:\t", owner_obj.centredegrees)


def calOwnerCentreDegree():
    excel_data = xlrd.open_workbook(tar_excel_path)
    table = excel_data.sheet_by_name(u'Sheet1')

    # 获得行数
    num_rows = table.nrows

    # 获得标签行
    row_index = table.row_values(0)
    row_index = [item.strip() for item in row_index]

    # 获得关键数据列号
    patent_index = row_index.index('GA')  # 专利唯一名称
    owner_index = row_index.index('AE')  # 所有人名称
    ref_index = row_index.index('CP')  # 引用专利
    subname_index = row_index.index('PN')  # 专利多个名称

    # 建立两个字典
    for i in range(1, num_rows):
        lineArray = table.row_values(i)

        # 建立唯一专利号字典
        namelist = lineArray[subname_index].split(';')
        nameset = set([item.strip() for item in namelist])
        get_ref_unique_name(nameset, lineArray[patent_index])

        # 建立所有人字典
        ownerlist = lineArray[owner_index].split(';')
        ownerset = set([item.strip() for item in ownerlist])
        addPatentToOwner(ownerset, lineArray[patent_index])

        # 建立专利字典 需要将被引用的专利名称改为唯一专利名
        # 同时建立引用专利的对象 和 被引用专利的对象
        addRefToPatent(lineArray[patent_index], lineArray[ref_index])

    # 获得中心度
    getCentreDegree()


if __name__ == '__main__':
    calOwnerCentreDegree()
