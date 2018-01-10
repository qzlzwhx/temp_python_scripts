# coding:utf8

import xlrd
import json
import os

subject_dict = {
    "历史": "history",
    "地理": "geography",
    "语文": "chinese",
    "政治": "politics",
}

type_tag = {
    u"知识": "knowledge",
    u"方法": "method",
    u"载体": "carrier",
    u"专题": "topic",
    u"模型": "topic",
    u"形式": "form",
    u"能力": "ability",
    u"评级": "rating",
    u"题型": "subtype",
    u"表现": "behavior",

}


type_info = dict(edu='junior', subject="geography", _type="knowledge", number=0)


class Node(object):

    def __init__(self, name=None, parent=None, children=None, _type="knowledge", edu="senior", subject="chinese"):
        self.name = name
        self.parent = parent
        self.children = children

        self.start_row_number = 0
        self.end_row_number = 0
        self.edu = edu
        self._type = _type
        self.subject = subject
        self._id = 0

    # dict(type="knowledge",
    #      edu="",
    #      subject="")
    def to_json(self):
        children = []
        # if self.children:
        #     for child in self.children:
        #         children.append(child.to_json())
        if not children: children = ""

        return {"name": self.name, "parent": self.parent._id if self.parent else "",
                "edu": self.edu, "subject": self.subject, "type": self._type, "id": self._id}


def create_node_list(start_row, end_row, colx, sheet):
    last_name = None
    node_list = []
    last_node = None
    for i in range(start_row, end_row):
        cell = sheet.cell(rowx=i, colx=colx)

        if cell.value and cell.value != last_name:
            last_node = Node()
            last_node.name = cell.value
            last_node.start_row_number = i + 1
            last_name = cell.value
            # TODO sheet name panduan
            auto_increment_number = type_info.get("number")
            last_node.edu = type_info.get("edu") # senior
            last_node.subject = type_info.get('subject')
            last_node._type = type_info.get('_type')
            last_node._id = auto_increment_number
            type_info["number"] += 1
            node_list.append(last_node)
        elif last_node:
            # print i
            last_node.end_row_number = i + 1

    return node_list


def get_children_nodes(parent_node, colx, sheet):
    start_row = parent_node.start_row_number - 1
    end_row = parent_node.end_row_number
    # print start_row, end_row, '--------------------------'
    node_list = create_node_list(start_row, end_row, colx, sheet)

    #
    # for node in node_list:
    #     print 'child:', colx, node.name, node.start_row_number, node.end_row_number
    return node_list


def create_json_by_file(sheets):
    result = []
    for sheet in sheets:

        if sheet.nrows == 0:
            continue
        # TODO sheet name
        name_key = sheet.name.strip()[:2]

        _type = type_tag.get(name_key)
        if not _type:
            print '----------------------', sheet.name
            continue
        print '======', name_key, _type
        type_info["_type"] = _type
        node_list = create_node_list(0, sheet.nrows, 0, sheet)
        rs = []

        for node in node_list:
            # print node.name, node.start_row_number, node.end_row_number
            node.children = get_children_nodes(node, 1, sheet)
            rs.append(node)
            print node.name
            # second level
            for child in node.children:
                rs.append(child)
                child.parent = node
                child.children = get_children_nodes(child, 2, sheet)

                # third level
                for grandson in child.children:
                    rs.append(grandson)
                    grandson.parent = child
                    grandson.children = get_children_nodes(grandson, 3, sheet)
                    rs += grandson.children
                    for g_g_son in grandson.children:
                        g_g_son.parent = grandson
            # rs.append(node.to_json())

            # rs.update({node.name: node.to_json().get("children")})
        # result.append(json.dumps(rs, encoding="UTF-8", ensure_ascii=False))
        result.append(rs)
    return result


def create_json_data(type_info):
    pathDir = os.listdir("./")
    whole_json = []
    for file in pathDir:
        edu = 'junior'
        if "高中" in file:
            edu = 'senior'  # senior

        subject = subject_dict.get(file.split('.')[0][6:])
        _type = "knowledge"
        type_info['edu'] = edu
        type_info['subject'] = subject
        type_info['_type'] = _type

        if "xls" in file:
            workbook = xlrd.open_workbook(file)
            # print file.split('.')[0][6:]
            # print file
            sheets = workbook.sheets()
            # 去掉0了，要修改判断的东西了
            k_json = create_json_by_file(sheets)
            for k in k_json:
                whole_json += k
            # break
    # print json.dumps(whole_json, encoding="UTF-8", ensure_ascii=False)
    print json.dumps([whole.to_json() for whole in whole_json], encoding="UTF-8", ensure_ascii=False)
    # f = open("output_json.txt", 'w')
    # f.write(json.dumps(whole_json))
    # f.close()


create_json_data(type_info)

workbook = xlrd.open_workbook("初中语文.xlsx")
for s in workbook.sheets():
    print s.name[:2], s.nrows
