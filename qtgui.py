# -*- coding: utf-8 -*-

__author__ = 'yaoyuanchao'
import sys
import json
from collections import OrderedDict

import xlrd
from PyQt4 import QtGui
from PyQt4 import QtCore


def is_excel_file(file_path):
    return file_path.rsplit('.', 1)[1] in {'xls', 'xlxs'}


def num_or_str(value):
    try:
        float_v = float(value)
        try:
            int_v = int(value)
            return int_v
        except ValueError:
            return float_v
    except ValueError:
        return value

def parse_field_data(field_type,field_data):
    if field_type in {'int', 'str', 'float'}:
        return eval(field_type)(field_data)
    if field_type in {'bool'}:
        return eval(field_data.capitalize())
    if field_type in {'array', 'list'}:
        return eval(field_data.replace('{', '[').replace('}', ']'))
    if field_type in {'table', 'dict', 'map', 'object'}:
        data_string = field_data.strip(' {}[]()')
        data_string = data_string.replace(';', ',').replace('=', ':').replace('"', '').replace("'", '')
        if len(data_string) > 0:
            return OrderedDict(
                [(k.strip(), num_or_str(v.strip())) for k, v in
                 (pair.split(':') for pair in data_string.split(','))])
        else:
            return OrderedDict()
    return field_data

def format_row_data(field_labels, field_data):
    for order in range(len(field_labels)):
        field_type = field_labels[order].split('_', 1)[0].lower()
        field_data[order]=parse_field_data(field_type,field_data[order])

    return zip(field_labels, field_data)


class DataEditor(QtGui.QWidget):
    def __init__(self):
        super(DataEditor, self).__init__()

        # left panel
        load_excel_button = QtGui.QPushButton(u"导入 表格")
        load_gdrive_button = QtGui.QPushButton(u"导入 Google文档")
        load_json_button = QtGui.QPushButton(u"导入 json")
        export_excel_button = QtGui.QPushButton(u"导出 表格")
        export_gdrive_button = QtGui.QPushButton(u"导出 Google文档")
        export_json_button = QtGui.QPushButton(u"导出 json")
        load_area_layout = QtGui.QHBoxLayout()
        load_area_layout.addWidget(load_excel_button)
        load_area_layout.addWidget(load_gdrive_button)
        load_area_layout.addWidget(load_json_button)
        load_area_box = QtGui.QWidget()
        load_area_box.setLayout(load_area_layout)

        self.line_editor_area_list = QtGui.QListWidget()
        self.dict_list = []
        self.json_list = []

        insert_new_line_button=QtGui.QPushButton(u"插入")
        delete_line_button=QtGui.QPushButton(u"删除")
        copy_line_button=QtGui.QPushButton(u"复制")
        line_controller_area_layout=QtGui.QHBoxLayout()
        line_controller_area_layout.addWidget(insert_new_line_button)
        line_controller_area_layout.addWidget(delete_line_button)
        line_controller_area_layout.addWidget(copy_line_button)
        line_controller_area_box=QtGui.QWidget()
        line_controller_area_box.setLayout(line_controller_area_layout)

        export_area_layout = QtGui.QHBoxLayout()
        export_area_layout.addWidget(export_excel_button)
        export_area_layout.addWidget(export_gdrive_button)
        export_area_layout.addWidget(export_json_button)
        export_area_box = QtGui.QWidget()
        export_area_box.setLayout(export_area_layout)

        left_panel_layout = QtGui.QVBoxLayout()
        left_panel_layout.addWidget(load_area_box)
        left_panel_layout.addWidget(self.line_editor_area_list)
        left_panel_layout.addWidget(line_controller_area_box)
        left_panel_layout.addWidget(export_area_box)
        left_panel=QtGui.QWidget()
        left_panel.setLayout(left_panel_layout)

        #right panel
        field_name_label=QtGui.QLabel(u"字段名")
        field_type_label=QtGui.QLabel(u"类型")
        field_value_label=QtGui.QLabel(u"值")
        self.item_editor_area_layout=QtGui.QGridLayout()
        self.item_editor_area_layout.addWidget(field_name_label,0,0)
        self.item_editor_area_layout.addWidget(field_type_label,0,1)
        self.item_editor_area_layout.addWidget(field_value_label,0,2)
        self.item_editor_area_layout.setColumnMinimumWidth(0,80)
        self.item_editor_area_layout.setColumnMinimumWidth(1,60)
        self.item_editor_area_layout.setRowMinimumHeight(0,40)
        self.item_list=[]
        item_editor_area_box=QtGui.QWidget()
        item_editor_area_box.setLayout(self.item_editor_area_layout)
        item_editor_area_scroll=QtGui.QScrollArea()
        item_editor_area_scroll.setWidget(item_editor_area_box)

        reset_item_button=QtGui.QPushButton(u"初始化")
        save_item_button=QtGui.QPushButton(u"保存")
        close_item_button=QtGui.QPushButton(u"关闭")
        item_controller_area_layout=QtGui.QHBoxLayout()
        item_controller_area_layout.addWidget(reset_item_button)
        item_controller_area_layout.addWidget(save_item_button)
        item_controller_area_layout.addWidget(close_item_button)
        item_controller_area_box=QtGui.QWidget()
        item_controller_area_box.setLayout(item_controller_area_layout)

        right_panel_layout=QtGui.QVBoxLayout()
        right_panel_layout.addWidget(item_editor_area_scroll)
        right_panel_layout.addWidget(item_controller_area_box)
        right_panel=QtGui.QWidget()
        right_panel.setLayout(right_panel_layout)

        main_box_layout=QtGui.QHBoxLayout()
        main_box_layout.addWidget(left_panel)
        main_box_layout.addWidget(right_panel)
        main_box_layout.setContentsMargins(15,0,15,0)

        self.display_line_view()
        self.setLayout(main_box_layout)
        #self.setContentsMargins(0,0,0,0)
        self.resize(550, 620)

    def clear_table_view(self,table):
        #self.data_area_line_list = []
        for i in reversed(range(table.count())):
            table.itemAt(i).widget().deleteLater()

    def display_line_view(self):
        path = "excel_input/all_buildings.xls"
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        field_names = ws.row_values(0)

        self.dict_list = []
        self.json_list = []
        for row_num in range(1, ws.nrows):
            new_row = OrderedDict(format_row_data(field_names, ws.row_values(row_num)))
            self.dict_list.append(new_row)
            self.json_list.append(json.dumps(new_row))
        # j = json.dumps(dict_list, indent=2)

        self.line_editor_area_list.addItems(self.json_list)

    def display_item_view(self):
        item_dict=self.dict_list[0]
        self.item_list=[]
        for (field,field_data) in item_dict.iteritems():
            row_count=self.item_editor_area_layout.rowCount()
            field_name=field.split('_')[1]
            field_type=field.split('_')[0].lower()
            new_item=[QtGui.QLabel(field_name),QtGui.QLabel(field_type),QtGui.QLineEdit(field_data)]
            self.item_list.append(new_item)
            self.item_editor_area_layout.addWidget(new_item[0],row_count,0)
            self.item_editor_area_layout.addWidget(new_item[1],row_count,1)
            self.item_editor_area_layout.addWidget(new_item[2],row_count,2)




app = QtGui.QApplication(sys.argv)
main_window = DataEditor()
main_window.show()
sys.exit(app.exec_())
