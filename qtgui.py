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
        load_area_layout = QtGui.QHBoxLayout()
        load_area_layout.addWidget(load_excel_button)
        load_area_layout.addWidget(load_gdrive_button)
        load_area_layout.addWidget(load_json_button)
        load_area_box = QtGui.QWidget()
        load_area_box.setLayout(load_area_layout)
        load_excel_button.clicked.connect(self.load_from_excel)

        self.line_editor_area_list = QtGui.QListWidget()
        self.dict_list = []
        self.line_editor_area_list.currentRowChanged.connect(self.display_item_view)

        insert_new_line_button=QtGui.QPushButton(u"插入")
        delete_line_button=QtGui.QPushButton(u"删除")
        copy_line_button=QtGui.QPushButton(u"复制")
        line_controller_area_layout=QtGui.QHBoxLayout()
        line_controller_area_layout.addWidget(insert_new_line_button)
        line_controller_area_layout.addWidget(delete_line_button)
        line_controller_area_layout.addWidget(copy_line_button)
        line_controller_area_box=QtGui.QWidget()
        line_controller_area_box.setLayout(line_controller_area_layout)

        export_excel_button = QtGui.QPushButton(u"导出 表格")
        export_gdrive_button = QtGui.QPushButton(u"导出 Google文档")
        export_json_button = QtGui.QPushButton(u"导出 json")
        export_area_layout = QtGui.QHBoxLayout()
        export_area_layout.addWidget(export_excel_button)
        export_area_layout.addWidget(export_gdrive_button)
        export_area_layout.addWidget(export_json_button)
        export_area_box = QtGui.QWidget()
        export_area_box.setLayout(export_area_layout)
        export_json_button.clicked.connect(self.save_to_json)

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
        self.item_editor_area_layout.setColumnMinimumWidth(0,100)
        self.item_editor_area_layout.setColumnMinimumWidth(1,100)
        self.item_editor_area_layout.setColumnMinimumWidth(2,200)
        self.item_editor_area_layout.setRowMinimumHeight(0,40)
        self.item_editor_area_layout.setRowStretch(1,1)
        self.item_list=[]
        item_editor_area_box=QtGui.QWidget()
        item_editor_area_box.setLayout(self.item_editor_area_layout)
        item_editor_area_scroll=QtGui.QScrollArea()
        item_editor_area_scroll.setWidgetResizable(True)
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
        self.right_panel=QtGui.QWidget()
        self.right_panel.setLayout(right_panel_layout)
        self.right_panel.hide()
        close_item_button.clicked.connect(self.right_panel.hide)
        save_item_button.clicked.connect(self.save_item_change)

        main_box_layout=QtGui.QHBoxLayout()
        main_box_layout.addWidget(left_panel)
        main_box_layout.addWidget(self.right_panel)
        main_box_layout.setContentsMargins(15,0,15,0)

        self.display_line_view()
        self.setLayout(main_box_layout)
        self.resize(1200, 700)

    def clear_item_view(self):
        #for row in range(self.item_editor_area_layout.rowCount()):
        #    self.item_editor_area_layout.setRowStretch(row,0)
        self.item_editor_area_layout.setRowStretch(self.item_editor_area_layout.rowCount()-1,0)
        for i in reversed(range(3,self.item_editor_area_layout.count())):
            self.item_editor_area_layout.itemAt(i).widget().setParent(None)

    def display_line_view(self):
        self.line_editor_area_list.clear()
        json_list=[]
        for row in self.dict_list:
            json_list.append(json.dumps(row))
        self.line_editor_area_list.addItems(json_list)

    def display_item_view(self,row_num):
        self.clear_item_view()
        self.right_panel.show()
        item_dict=self.dict_list[row_num]
        self.item_list=[]
        for (field,field_data) in item_dict.iteritems():
            row_count=self.item_editor_area_layout.rowCount()
            field_name=field.split('_',1)[1]
            field_type=field.split('_',1)[0]
            field_name_label=QtGui.QLabel(field_name)
            field_type_label=QtGui.QLabel(field_type)
            field_data_line=QtGui.QLineEdit(json.dumps(field_data).strip('"').strip("'"))
            field_data_line.setValidator(self.type_validator(field_type.lower()))
            new_item=[field_name_label,field_type_label,field_data_line]
            self.item_list.append(new_item)
            self.item_editor_area_layout.addWidget(field_name_label,row_count,0)
            self.item_editor_area_layout.addWidget(field_type_label,row_count,1)
            self.item_editor_area_layout.addWidget(field_data_line,row_count,2)
        self.item_editor_area_layout.setRowStretch(self.item_editor_area_layout.rowCount(),1)

    def line_selected(self,line_num):
        self.display_item_view(line_num)

    def save_to_json(self):
        j=json.dumps(self.dict_list,indent=2)
        file_name=QtGui.QFileDialog.getSaveFileName(self,u"json文件保存",QtCore.QDir.currentPath()+"/json_output/"+".json","json files (*.json)")
        if file_name<>"":
            with open(file_name, 'w') as f:
                f.write(j)

    def load_from_excel(self):
        file_name=QtGui.QFileDialog.getOpenFileName(self,u"excel文件打开",QtCore.QDir.currentPath()+"/excel_input/","excel files (*.xls)")
        if file_name<>"":
            self.dict_list = []
            wb = xlrd.open_workbook(file_name)
            ws = wb.sheet_by_index(0)
            field_names = ws.row_values(0)
            for row_num in range(1, ws.nrows):
                new_row = OrderedDict(format_row_data(field_names, ws.row_values(row_num)))
                self.dict_list.append(new_row)
            self.display_line_view()
            self.right_panel.hide()

    def save_item_change(self):
        selected_row=self.line_editor_area_list.currentRow()
        field_names=[]
        field_values=[]

        for field_order in range(len(self.dict_list[selected_row])):
            widget_item_row=self.item_list[field_order]
            field_names.append(str(widget_item_row[1].text()+'_'+widget_item_row[0].text()))
            field_values.append(str(widget_item_row[2].text()))
        self.dict_list[selected_row]=OrderedDict(format_row_data(field_names,field_values))
        self.line_editor_area_list.item(selected_row).setText(json.dumps(self.dict_list[selected_row]))

    def type_validator(self,type):
        if type in {'int'}:
            return QtGui.QIntValidator()
        if type in {'float'}:
            return QtGui.QDoubleValidator()
        return None

app = QtGui.QApplication(sys.argv)
main_window = DataEditor()
main_window.show()
sys.exit(app.exec_())
