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


def format_row_data(field_labels, field_data):
    for order in range(len(field_labels)):
        field_type = field_labels[order].split('_', 1)[0].lower()
        if field_type in {'int', 'str', 'float'}:
            field_data[order] = eval(field_type)(field_data[order])
        if field_type in {'bool'}:
            field_data[order] = eval(field_data[order].capitalize())
        if field_type in {'array', 'list'}:
            field_data[order] = eval(field_data[order].replace('{', '[').replace('}', ']'))
        if field_type in {'table', 'dict', 'map', 'object'}:
            data_string = field_data[order].strip(' {}[]()')
            data_string = data_string.replace(';', ',').replace('=', ':').replace('"', '').replace("'", '')
            if len(data_string) > 0:
                field_data[order] = OrderedDict(
                    [(k.strip(), num_or_str(v.strip())) for k, v in
                     (pair.split(':') for pair in data_string.split(','))])
            else:
                field_data[order] = OrderedDict()

    return zip(field_labels, field_data)


class DataEditor(QtGui.QWidget):
    def __init__(self):
        super(DataEditor, self).__init__()
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

        self.data_area_layout = QtGui.QVBoxLayout()
        # self.data_area_layout.addStretch(1)
        self.data_area_line_list = []
        data_area_box = QtGui.QWidget()
        data_area_box.setLayout(self.data_area_layout)
        data_area_scroll = QtGui.QScrollArea()
        data_area_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        data_area_scroll.setWidget(data_area_box)
        data_area_scroll.setWidgetResizable(True)

        export_area_layout = QtGui.QHBoxLayout()
        export_area_layout.addWidget(export_excel_button)
        export_area_layout.addWidget(export_gdrive_button)
        export_area_layout.addWidget(export_json_button)
        export_area_box = QtGui.QWidget()
        export_area_box.setLayout(export_area_layout)

        main_box = QtGui.QVBoxLayout()
        main_box.addWidget(load_area_box)
        main_box.addWidget(data_area_scroll)
        main_box.addWidget(export_area_box)

        self.display_table_view()

        self.setLayout(main_box)
        self.resize(550, 620)

    def clear_table_view(self):
        self.data_area_line_list = []
        for i in reversed(range(self.data_area_layout.count())):
            self.data_area_layout.itemAt(i).widget().deleteLater()

    def display_table_view(self):
        path = "excel_input/all_buildings.xls"
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        field_names = ws.row_values(0)

        item_list = []
        json_list = []
        for row_num in range(1, ws.nrows):
            new_row = OrderedDict(format_row_data(field_names, ws.row_values(row_num)))
            item_list.append(new_row)
            json_list.append(json.dumps(new_row))
        # j = json.dumps(item_list, indent=2)

        for json_row in json_list:
            new_line_widget = QtGui.QLabel()
            new_line_widget.setText(json_row)
            self.data_area_layout.addWidget(new_line_widget)
            self.data_area_line_list.append(new_line_widget)


app = QtGui.QApplication(sys.argv)
main_window = DataEditor()
main_window.show()
sys.exit(app.exec_())
