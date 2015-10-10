# -*- coding: utf-8 -*-

__author__ = 'yaoyuanchao'
import sys
import json
from collections import OrderedDict

import xlrd
import xlwt
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


def parse_field_data(field_type, field_data):
    if field_type in {'int', 'float'}:
        try:
            return eval(field_type)(field_data)
        except ValueError:
            return field_data
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
        field_data[order] = parse_field_data(field_type, field_data[order])

    return zip(field_labels, field_data)


def type_validator(field_type):
    if field_type in {'int'}:
        return QtGui.QIntValidator()
    if field_type in {'float'}:
        return QtGui.QDoubleValidator()
    return None


def initial_value(field_type):
    if field_type in {'int'}:
        return '0'
    if field_type in {'float'}:
        return '0.0'
    if field_type in {'bool'}:
        return 'false'
    if field_type in {'array', 'list'}:
        return '[]'
    if field_type in {'table', 'dict', 'map', 'object'}:
        return '{}'
    return ''


class DataEditor(QtGui.QWidget):
    def __init__(self):
        super(DataEditor, self).__init__()

        # left panel
        load_excel_button = QtGui.QPushButton(u"导入 表格")
        load_gdrive_button = QtGui.QPushButton(u"导入 Google文档")
        load_gdrive_button.setEnabled(False)
        load_json_button = QtGui.QPushButton(u"导入 json")
        load_area_layout = QtGui.QHBoxLayout()
        load_area_layout.addWidget(load_excel_button)
        load_area_layout.addWidget(load_gdrive_button)
        load_area_layout.addWidget(load_json_button)
        load_area_layout.setContentsMargins(0, 0, 0, 0)
        load_area_box = QtGui.QWidget()
        load_area_box.setLayout(load_area_layout)
        load_excel_button.clicked.connect(self.load_from_excel)
        load_json_button.clicked.connect(self.load_from_json)

        export_excel_button = QtGui.QPushButton(u"导出 表格")
        # export_excel_button.setEnabled(False)
        export_gdrive_button = QtGui.QPushButton(u"导出 Google文档")
        export_gdrive_button.setEnabled(False)
        export_json_button = QtGui.QPushButton(u"导出 json")
        # export_json_button.setEnabled(False)

        export_area_layout = QtGui.QHBoxLayout()
        export_area_layout.addWidget(export_excel_button)
        export_area_layout.addWidget(export_gdrive_button)
        export_area_layout.addWidget(export_json_button)
        export_area_layout.setContentsMargins(0, 0, 0, 20)
        export_area_box = QtGui.QWidget()
        export_area_box.setLayout(export_area_layout)
        export_json_button.clicked.connect(self.save_to_json)
        export_excel_button.clicked.connect(self.save_to_excel)

        self.sheet_selection = QtGui.QTabBar()
        self.sheet_selection.setUsesScrollButtons(True)
        self.sheet_names = []
        self.sheet_selection.currentChanged.connect(self.change_sheet_selection)

        self.dict_list = []
        self.data_loaded = False
        self.schema_dict = []
        self.line_editor_area_list = QtGui.QListWidget()
        self.line_editor_area_list.setAlternatingRowColors(True)
        self.line_editor_area_list.addItem(u"点击上方按钮载入数据")
        self.line_editor_area_list.currentRowChanged.connect(self.line_selected)

        insert_new_line_button = QtGui.QPushButton(u"插入")
        delete_line_button = QtGui.QPushButton(u"删除")
        copy_line_button = QtGui.QPushButton(u"复制")
        line_controller_area_layout = QtGui.QHBoxLayout()
        line_controller_area_layout.addWidget(insert_new_line_button)
        line_controller_area_layout.addWidget(delete_line_button)
        line_controller_area_layout.addWidget(copy_line_button)
        line_controller_area_layout.setContentsMargins(0, 15, 0, 15)
        line_controller_area_box = QtGui.QWidget()
        line_controller_area_box.setLayout(line_controller_area_layout)
        delete_line_button.clicked.connect(self.delete_row)
        copy_line_button.clicked.connect(self.copy_row)
        insert_new_line_button.clicked.connect(self.insert_row)

        left_panel_layout = QtGui.QVBoxLayout()
        left_panel_layout.addWidget(load_area_box)
        left_panel_layout.addWidget(export_area_box)
        left_panel_layout.addWidget(self.sheet_selection)
        left_panel_layout.addWidget(self.line_editor_area_list)
        left_panel_layout.addWidget(line_controller_area_box)
        left_panel = QtGui.QWidget()
        left_panel.setLayout(left_panel_layout)

        # right panel
        field_name_label = QtGui.QLabel(u"字段名")
        field_name_label.setAlignment(QtCore.Qt.AlignTop)
        field_type_label = QtGui.QLabel(u"类型")
        field_type_label.setAlignment(QtCore.Qt.AlignTop)
        field_value_label = QtGui.QLabel(u"值")
        field_value_label.setAlignment(QtCore.Qt.AlignTop)
        self.item_editor_area_layout = QtGui.QGridLayout()
        self.item_editor_area_layout.addWidget(field_name_label, 0, 0)
        self.item_editor_area_layout.addWidget(field_type_label, 0, 1)
        self.item_editor_area_layout.addWidget(field_value_label, 0, 2)
        self.item_editor_area_layout.setColumnMinimumWidth(0, 100)
        self.item_editor_area_layout.setColumnMinimumWidth(1, 100)
        self.item_editor_area_layout.setColumnMinimumWidth(2, 200)
        self.item_editor_area_layout.setRowMinimumHeight(0, 30)
        self.item_editor_area_layout.setRowStretch(1, 1)
        self.item_editor_area_layout.setSizeConstraint(2)
        self.item_editor_area_layout.setContentsMargins(15, 5, 15, 5)
        self.item_list = []
        item_editor_area_box = QtGui.QWidget()
        item_editor_area_box.setLayout(self.item_editor_area_layout)
        item_editor_area_scroll = QtGui.QScrollArea()
        item_editor_area_scroll.setWidgetResizable(True)
        item_editor_area_scroll.setWidget(item_editor_area_box)

        reset_item_button = QtGui.QPushButton(u"初始化")
        save_item_button = QtGui.QPushButton(u"保存")
        close_item_button = QtGui.QPushButton(u"关闭")
        item_controller_area_layout = QtGui.QHBoxLayout()
        item_controller_area_layout.addWidget(reset_item_button)
        item_controller_area_layout.addWidget(save_item_button)
        item_controller_area_layout.addWidget(close_item_button)
        item_controller_area_layout.setContentsMargins(0, 15, 0, 15)
        item_controller_area_box = QtGui.QWidget()
        item_controller_area_box.setLayout(item_controller_area_layout)

        right_panel_layout = QtGui.QVBoxLayout()
        right_panel_layout.addWidget(item_editor_area_scroll)
        right_panel_layout.addWidget(item_controller_area_box)
        self.right_panel = QtGui.QWidget()
        self.right_panel.setLayout(right_panel_layout)
        self.right_panel.setMinimumWidth(410)
        self.right_panel.hide()
        close_item_button.clicked.connect(self.right_panel.hide)
        save_item_button.clicked.connect(self.save_item_change)
        reset_item_button.clicked.connect(self.clear_item_editor_content)

        main_box_layout = QtGui.QHBoxLayout()
        main_box_splitter = QtGui.QSplitter()
        main_box_splitter.addWidget(left_panel)
        main_box_splitter.addWidget(self.right_panel)
        main_box_splitter.setStyle(QtGui.QStyleFactory.create("plastique"))
        main_box_layout.addWidget(main_box_splitter)
        main_box_layout.setContentsMargins(15, 0, 15, 0)

        # self.display_line_view()
        self.setLayout(main_box_layout)
        self.resize(1200, 700)
        self.setWindowTitle(u"数据编辑器")

    def clear_item_view(self):
        self.item_editor_area_layout.setRowStretch(self.item_editor_area_layout.rowCount() - 1, 0)
        for i in reversed(range(3, self.item_editor_area_layout.count())):
            self.item_editor_area_layout.itemAt(i).widget().setParent(None)

    def clear_tab_bar(self):
        for i in reversed(range(self.sheet_selection.count())):
            self.sheet_selection.removeTab(i)

    def change_sheet_selection(self):
        if self.data_loaded and self.sheet_selection.currentIndex() >= 0:
            self.right_panel.hide()
            self.display_line_view()

    def line_selected(self):
        if self.data_loaded and self.line_editor_area_list.currentRow() >= 0:
            self.display_item_view()

    def adjust_line_height(self):
        for i in range(self.line_editor_area_list.count()):
            item = self.line_editor_area_list.item(i)
            item.setSizeHint(QtCore.QSize(0, 22))

    def display_line_view(self):
        self.line_editor_area_list.clear()
        sheet_order = self.sheet_selection.currentIndex()
        if sheet_order < 0:
            return
        json_list = []
        for row in self.dict_list[sheet_order]:
            json_list.append(json.dumps(row, ensure_ascii=False))
        self.line_editor_area_list.addItems(json_list)
        self.adjust_line_height()

    def display_item_view(self):
        self.clear_item_view()
        self.right_panel.show()
        sheet_order = self.sheet_selection.currentIndex()
        row_num = self.line_editor_area_list.currentRow()
        item_dict = self.dict_list[sheet_order][row_num]
        self.item_list = []
        for (field, field_data) in item_dict.iteritems():
            row_count = self.item_editor_area_layout.rowCount()
            field_name = field.split('_', 1)[1]
            field_type = field.split('_', 1)[0]
            field_name_label = QtGui.QLabel(field_name)
            field_type_label = QtGui.QLabel(field_type)
            field_data_line = QtGui.QLineEdit(json.dumps(field_data, ensure_ascii=False).strip('"').strip("'"))
            field_data_line.setValidator(type_validator(field_type.lower()))
            new_item = [field_name_label, field_type_label, field_data_line]
            self.item_list.append(new_item)
            self.item_editor_area_layout.addWidget(field_name_label, row_count, 0)
            self.item_editor_area_layout.addWidget(field_type_label, row_count, 1)
            self.item_editor_area_layout.addWidget(field_data_line, row_count, 2)
        self.item_editor_area_layout.setRowStretch(self.item_editor_area_layout.rowCount(), 1)

    def load_from_excel(self):
        file_name = QtGui.QFileDialog.getOpenFileName(self, u"excel文件打开", "", "excel files (*.xls *.xlsx)")
        if file_name != "":
            self.dict_list = []
            self.sheet_names = []
            self.schema_dict = []
            self.clear_tab_bar()
            wb = xlrd.open_workbook(unicode(file_name))
            for ws in wb.sheets():
                new_dict_list = []
                field_names = ws.row_values(0)
                for row_num in range(1, ws.nrows):
                    new_row = OrderedDict(format_row_data(field_names, ws.row_values(row_num)))
                    new_dict_list.append(new_row)
                self.schema_dict.append(new_dict_list[0])
                self.dict_list.append(new_dict_list)
                self.sheet_selection.addTab(ws.name)
                self.sheet_names.append(ws.name)
            self.display_line_view()
            self.right_panel.hide()
            self.data_loaded = True

    def load_from_json(self):
        file_name = QtGui.QFileDialog.getOpenFileName(self, u"json文件打开", "", "json files (*.json)")
        if file_name != "":
            with open(unicode(file_name)) as json_file:
                self.dict_list = []
                self.sheet_names = []
                self.schema_dict = []
                self.clear_tab_bar()
                json_sheets = json.load(json_file, object_pairs_hook=OrderedDict)
                for (name, data) in json_sheets.items():
                    self.sheet_selection.addTab(name)
                    self.sheet_names.append(name)
                    self.dict_list.append(data)
                    self.schema_dict.append(data[0])
                self.display_line_view()
                self.right_panel.hide()
                self.data_loaded = True

    def save_to_json(self):
        file_name = QtGui.QFileDialog.getSaveFileName(self, u"json文件保存", "", "json files (*.json)")
        json_sheets = OrderedDict(zip(self.sheet_names, self.dict_list))
        j = json.dumps(json_sheets, ensure_ascii=False, indent=2)
        if file_name != "":
            with open(file_name, 'w') as f:
                f.write(j.encode('utf-8'))

    def save_to_excel(self):
        file_name = QtGui.QFileDialog.getSaveFileName(self, u"excel文件保存", "", "xls files (*.xls)")
        xls = xlwt.Workbook(encoding='UTF-8')
        for order in range(len(self.sheet_names)):
            sheet = xls.add_sheet(self.sheet_names[order])
            for col in range(len(self.schema_dict[order])):
                sheet.write(0, col, self.schema_dict[order].keys()[col])
            for row in range(len(self.dict_list[order])):
                for col in range(len(self.schema_dict[order])):
                    sheet.write(row + 1, col,
                                json.dumps(self.dict_list[order][row].values()[col], ensure_ascii=False).strip(
                                    '"').strip("'"))
        xls.save(file_name)

    def save_item_change(self):
        selected_row = self.line_editor_area_list.currentRow()
        sheet_order = self.sheet_selection.currentIndex()
        field_names = []
        field_values = []

        for field_order in range(len(self.dict_list[sheet_order][selected_row])):
            widget_item_row = self.item_list[field_order]
            field_names.append(str(widget_item_row[1].text() + '_' + widget_item_row[0].text()))
            field_values.append(str(widget_item_row[2].text()))
        self.dict_list[sheet_order][selected_row] = OrderedDict(format_row_data(field_names, field_values))
        self.line_editor_area_list.item(selected_row).setText(
            json.dumps(self.dict_list[sheet_order][selected_row], ensure_ascii=False))

    def clear_item_editor_content(self):
        for [_, field_type, field_data] in self.item_list:
            field_data.setText(initial_value(str(field_type.text()).lower()))

    def delete_row(self):
        row = self.line_editor_area_list.currentRow()
        sheet_order = self.sheet_selection.currentIndex()
        if row == -1 or not self.data_loaded:
            return
        row_item = self.line_editor_area_list.takeItem(row)
        del self.dict_list[sheet_order][row]
        del row_item
        if len(self.dict_list[sheet_order]) == 0:
            self.right_panel.hide()
            return
        self.line_selected()

    def copy_row(self):
        row = self.line_editor_area_list.currentRow()
        sheet_order = self.sheet_selection.currentIndex()
        if row == -1 or not self.data_loaded or sheet_order < 0:
            return
        father_item = self.line_editor_area_list.item(row)
        self.line_editor_area_list.insertItem(row + 1, father_item.text())
        self.adjust_line_height()
        father_dict = self.dict_list[sheet_order][row]
        self.dict_list[sheet_order].insert(row + 1, OrderedDict(father_dict))
        self.line_editor_area_list.setCurrentRow(row + 1)
        self.line_selected()

    def insert_row(self):
        row = self.line_editor_area_list.currentRow()
        sheet_order = self.sheet_selection.currentIndex()
        if len(self.schema_dict[sheet_order]) == 0:
            return
        self.line_editor_area_list.insertItem(row + 1, '')
        self.adjust_line_height()
        father_dict = self.schema_dict[sheet_order]
        self.dict_list[sheet_order].insert(row + 1, OrderedDict(father_dict))
        self.line_editor_area_list.setCurrentRow(row + 1)
        self.clear_item_editor_content()
        self.save_item_change()
        self.line_selected()


app = QtGui.QApplication(sys.argv)
main_window = DataEditor()
main_window.show()
sys.exit(app.exec_())
