__author__ = 'yaoyuanchao'

from collections import OrderedDict
import json

import xlrd


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


filename = "all_buildings.xls"
wb = xlrd.open_workbook(filename)

for ws in wb.sheets():

    field_names = ws.row_values(0)

    item_list = []
    for row_num in range(1, ws.nrows):
        new_row = OrderedDict(format_row_data(field_names, ws.row_values(row_num)))
        item_list.append(new_row)

    j = json.dumps(item_list, indent=2)
    with open(filename.split('.')[0] + '_' + ws.name + '.json', 'w') as f:
        f.write(j)
