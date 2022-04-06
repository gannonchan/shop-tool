import openpyxl
import decimal
import datetime
import time
from sheetvo import SheetVo


def parse_worksheet(src_workbook_path, date):
    workbook = openpyxl.load_workbook(src_workbook_path)
    worksheet_name = 'Sheet'+date
    worksheet = workbook[worksheet_name]
    row_index = 0
    col_dict = {}
    for cIndex in range(0, 26):
        col_dict.setdefault(cIndex + 1, chr(cIndex + 65))
    title_attr_dict_cn_keys = {}
    title_attr_dict_index_keys = {}
    price_attr_dict = {}
    for row in worksheet.rows:
        row_index += 1
        cell_index = 0
        for cell in row:
            cell_index += 1
            value = cell.value
            # print(cell.ctype)
            if not value is None:
                if isinstance(value, str):
                    index = col_dict.get(cell_index) + str(row_index)
                    title_attr_dict_cn_keys.setdefault(value, index)
                    title_attr_dict_index_keys.setdefault(index, value)
                if isinstance(value, float):
                    index = col_dict.get(cell_index) + str(row_index)
                    price_attr_dict.setdefault(value, index)
    final_dict = get_title_index(price_attr_dict, title_attr_dict_index_keys, title_attr_dict_cn_keys)
    shop_name = worksheet['B2'].value
    vo = SheetVo(final_dict,date,shop_name)
    return vo


# 获取金额头文本下标
def get_title_index(price_attr_dict, title_attr_dict_index_keys, title_attr_dict_cn_keys):
    final_dict = {}
    for key in price_attr_dict.keys():
        cell_index = price_attr_dict.get(key)
        col_index = cell_index[0]
        row_index = cell_index[1:]
        col_index_ascii = ord(col_index)
        title_index = chr(col_index_ascii - 1) + row_index
        title_text = title_attr_dict_index_keys.get(title_index)
        final_dict.setdefault(title_text, key)
    return final_dict


def padding_sheet(target_workbook_path, src_sheet_vo):
    target_workbook = openpyxl.load_workbook(target_workbook_path)
    shop_name = src_sheet_vo.getshopName()
    target_worksheet = target_workbook[shop_name]
    day_str = src_sheet_vo.getday()
    day_column = target_worksheet['A']
    exec_row_index = 0
    for cell in day_column:
        exec_row_index += 1
        if cell.value == day_str:
            print(cell.value)
            break
    for column in target_worksheet['B':'AA']:
        # print(column)
        row_index = 0
        value = 0
        flag = False
        for cell in column:
            row_index += 1
            if row_index == 2:
                #titleRowIndex
                cell_value = cell.value
                print(cell_value)
                dict = src_sheet_vo.getdict()
                if not flag:
                    value = dict.get(cell_value)
                    if value is None:
                        break
                    else:
                        flag = True
                        print(value)
            if row_index == exec_row_index:
                cell.value = value
                print(value)

