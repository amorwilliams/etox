# -*- coding: utf-8 -*-

import xlrd
import math
import json
from slpp import slpp as lua

class SheetManager(object):
    """docstring for SheetManager"""
    def __init__(self):
        self.sheet_dic = {}
        self.sheet_name_list = []


    def add_work_book(self, filepath):
        wb = xlrd.open_workbook(filepath)

        for sheet_index in range(wb.nsheets):
            sh = wb.sheet_by_index(sheet_index)
            sheet = Sheet(self, sh)
            self.add_sheet(sheet)


    def add_sheet(self, sheet):
        self.sheet_dic[sheet.name] = sheet
        self.sheet_name_list.append(sheet.name)


    def get_sheet(self, name):
        if name in self.sheet_dic:
            return self.sheet_dic[name]
        return None


    def get_sheet_name_list(self):
        return self.sheet_name_list


    def export_json(self, name, sheet_output_field = []):
        sheet = self.sheet_dic[name]
        data = sheet.to_python(sheet_output_field)
        return json.dumps(data, sort_keys=True, indent=2, ensure_ascii=False)

    def export_lua(self, name, sheet_output_field=[]):
        sheet = self.sheet_dic[name]
        data = sheet.to_python(sheet_output_field)
        s = ('%s' % '\n').join([ 't[%s] = %s' % (k, lua.encode(v)) for k, v in data.items() ])
        return 't={}\n %s \n ' % s

    def is_ref_sheet(self, name):
        for sheet_name in self.sheet_dic:
            if name in self.sheet_dic[sheet_name].ref_sheets:
                return  True

        return False




class Field(object):
    """Field defined of sheet"""
    def __init__(self):
        self.name = None
        self.type = None
        self.default = None
        self.desc = None


    def __str__(self):
        return "name:%r,type:%r,defautl:%r,desc:%r" % (self.name, self.type, self.default, self.desc)




class Sheet:
    def __init__(self, mgr, sh):
        self.mgr = mgr
        self.sh = sh
        self.name = sh.name
        self.initialized = False
        self.field_list = []
        self.ref_sheets = set()
        self.p_data = {}

        self.__find_row()
        self.__find_col()

        self.__parse_field()
        self.__parse_ref_sheet()

        self.__convert_to_python_data()


    def __find_row(self):

        '''
        查找数据起始行，名称行，描述行，缺省值行，类型行，数据终止行数
        '''
        self.defualt_row = -1
        self.desc_row = -1

        for i in range(0, 5):
            value = self.sh.cell(i,0).value
            if value == '__default__':
                self.defualt_row = i
            elif value == '__type__':
                self.type_row = i
            elif value == '__name__':
                self.name_row = i
            elif value == '__desc__':
                self.desc_row = i
            else:
                self.data_begin_row = i
                break

        for row in range(self.sh.nrows):
            if self.sh.cell(row, 0).ctype == xlrd.XL_CELL_EMPTY:
                self.data_end_row = row
                break

        if row == self.sh.nrows - 1:
            self.data_end_row = self.sh.nrows


    def __find_col(self):
        '''
        查找数据终止列数
        '''

        for col in range(self.sh.ncols):
            if self.sh.cell(self.name_row, col).ctype == xlrd.XL_CELL_EMPTY:
                self.data_end_col = col
                break

        if col == self.sh.ncols - 1:
            self.data_end_col = self.sh.ncols


    def __parse_field(self):
        '''
        解析字段属性
        '''

        for col in range(self.data_end_col):
            field = Field()
            self.field_list.append(field)

            field.type = self.sh.cell(self.type_row, col).value
            field.name = self.sh.cell(self.name_row, col).value

            if self.defualt_row == -1:
                field.desc = None
            else:
                field.desc = self.sh.cell(self.desc_row, col).value

            if self.defualt_row == -1:
                field.default = None
            else:
                type = field.type
                ctype = self.sh.cell(self.defualt_row, col).ctype
                value = self.sh.cell(self.defualt_row, col).value

                if col == 0:
                    field.default = None #第一位缺省值，占位符
                elif ctype == xlrd.XL_CELL_EMPTY: #空白格
                    field.default = None
                elif value == 'null':  #null格
                    field.default = None
                elif type == 'int':
                    field.default = int(value)
                elif type == 'float':
                    field.default = float(value)
                elif type == 'string':
                    field.default = value
                elif type == 'boolean':
                    field.default = bool(value)
                elif type == 'object':
                    field.default = self.__convert_str_to_dic(value)
                elif type == 'int[]' or type == 'float[]' or type == 'string[]' or type == 'object[]':  #数组
                    field.default = self.__convert_str_to_list(value, type)
                elif type == 'ref':  #引用
                    field.default = value


    def __parse_ref_sheet(self):
        for row in range(self.data_begin_row, self.data_end_row):
            for col in range(1, self.data_end_col):
                field = self.field_list[col]
                field_type = field.type
                value = self.sh.cell(row, col).value

                if field_type == 'ref':
                    sheet_name = value.split('.')[0]
                    self.ref_sheets.add(sheet_name)


    def __convert_str_to_list(self, str, type):
        '''
        转换字符串为list
        '''
        type = type.split('[')[0]
        list = str.split(',')

        for i in range(len(list)):
            if type == 'int':
                list[i] = int(list[i])
            elif type == 'float':
                list[i] = float(list[i])
            elif type == 'string':
                list[i] = list[i]
            elif type == 'object':
                list[i] = self.__convert_str_to_dic(list[i])

        return list


    def __convert_str_to_dic(self, str):
        '''
        转换字符串为
        '''
        dict = {}
        list = str.split(',')

        for i in range(len(list)):
            kv = list[i].split(':')
            key = kv[0]
            value = kv[1]

            if value.isdigit() and '.' in value:
                dict[key] = float(value)
            elif value.isdigit():
                dict[key] = int(value)
            else:
                dict[key] = value

        return dict


    def log(self):
        print 'Default line:', self.defualt_row + 1
        print 'Type line:', self.type_row + 1
        print 'Name line:', self.name_row + 1
        print 'Description line:', self.desc_row + 1
        print 'Data begin line:', self.data_begin_row + 1
        print 'Data end line:', self.data_end_row + 1
        print 'Data end column', self.data_end_col + 1
        print 'Field properties:'
        for field in self.field_list:
            print field
        print 'Refrence sheet:', self.ref_sheets


    def __convert_to_python_data(self):
        '''
        将数据解析成python格式。不包括引用数据。
        '''
        
        #dump data
        for row in range(self.data_begin_row, self.data_end_row):
            recordId = self.__get_record_id(row)
            record = self.p_data[recordId] = {}

            for col in range(1, self.data_end_col):
                field = self.field_list[col]

                field_name = field.name
                filed_type = field.type

                value = self.sh.cell(row, col).value
                ctype = self.sh.cell(row, col).ctype

                if ctype == xlrd.XL_CELL_EMPTY:
                    record[field_name] = field.default
                elif value == 'null':
                    record[field_name] = None
                else:
                    #如果没有类型字段,自动判断类型，只支持int,float,string
                    if filed_type == '' or filed_type == None:
                        filed_type = self.__auto_decide_type(value)

                    if filed_type == 'int':
                        record[field_name] = int(value)
                    elif filed_type == 'float':
                        record[field_name] = float(value)
                    elif filed_type == 'string':
                        record[field_name] = value
                    elif filed_type == 'boolean':
                        record[field_name] = bool(value)
                    elif filed_type == 'object':
                        record[field_name] = self.__convert_str_to_dic(value)
                    elif filed_type == 'int[]' or filed_type == 'float[]' or filed_type == 'string[]' or filed_type == 'object[]':  #数组
                        record[field_name]= self.__convert_str_to_list(value, filed_type)
                    elif filed_type == 'ref':  #引用
                        record[field_name] = value


    def __get_record_id(self, row):
        '''
        获得当前行的recordId
        '''
        recordId = self.sh.cell(row, 0).value
        ctype = self.sh.cell(row, 0).ctype
        if ctype == xlrd.XL_CELL_TEXT:
            pass
        elif ctype == xlrd.XL_CELL_NUMBER:
            recordId = int(recordId)
            #TODO 不支持浮点数做主键

        return recordId


    def __auto_decide_type(self, value):
        if isinstance(value, float):
            if math.ceil(value) == value:
                return 'int'
            else:
                return 'float'
        else:
            return 'string'
    

    def to_python(self, sheet_output_field=[]):
        '''
        插入引用表
        '''
        if not self.initialized:
            self.__merge()
            self.initialized = True

        if sheet_output_field == []:
            return self.p_data
        else:
            new_p_data = self.p_data.copy()
            for recordId in new_p_data:
                del_field_name_list = []

                for field_name in new_p_data[recordId]:
                    if field_name in sheet_output_field:
                        pass
                    else:
                        del_field_name_list.append(field_name)

                for del_field_name in del_field_name_list:
                    del new_p_data[recordId][del_field_name]

        return new_p_data


    def __merge(self):
        '''
        合并引用表到当前表
        '''
        for row in range(self.data_begin_row, self.data_end_row):
            
            recordId = self.__get_record_id(row)
            record = self.p_data[recordId]

            for col in range(1, self.data_end_col):
                field = self.field_list[col]
                field_name = field.name
                field_type = field.type

                if field_type == 'ref':
                    value = record[field_name]
                    ref_sheet_name = value.split('.')[0]
                    ref_recordId = value.split('.')[1]

                    if ref_recordId.isdigit():
                        ref_recordId = int(ref_recordId)
                    #TODO 不支持浮点数做主键

                    ref_sheet = self.mgr.get_sheet(ref_sheet_name)
                    if not ref_sheet:
                        raise Exception('Failed to get sheet in cell [%s%s] from [%s]' % (chr(col + ord('A')), row+1, self.sh.name))

                    ref_p_data = ref_sheet.to_python()
                    record[field_name] = ref_p_data[ref_recordId]


shm = SheetManager()

__all__ = ['shm']