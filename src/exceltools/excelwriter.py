# -*- coding: utf-8 -*-
import os
import copy
import xlsxwriter
from collections import defaultdict, OrderedDict
from xlsxwriter.workbook import Workbook
import datetime

def round_numeric(val, n=3):
    if isinstance(val, float):
        val = round(val, n)
    try:
        if float(val) == int(val):
            val = int(val)
    except (ValueError, TypeError):
        pass
    return val

class ExcelWriter:

    def __init__(self):
        self.__book__ = defaultdict(None)
        self.__sheet__ = defaultdict(dict)
        self.__fields__ = defaultdict(dict)
        self.__firstrow__ = defaultdict(int)
        self.__data__ = OrderedDict()
        self.__fmt__ = defaultdict(dict)

    def setworkbook(self, wb):
        workbooks = self.__book__
        if wb not in workbooks:
            workbook = workbooks[wb] = Workbook(wb, {'strings_to_numbers': True})
        self.wb = wb
        self.workbook = workbooks[wb]
        self.setformats()
        return self

    def setworksheet(self, wb, sheet):        
        workbooks = self.__book__
        sheets = self.__sheet__
        self.setworkbook(wb)
        workbook = self.workbook        
        try:
            ws = sheets[wb][sheet]
        except KeyError:
            sheets[wb][sheet] = workbook.add_worksheet(sheet)
        self.sheet = sheet
        self.worksheet = sheets[wb][sheet]
        return self     

    def setformats(self):
        wb = self.wb
        workbook = self.workbook
        fmt = self.__fmt__[wb]
        fmt['Bold'] = workbook.add_format({'bold': 1, 'align': 'center'})
        fmt['data'] = workbook.add_format()
        fmt['int'] = workbook.add_format({'num_format': '0'})
        fmt['float'] = workbook.add_format({'num_format': '0.00'})
        fmt['time'] = workbook.add_format({'num_format': 'yyyy-m-d h:mm;@'}) 
        fmt["center"] = workbook.add_format({"align": "center", "border": 1})
        fmt["date"] = workbook.add_format({"num_format": 'yyyy-mm-dd', "align": "left", "border": 1})
        fmt["left"] = workbook.add_format({"align": "left", "border": 1})
        fmt["percentage"] = workbook.add_format({"num_format": '0.00%', "border": 1})
        fmt["0"] = workbook.add_format({"border": 1})
        fmt["1"] = workbook.add_format({"num_format": '0.00', "border": 1})        
        for cfmt in fmt.values():
            cfmt.set_font_name(u'Arial')
            cfmt.set_font_size(10)
            cfmt.set_align('vcenter')
            cfmt.set_text_wrap(True)
        self.fmt = fmt
        return self

    def setfields(self, fields, firstrow=0):
        key = (self.wb, self.sheet)
        if key not in self.__fields__:
            self.__fields__[key] = fields
        if key not in self.__firstrow__:
            self.__firstrow__[key] = firstrow

    def __call__(self, msg):
        key = (self.wb, self.sheet)
        if key not in self.__data__:
            self.__data__[key] = []
        self.__data__[key].append(msg)

    def setwidth(self, wb, sheet):
        sh = self.__sheet__[wb][sheet]
        data = self.__data__[(wb, sheet)]
        fields = self.__fields__[(wb, sheet)]
        width = [[len(str(s)) for s in fields]]
        for row in data:
            width.append([len(str(s)) for s in list(row)])
        width = list(map(max, list(zip(*width))))
        for (col, width) in enumerate(width):
            col_name = xlsxwriter.utility.xl_col_to_name(col)
            col_name_range = '%s:%s' % (col_name, col_name)
            sh.set_column(col_name_range, min(width+1, 45))

    def writefields(self, wb, sheet, fields=None, firstrow=None):
        self.setworksheet(wb, sheet)
        ws = self.worksheet
        fmt = self.fmt
        fields = fields or self.__fields__[(wb, sheet)]
        firstrow = firstrow or self.__firstrow__[(wb, sheet)]
        for col, value in enumerate(fields):
            ws.write(firstrow, col, value, fmt["Bold"])       
        ws.freeze_panes(1, 1)

    def write(self, row, col, value):
        ws = self.worksheet
        fmt = self.fmt
        if isinstance(value, (int, )):
            if value < 1E15:
                ws.write(row, col, value, fmt["int"])
            else:
                ws.write(row, col, str(value), fmt["data"])                        
        elif isinstance(value, float):
            ws.write(row, col, str(value), fmt["float"])
        elif isinstance(value, datetime.datetime):
            ws.write(row, col, value, fmt["time"]) 
        elif isinstance(value, datetime.date):
            ws.write(row, col, value, fmt["date"])                  
        else:
            ws.write(row, col, value, fmt["data"])        

    def writedata(self, wb, sheet):
        ws = self.__sheet__[wb][sheet]
        start_row = self.__firstrow__[(wb, sheet)] + 1
        data = self.__data__[(wb, sheet)]
        self.setworksheet(wb, sheet)
        for rownum, data in enumerate(data, start_row):
            for colnum, value in enumerate(data):
                self.write(rownum, colnum, value)               

    def write_df(self, df, workbook, worksheet, precision=3):
        fields = df.columns.tolist()
        self.setworksheet(workbook, worksheet)
        self.setfields(fields)
        for k, record in df.iterrows():
            msg = [round_numeric(v, precision) for v in record]
            self(msg)         
        return self

    def save(self):        
        datas = self.__data__.items()
        for (wb, sheet), data in datas:
            fields = self.__fields__[(wb, sheet)]
            self.writefields(wb, sheet, fields)
            self.writedata(wb, sheet)
            self.setwidth(wb, sheet)
        for name, wb in self.__book__.items():
            path = os.path.split(name)[0]
            if path and (not os.path.exists(path)):
                os.makedirs(path)
            wb.close()

__all__ = ['ExcelWriter']
