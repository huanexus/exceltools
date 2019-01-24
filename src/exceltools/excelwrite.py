# -*- coding: utf-8 -*-
import os
import copy
import xlsxwriter
from collections import defaultdict, OrderedDict
from xlsxwriter.workbook import Workbook
import datetime

def round_num(v, n=3):
    value = v
    if isinstance(v, float):
        value = round(v, n)
    try:
        if float(v) == int(v):
            value = int(v)
    except:
        pass
    return value

class excelwrite(object):

    def __init__(self, fields = None, firstrow=None, sheetordered=False):
        self.fields = fields
        self.firstrow = firstrow or 0
        self.__book__ = defaultdict(None)
        self.__sheet__ = defaultdict(dict)
        self.__fields__ = defaultdict(dict)
        self.__firstrow__ = defaultdict(int)
        self.__data__ = OrderedDict()
        self.__fmt__ = defaultdict(dict)
        self.__sheetordered__ = sheetordered

    def __call__(self, wb, sheet, msg, fields=None, firstrow=None):
        data = list(copy.copy(msg))
        if (wb, sheet) not in self.__fields__:
            self.__fields__[(wb, sheet)] = fields or self.fields
        if (wb, sheet) not in self.__firstrow__:
            self.__firstrow__[(wb, sheet)] = firstrow or self.firstrow
        if (wb, sheet) not in self.__data__:
            self.__data__[(wb, sheet)] = []
        self.__data__[(wb, sheet)].append(data)

    def write_df(self, df, workbook, worksheet, precision=3):
        fields = df.columns.tolist()            
        for k, record in df.iterrows():
            msg = [round_num(v, precision) for v in record]
            self(workbook, worksheet, msg, fields)         
        return self        
    
    def set_column_width(self, wb, sheet):
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
            sh.set_column(col_name_range, min(width, 45))

    def write_fields(self, wb, sheet, fields=None, firstrow=None):
        workbooks = self.__book__
        sheets = self.__sheet__
        try:
            ws = sheets[wb][sheet]
        except KeyError:
            if wb not in workbooks:
                workbook = workbooks[wb] = Workbook(wb, {'strings_to_numbers': True})
                self.__fmt__[wb]['Bold'] = workbook.add_format({'bold': 1, 'align': 'center'})
                self.__fmt__[wb]['data'] = workbook.add_format()
                self.__fmt__[wb]['int'] = workbook.add_format({'num_format': '0'})
                self.__fmt__[wb]['float'] = workbook.add_format({'num_format': '0.00'})
                self.__fmt__[wb]['time'] = workbook.add_format({'num_format': 'yyyy-m-d h:mm;@'})
                self.__fmt__[wb]['date'] = workbook.add_format({'num_format': 'yyyy-m-d'})
                for fmt in self.__fmt__[wb].values():
                    fmt.set_font_name(u'Arial')
                    fmt.set_font_size(10)
##                    fmt.set_border(1)
                    fmt.set_align('vcenter')
                    fmt.set_text_wrap(True)
                    
            sheets[wb][sheet] = workbooks[wb].add_worksheet(sheet)
        ws = sheets[wb][sheet]
        bold = self.__fmt__[wb]['Bold']
        fields = fields or self.__fields__[(wb, sheet)]
        firstrow = firstrow or self.__firstrow__[(wb, sheet)]
        for col, value in enumerate(fields):
            ws.write(firstrow, col, value, bold)       
        ws.freeze_panes(1, 1)

    def write_data(self, wb, sheet):
        ws = self.__sheet__[wb][sheet]
        start_row = self.__firstrow__[(wb, sheet)] + 1
        data = self.__data__[(wb, sheet)]
        fmt_data = self.__fmt__[wb]['data']
        fmt_int = self.__fmt__[wb]['int']
        fmt_float = self.__fmt__[wb]['float']
        fmt_time = self.__fmt__[wb]['time']
        fmt_date = self.__fmt__[wb]['date']
        
        for rownum, data in enumerate(data, start_row):
            for colnum, value in enumerate(data):
                if isinstance(value, (int, )):
                    if value < 1E15:
                        ws.write(rownum, colnum, value, fmt_int)
                    else:
                        ws.write(rownum, colnum, str(value), fmt_data)
                elif isinstance(value, float):
                    ws.write(rownum, colnum, str(value), fmt_float)
                elif isinstance(value, datetime.datetime):
                    ws.write(rownum, colnum, value, fmt_time) 
                elif isinstance(value, datetime.date):
                    ws.write(rownum, colnum, value, fmt_date)                  
                else:
                    ws.write(rownum, colnum, value, fmt_data)

    def save(self, overwrite=None):
        overwrite = overwrite or self.overwrite
        data = self.__data__.items()
        if self.__sheetordered__ == True:
            data = sorted(data)
        for (wb, sheet), data in data:
            fields = self.__fields__[(wb, sheet)]
            self.write_fields(wb, sheet, fields)
            self.write_data(wb, sheet)
            self.set_column_width(wb, sheet)
        for name, wb in self.__book__.items():
            path = os.path.split(name)[0]
            if path and (not os.path.exists(path)):
                os.makedirs(path)
            wb.close()

__all__ = ['excelwrite']
