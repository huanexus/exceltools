import datetime
import os
from collections import OrderedDict, defaultdict
from contextlib import contextmanager
from itertools import count

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.workbook import Workbook
today = datetime.date.today()

class EOWriter:

    def __init__(self):
        self.__book__ = defaultdict()
        self.__sheet__ = defaultdict(dict)
        self.__fields__ = defaultdict(dict)
        self.__firstrow__ = defaultdict(int)
        self.__data__ = OrderedDict()
        self.__fmt__ = defaultdict(dict)
        self.__sheetordered__ = True
        self.__eos__ = defaultdict(dict)
        self.msg = datetime.date.today().isoformat()


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

    def writeheader(self, row=0):
        worksheet = self.worksheet
        fmt = self.fmt
        msg = self.msg
        r0, r1, r = row, row+1, row+2
        worksheet.merge_range(r0, 0, r0, 4, msg)
        worksheet.merge_range(r1, 0, r1, 1, "监控设置", fmt["center"])
        worksheet.merge_range(r1, 2, r1, 4, "参考值", fmt["center"])
        worksheet.merge_range(r1, 5, r1, 6, "调整基准", fmt["center"])
        worksheet.merge_range(r1, 7, r1, 9, "调整系数", fmt["center"])
        worksheet.merge_range(r1, 10, r1, 11, "调整值", fmt["center"])
        worksheet.merge_range(r1, 12, r1, 13, "设置值", fmt["center"])
        worksheet.merge_range(r1, 14, r1, 16, "实际设置区间百分比", fmt["center"])
        worksheet.write(r, 0, "通道", fmt["center"])
        worksheet.write(r, 1, "监控条件", fmt["center"])
        worksheet.write(r, 2, "下限", fmt["center"])
        worksheet.write(r, 3, "上限", fmt["center"])
        s = "中心 基准 基准系数 移动M 下限A 上限B 下限C 上限D 下限L 上限U 下限% 上限% 区间宽度%".split()
        for c in range(4, len(s)+4):
            worksheet.write(r, c, s[c-4], fmt["center"])
        return row+3

    @contextmanager
    def writerow(self, row):
        worksheet = self.worksheet
        fmt = self.fmt        
        r = row
        fmt0 = fmt["0"]
        fmt1 = fmt["1"]
        worksheet.write_formula(r, 4, "($D{0}+$C{0})/2".format(r+1), fmt1)
        worksheet.write_formula(r, 5, "($D{0}-$C{0})/10".format(r+1), fmt1)
        worksheet.write_formula(r, 6, "$F{0}/$E{0}".format(r+1), fmt["percentage"])
        formula = "$C4+$F4*($H4+$I4)".replace("4", "{0}")
        worksheet.write_formula(r, 10, formula.format(r+1), fmt1)
        formula = "$F4*($H4+$J4)+$D4".replace("4", "{0}")
        worksheet.write_formula(r, 11, formula.format(r+1), fmt1)
        formulas = ["$K4", "$L4", "1-$M4/$E4", "$N4/$E4-1", "($N4-$M4)/$E4"]
        for c in range(12, 17):
            formula = formulas[c-12].replace("4", "{0}")
            if c < 14:
                worksheet.write_formula(r, c, formula.format(r+1), fmt1)
            else:
                worksheet.write_formula(r, c, formula.format(r+1), fmt["percentage"])
        yield
        cs = range(7, 10)
        for c in cs:
            worksheet.write(r, c, "", fmt0)


    def __call__(self, eos, names, channel_nr=1, msg=None):
        worksheet = self.worksheet
        worksheet.set_column(0, 0, 4)
        worksheet.set_column(2, 3, 6)
        fmt = self.fmt
        fmt0 = fmt["0"]
        if msg is not None:
            self.msg = msg            
        row = 0
        row = self.writeheader(row)
        row_id = count(row)
        for name, eo in zip(names, eos):
            r = next(row_id)
            with self.writerow(r):
                worksheet.write(r, 0, channel_nr, fmt["center"])
                worksheet.write(r, 1, name, fmt0) 
                worksheet.write(r, 2, eo[0], fmt0)
                worksheet.write(r, 3, eo[1], fmt0)
        self.__eos__[self.wb][self.sheet] = (eos, names)
        self.eos = self.__eos__[self.wb][self.sheet]
        return self

    def write_df(self, df):
        names = dict(ym="最大值", y0="垂直阈值", i0="积分值")
        worksheet = self.worksheet
        worksheet.set_column(0, 0, 4)
        worksheet.set_column(2, 3, 6)
        fmt = self.fmt
        fmt0 = fmt["0"]     
        row = 0
        row = self.writeheader(row)
        row_id = count(row)
        for column in df.columns:
            ya, yb = df[column].loc[["ya", "yb"]]
            r = next(row_id)
            channel_nr = 0       
            name = names.get(column.lower(), column)
            with self.writerow(r):
                worksheet.write(r, 0, channel_nr, fmt["center"])
                worksheet.write(r, 1, name, fmt0) 
                worksheet.write(r, 2, ya, fmt0)
                worksheet.write(r, 3, yb, fmt0)

    def save(self):
        for name, wb in self.__book__.items():
            path = os.path.split(name)[0]
            if path and (not os.path.exists(path)):
                os.makedirs(path)
            wb.close()
