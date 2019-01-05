# -*- coding: utf-8 -*-
"""
Created on Tue Dec 18 07:55:35 2018

@author: eamdf
"""

import xlrd #导入x1rd库
data=xlrd.open_workbook('AutoReport.xlsm') #打开Exce1文件
sh=data.sheet_by_name('挠度') #获得需要的表单
print (sh.cell_value(13,3)) #打印表单中B2值