# -*- coding: utf-8 -*- 
import re
import pandas as pd
import numpy as np
import xlrd 
import xlsxwriter
from collections import Counter
from collections import defaultdict

sms_file = r'\\172.16.2.51\general\ДОКУМЕНТЫ ОТДЕЛОВ\ОКТП\Индикатор Статистика\Качество сотрудников\sms.xlsx'
sms_result = r'\\172.16.2.51\general\ДОКУМЕНТЫ ОТДЕЛОВ\ОКТП\Индикатор Статистика\Качество сотрудников\result.xlsx'
sms = pd.read_excel(sms_file)
sms_phone = sms['Контактное лицо'].str.replace('[^1234567890]','')
sms['Контактное лицо'] = sms_phone
sms_vd = sms['Присвоено на дату']
sms_ls = sms['№ лиц. счета']
sms_df = pd.DataFrame(sms, columns = ['№ лиц. счета','Присвоено на дату', 'Контактное лицо']).drop_duplicates('№ лиц. счета')

writer_orig = pd.ExcelWriter(sms_result, engine='xlsxwriter')
sms_df.to_excel(writer_orig, index=False, sheet_name='result')
writer_orig.save()
