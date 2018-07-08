# -*- coding: utf-8 -*-

import re
import pandas as pd
import numpy as np
import xlrd 
import xlsxwriter
from collections import Counter
from collections import defaultdict





sms_file = r'E:\\mdlv\\yumco\\Yumco\\sms.xlsx'
sms_result = r'E:\\mdlv\\yumco\\Yumco\\result.xlsx'
sms = pd.read_excel(sms_file)
sms_phone = sms['Контактное лицо']
sms_ls = sms['№ лиц. счета']
sms_type = sms['Тип выезда']
sms_brigada = ['Бригада']


counter = Counter(sms_ls)
sms_len = len(sms.index)

for i in range(0, sms_len):
    lf1 = sms_ls[i]
    sms_phone[i] = sms_phone[i].lower()
    sms_phone[i] = re.sub(r'[_-йцукенгшщзхъфывапролджэячсёмитьбю().+;,\s]','', sms_phone[i])
    for j in range(0, sms_len):    
        lf2 = sms_ls[j]
        if j != i:
            if lf1 == lf2:
                sms_phone[i] = ''
print('done')

          
writer_orig = pd.ExcelWriter(sms_result, engine='xlsxwriter')
sms.to_excel(writer_orig, index=False, sheet_name='report')
writer_orig.save()

