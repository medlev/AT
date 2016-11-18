import xlrd, xlwt
from xlwt import *
from datetime import *
import datetime

#работа с датами, верный формат, больше суток, двое, 14 дней
now = datetime.datetime.now()
today = now.date()
twodays = today - datetime.timedelta(2)
onedays = today - datetime.timedelta(1)
fo14 = today - datetime.timedelta(14)
dt=today.strftime('%d.%m.%Y')
dt1=onedays.strftime('%d.%m.%Y')
dt2=twodays.strftime('%d.%m.%Y')
dt14=fo14.strftime('%d.%m.%Y')
#указываем откуда читать ексель
rb = xlrd.open_workbook(r'C:/Users/yusmedlev/OneDrive/stat/all.xlsx')
sheet = rb.sheet_by_index(0)
rb1 = xlrd.open_workbook(r'C:/Users/yusmedlev/OneDrive/stat/all_result.xls')
sheet1 = rb.sheet_by_index(0)
#указываем куда сохраняем
wb = xlwt.Workbook(r'C:/Users/yusmedlev/OneDrive/stat/all_result.xls')
#форматирование экселя
font0 = Font()
font0.name = 'Tahoma'
font0.bold = True
style0 = XFStyle()
style0.font = font0
style1 = xlwt.easyxf('align: rotation 90')

#1 - всего заявок
def oldtable(x):
    soch=sotz=semail=swelcomecall=soch_eng=sotz_eng=semail_eng=swelcomecall_eng=0
    ws.row(0).height = 2500
    ws.col(0).width = 3500
    for i in range(50):
        ws.col(i+1).width = 1000
    list1 = ['Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова', 'Виталий БУРЫЙ', 'Андрей Арапов',
    'Владимир Михаловский', 'Дмитрий Гуменников', 'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз',
    'Виктор Журавель', 'Никита Зайченко', 'Павел Жаворонок', 'Николай Волынец', 'Семен Дубовский',
    'Никита Пастухов', 'Дмитрий Бородин', 'Александра Моисеенко', 'Вадим Писарев', 'Илья Хрисанфов',
    'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    list2 = ['Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова', 'Виталий БУРЫЙ', 'Андрей Арапов',
    'Владимир Михаловский', 'Дмитрий Гуменников', 'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз',
    'Виктор Журавель', 'Никита Зайченко', 'Павел Жаворонок', 'Николай Волынец', 'Семен Дубовский',
    'Никита Пастухов', 'Дмитрий Бородин', 'Александра Моисеенко', 'Вадим Писарев', 'Илья Хрисанфов',
    'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    list3 = ['']
    for i in range(1,sheet.nrows):
        if sheet.row_values(i)[14] == x:
            if sheet.row_values(i)[6] == 'ОЧЕРЕДЬ':
                soch+=1
            elif sheet.row_values(i)[6] == 'ОТЗВОН':
                sotz+=1
            elif sheet.row_values(i)[6] == 'E-MAIL':
                semail+=1
            elif sheet.row_values(i)[6] == 'WELCOME CALL':
                swelcomecall+=1
    for i in range(1,sheet.nrows):
        if sheet.row_values(i)[14] != '':
            if sheet.row_values(i)[6] == 'ОЧЕРЕДЬ':
                soch_eng+=1
            elif sheet.row_values(i)[6] == 'ОТЗВОН':
                sotz_eng+=1
            elif sheet.row_values(i)[6] == 'E-MAIL':
                semail_eng+=1
            elif sheet.row_values(i)[6] == 'WELCOME CALL':
                swelcomecall_eng+=1
    sum=[semail_eng,sotz_eng+swelcomecall_eng,soch_eng, x, semail_eng+sotz_eng+swelcomecall_eng+soch_eng]
    sum1=[semail,sotz+swelcomecall,soch, x, semail+sotz+swelcomecall+soch]
    for i in range(23):
        if sum1[3] == list1[i]:
            ws.write(0, i+1, x, style1)
            ws.write(1, i+1, semail, style0)
            ws.write(2, i+1, sotz+swelcomecall, style0)
            ws.write(3, i+1, soch, style0)
            ws.write(4, i+1, semail+sotz+swelcomecall+soch, style0)
    if sum1[3] == '':
            ws.write(6, 1, semail, style0)
            ws.write(7, 1, sotz+swelcomecall, style0)
            ws.write(8, 1, soch, style0)
            ws.write(9, 1, semail+sotz+swelcomecall+soch, style0)
    if sum1[3] != '':
            soch_eng+=1
            sotz_eng+=1
            semail_eng+=1
            swelcomecall_eng+=1
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/all_result.xls')

#2 - табель по заявкам
def newtable(x):
    table17=table16=table1=table2=table3=table4=table5=table6=table7=table8=table9=table10=table11=table12=table13=table14=table15=0
    ws.row(0).height = 2500
    ws.col(0).width = 3000
    for i in range(50):
        ws.col(i+1).width = 1000
    list1 = ['Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова', 'Виталий БУРЫЙ', 'Андрей Арапов',
    'Владимир Михаловский', 'Дмитрий Гуменников', 'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз',
    'Виктор Журавель', 'Никита Зайченко', 'Павел Жаворонок', 'Николай Волынец', 'Никита Пастухов', 'Дмитрий Бородин',
    'Александра Моисеенко', 'Вадим Писарев', 'Илья Хрисанфов', 'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    for i in range(1,sheet.nrows):
        if sheet.row_values(i)[14] == x:
            sh2 = sheet.row_values(i)[12]
            if sh2 == 'ОДТП физ > Срочнопорт > Пересоздать MVR' or sh2 == 'ОДТП физ > Срочнопорт > Изменение параметров на порту' or sh2 == 'ОДТП физ > Срочнопорт > Включение порта' or sh2 == 'ОДТП физ > Срочнопорт > Изменение зарезки ADSL' or sh2 == 'ОДТП физ > Срочнопорт > Изменение зарезки Ethernet' or sh2 == 'ОДТП физ > Срочнопорт > Изменение физической скорости Ethern':
                table1+=1
            elif sh2 == 'ОДТП физ > Консультация > Подключение' or sh2 == 'ОДТП физ > Консультация > NAT' or sh2 == 'ОДТП физ > Консультация > DNS' or sh2 == 'ОДТП физ > Консультация > Статический адрес' or sh2 == 'ОДТП физ > Консультация > Технология подключения' or sh2 == 'ОДТП физ > Консультация > Выбор оборудования для подлкючения' or sh2 == 'ОДТП физ > Консультация > Стороннее ПО' or sh2 == 'ОДТП физ > Консультация > Стороннее оборудование клиента':
                table2+=1
            elif sh2 == 'ОДТП физ > Прочее > Другое' or sh2 == 'ОДТП физ > Неидентифицировано > Клиент неизвестен' or sh2 == 'ОДТП физ > Тех. Проблема > Другое' or sh2 == 'ОДТП физ > Тех. Проблема > Не доступен узел ДС/БЦ/АТС' or sh2 == 'ОДТП физ > Тех. Проблема > Проблемы с BRAS':
                table3+=1
            elif sh2 == 'ОДТП физ > Не работает интернет > Нет линка':
                table4+=1
            elif sh2 == 'ОДТП физ > Диагностика > Стороннее оборудование клиента':
                table5+=1
            elif sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Другое' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Не включается/загружает' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Нет вещания всех канало' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Нет вещания одного или' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Нет списка каналов / ош' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Отсутствует видео, звук' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Ошибка загрузки портала' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Пропадает изображение' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Рассыпается изображение' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Отложенный просмотр':
                table5+=1
            elif sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Подвисает изображение' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Пульт' or sh2 == 'ОДТП физ > IPTV диагностика MAG250 > Смещается время в прист':
                table5+=1
            elif sh2 == 'ОДТП физ > IPTV диагностика RedBox > Другое' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Не включается/загружает' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Нет вещания всех канало' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Нет вещания одного или' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Нет списка каналов / ош' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Отсутствует видео, звук' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Ошибка загрузки портала' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Пропадает изображение' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Рассыпается изображение' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Отложенный просмотр':
                table5+=1
            elif sh2 == 'ОДТП физ > IPTV диагностика RedBox > Подвисает изображение' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Внутренние ресурсы прис' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Не совместим ТВ с RedBo' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Нет доп. лицензии' or sh2 == 'ОДТП физ > IPTV диагностика RedBox > Пульт':
                table5+=1
            elif sh2 == 'ОДТП физ > IPTV-диагностика ПК > Другое' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Нет вещаниея одного или нес' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Нет вещания всех каналов' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Нет списка каналов/ошибка п' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Ошибка при загрузке каналов' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Отсутствует видео, звук' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Пропадает изображение' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Рассыпается изображение' or sh2 == 'ОДТП физ > IPTV-диагностика ПК > Подвисает изображение':
                table5+=1
            elif sh2 == 'ОДТП физ > IPTV-диагностика > Нет вещаниея одного или нескол' or sh2 == 'ОДТП физ > IPTV-диагностика > Нет списка каналов/ошибка при' or sh2 == 'ОДТП физ > IPTV-диагностика > Отсутствует видео, звук есть' or sh2 == 'ОДТП физ > IPTV-диагностика > Ошибка при загрузке каналов' or sh2 == 'ОДТП физ > IPTV-диагностика > Пропадает изображение' or sh2 == 'ОДТП физ > IPTV-диагностика > Рассыпается изображение':
                table5+=1
            elif sh2 == 'ОДТП физ > Диагностика > Dr. Web' or sh2 == 'ОДТП физ > Диагностика > IPTV Smart TV' or sh2 == 'ОДТП физ > Диагностика > video.telecom.by' or sh2 == 'ОДТП физ > Диагностика > Wi-Fi' or sh2 == 'ОДТП физ > Диагностика > Внутренние ресурсы (остальное)':
                table5+=1
            elif sh2 == 'ОДТП физ > Не работает интернет > Нет мака' or sh2 == 'ОДТП физ > Не работает интернет > Нет авторизации' or sh2 == 'ОДТП физ > Не работает интернет > Не грузятся страницы при н' or sh2 == 'ОДТП физ > Не работает интернет > Не получает IP-адрес' or sh2 == 'ОДТП физ > Не работает интернет > Нет доступа к конкретному' or sh2 == 'ОДТП физ > Не работает интернет > Оборудование провайдера не' or sh2 == 'ОДТП физ > Не работает интернет > Viloation rate' or sh2 == 'ОДТП физ > Не работает интернет > Hash limit':
                table6+=1
            elif sh2 == 'ОДТП физ > Низкая скорость > Интернет' or sh2 == 'ОДТП физ > Низкая скорость > Внутренние ресурсы':
                table7+=1
            elif sh2 == 'ОДТП физ > Дисконнекты > Интернет':
                table8+=1
            elif sh2 == 'ОДТП физ > Прочее > SPAM':
                table9+=1
            elif  sh2 == 'ОДТП физ > Настройка > IPTV mag250' or sh2 == 'ОДТП физ > Настройка > IPTV RedBox' or sh2 == 'ОДТП физ > Настройка > IPTV Smart TV' or sh2 == 'ОДТП физ > Настройка > IPTV ПК' or sh2 == 'ОДТП физ > Настройка > Smart TV':
                table10+=1
            elif sh2 == 'ОДТП физ > Не работает > IPTV mag250' or sh2 == 'ОДТП физ > Не работает > IPTV ПК' or sh2 == 'ОДТП физ > Не работает > IPTV mag250' or sh2 == 'ОДТП физ > Не работает > IPTV ПК':
                table10+=1
            elif  sh2 == 'ОДТП физ > Не работает > video.telecom.by' or sh2 == 'ОДТП физ > Не работает > DC++' or sh2 == 'ОДТП физ > Не работает > Внутренние ресурсы (остальное)' or sh2 == 'ОДТП физ > Не работает > Почта' or sh2 == 'ОДТП физ > Настройка > DC++' or sh2 == 'ОДТП физ > Настройка > Почта' or sh2 == 'ОДТП физ > Настройка > Dr. Web':
                table11+=1
            elif sh2 == 'ОДТП физ > Прочее > Деталька':
                table12+=1
            elif sh2 == 'ОДТП физ > Настройка > Маршрутизатор АТ' or sh2 == 'ОДТП физ > Настройка > Маршрутизатор клиента' or sh2 == 'ОДТП физ > Настройка > Модем' or sh2 == 'ОДТП физ > Настройка > ПК' or sh2 == 'ОДТП физ > Настройка > Другое' or sh2 == '`ОДТП физ > Настройка > Стороннее оборудование клиента`':
                table13+=1
            elif sh2 == 'ОДТП физ > Не работает > Wi-Fi' or sh2 == 'ОДТП физ > Настройка > WI-FI' or sh2 == 'ОДТП физ > Низкая скорость > WI-FI' or sh2 == 'ОДТП физ > Дисконнекты > WI-FI':
                table14+=1
            elif sh2 == 'ОДТП физ > Настройка > Видеонаблюдение':
                table15+=1
            elif sh2 == 'ОДТП физ > Спам > Спам':
                table16+=1
            else:
                table17+=1
    sum1=[table1, table2, table3, table4, table5, table6, table7, table8, table9, table10, table11, table12, table13, table14, table15, table16, x]
    for i in range(23):
        if sum1[16] == list1[i]:
            ws.write(0, i+1, x, style1)
            ws.write(1, i+1, table1, style0)
            ws.write(2, i+1, table2, style0)
            ws.write(3, i+1, table3, style0)
            ws.write(4, i+1, table4, style0)
            ws.write(5, i+1, table5, style0)
            ws.write(6, i+1, table6, style0)
            ws.write(7, i+1, table7, style0)
            ws.write(8, i+1, table8, style0)
            ws.write(9, i+1, table9, style0)
            ws.write(10, i+1, table10, style0)
            ws.write(11, i+1, table11, style0)
            ws.write(12, i+1, table12, style0)
            ws.write(13, i+1, table13, style0)
            ws.write(14, i+1, table14, style0)
            ws.write(15, i+1, table15, style0)
            ws.write(16, i+1, table16, style0)
            ws.write(17, i+1, table17, style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/res.xls')

#3 - без изменений больше 2 суток по кол-ву заявок
def onemore(x):
    iz=0
    ws.row(0).height = 2500
    for i in range(50):
        ws.col(i+1).width = 3000
    list1 = ['Александра Моисеенко', 'Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова',
            'Вадим Писарев', 'Виталий БУРЫЙ', 'Андрей Арапов', 'Владимир Михаловский', 'Дмитрий Гуменников',
            'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз', 'Виктор Журавель', 'Никита Зайченко',
            'Николай Волынец', 'Семен Дубовский', 'Никита Пастухов', 'Дмитрий Бородин', 'Илья Хрисанфов',
            'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    sum=[x]
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[7]
        a = int(sh[:2])
        b = int(sh[3:5])
        c = int(sh[6:10])
        fo1 = today - datetime.timedelta(2)
        if datetime.date(c, b, a) < fo1:
            if sheet.row_values(i)[11] == x:
                for im in range(23):
                    if sum[0]== list1[im]:
                        iz+=4
                        ws.write(iz+5,im,sheet.row_values(i)[0], style0)
                        ws.write(iz+6,im,sheet.row_values(i)[7], style0)
                        ws.write(iz+7,im,sheet.row_values(i)[2], style0)
                        ws.write(iz+8,im,sheet.row_values(i)[9], style0)
    for i in range(23):
        if sum[0] == list1[i]:
            ws.write(0, i, x, style1)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/2day_result_id.xls')

#4 - без изменений больше 2 суток по ID
def twoday(x):
    soch=0
    sotz=0
    semail=0
    swelcomecall=0
    ws.row(0).height = 2500
    for i in range(50):
        ws.col(i+1).width = 1000
    list1 = ['Александра Моисеенко', 'Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова',
            'Вадим Писарев', 'Виталий БУРЫЙ', 'Андрей Арапов', 'Владимир Михаловский', 'Дмитрий Гуменников',
            'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз', 'Виктор Журавель', 'Никита Зайченко',
            'Николай Волынец', 'Семен Дубовский', 'Никита Пастухов', 'Дмитрий Бородин', 'Илья Хрисанфов',
            'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[7]
        a = int(sh[:2])
        b = int(sh[3:5])
        c = int(sh[6:10])
        fo1 = today - datetime.timedelta(2)
        if datetime.date(c, b, a) < fo1:
            if sheet.row_values(i)[11] == x:
                if sheet.row_values(i)[6] == 'ОЧЕРЕДЬ':
                    soch+=1
                elif sheet.row_values(i)[6] == 'ОТЗВОН':
                    sotz+=1
                elif sheet.row_values(i)[6] == 'E-MAIL':
                    semail+=1
                elif sheet.row_values(i)[6] == 'WELCOME CALL':
                    swelcomecall+=1
    sum1=[semail,sotz+swelcomecall,soch, x, semail+sotz+swelcomecall+soch]
    sum =[semail,sotz+swelcomecall,soch]
    for i in range(23):
        if sum1[3] == list1[i]:
            ws.write(0, i+1, x, style1)
            ws.write(1, i+1, semail, style0)
            ws.write(2, i+1, sotz+swelcomecall, style0)
            ws.write(3, i+1, soch, style0)
            ws.write(4, i+1, semail+sotz+swelcomecall+soch, style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/2day_result.xls')

#5 - без изменений больше 14 суток по кол-ву заявок
def fo14(x):
    #ws = wb.add_sheet('>14')
    soch=0
    sotz=0
    semail=0
    swelcomecall=0
    ws.row(0).height = 2500
    for i in range(50):
        ws.col(i+1).width = 1000
    list1 = ['Александра Моисеенко', 'Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова',
            'Вадим Писарев', 'Виталий БУРЫЙ', 'Андрей Арапов', 'Владимир Михаловский', 'Дмитрий Гуменников',
            'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз', 'Виктор Журавель', 'Никита Зайченко',
            'Николай Волынец', 'Семен Дубовский', 'Никита Пастухов', 'Дмитрий Бородин', 'Илья Хрисанфов',
            'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[3]
        a = int(sh[:2])
        b = int(sh[3:5])
        c = int(sh[6:10])
        fo = today - datetime.timedelta(14)
        if datetime.date(c, b, a) < fo:
            if sheet.row_values(i)[11] == x:
                if sheet.row_values(i)[6] == 'ОЧЕРЕДЬ':
                    soch+=1
                elif sheet.row_values(i)[6] == 'ОТЗВОН':
                    sotz+=1
                elif sheet.row_values(i)[6] == 'E-MAIL':
                    semail+=1
                elif sheet.row_values(i)[6] == 'WELCOME CALL':
                    swelcomecall+=1
    sum1=[semail,sotz+swelcomecall,soch, x, semail+sotz+swelcomecall+soch]
    sum =[semail,sotz+swelcomecall,soch]
    for i in range(23):
        if sum1[3] == list1[i]:
            ws.write(0, i+1, x, style1)
            ws.write(1, i+1, semail, style0)
            ws.write(2, i+1, sotz+swelcomecall, style0)
            ws.write(3, i+1, soch, style0)
            ws.write(4, i+1, semail+sotz+swelcomecall+soch, style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/fo14.xls')

#6 - без изменений больше 14 суток по ID
def fo14id(x):
    iz=0
    ws.row(0).height = 2500
    ws.col(0).width = 3500
    ws.col(1).width = 2500
    ws.col(2).width = 2700
    ws.col(3).width = 5000
    ws.col(4).width = 16000
    ws.col(5).width = 3000
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[3]
        a = int(sh[:2])
        b = int(sh[3:5])
        c = int(sh[6:10])
        fo = today - datetime.timedelta(14)
        if datetime.date(c, b, a) < fo:
            iz+=1
            ws.write(iz,0, sheet.row_values(i)[0], style0)
            ws.write(iz,1, sheet.row_values(i)[1], style0)
            ws.write(iz,2, sheet.row_values(i)[3], style0)
            ws.write(iz,3, sheet.row_values(i)[11], style0)
            ws.write(iz,4, sheet.row_values(i)[10], style0)
            ws.write(iz,5, sheet.row_values(i)[13], style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/fo14id.xls')

#7 - закрытие в течении 24 часов
def oneday(x):
    closed=closed1=uncl=k=0
    listclosed = ['01.10.2016', '02.10.2016', '03.10.2016', '04.10.2016', '05.10.2016', '06.10.2016', '07.10.2016',
              '08.10.2016', '09.10.2016', '10.10.2016', '11.10.2016', '12.10.2016', '13.10.2016', '14.10.2016',
              '15.10.2016', '16.10.2016', '17.10.2016', '18.10.2016', '19.10.2016', '20.10.2016', '21.10.2016',
              '22.10.2016', '23.10.2016', '24.10.2016', '25.10.2016', '26.10.2016', '27.10.2016', '28.10.2016',
              '29.10.2016', '30.10.2016', '31.10.2016']
    ws.row(0).height = 2500
    ws.col(0).width = 3500
    ws.col(1).width = 1000
    ws.col(2).width = 1000
    for i in range(4,50):
        ws.col(i).width = 3000
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[3]
        a = int(sh[:2])
        b = int(sh[3:5])
        c = int(sh[6:10])
        see = ''
        one = datetime.date(c, b, a)
        oneone = one + datetime.timedelta(1)
        onetrue=oneone.strftime('%d.%m.%Y')
        if str(sheet.row_values(i)[3]) == x:
            closed+=1
            if str(sheet.row_values(i)[6]) == "E-MAIL":
                closed1+=1
            elif str(sheet.row_values(i)[9]) == x:
                closed1+=1
            elif str(sheet.row_values(i)[9]) == onetrue:
                closed1+=1
            elif str(sheet.row_values(i)[5]) == 'Есть объединяющая заявка':
                ##fix
                #closed-=1
                True
            else:
                uncl += 1
                see = see + str(sheet.row_values(i)[0])+ ' ' + str(sheet.row_values(i)[2]) + ' ' + str(sheet.row_values(i)[9])
                ws.write(uncl, int(sh[:2])+3, see)
    sum1=[x, closed, closed1, uncl]
    for i in range(len(listclosed)):
        if sum1[0] == listclosed[i]:
            ws.write(i, 1, sum1[1], style0)
            ws.write(i, 2, sum1[2], style0)
            ws.write(i, 0, listclosed[i], style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/closedres.xls')

# закрытие без левых об.заявок
def oneday_mix(x):
    closed=closed1=uncl=k=0
    listclosed = ['01.10.2016', '02.10.2016', '03.10.2016', '04.10.2016', '05.10.2016', '06.10.2016', '07.10.2016',
              '08.10.2016', '09.10.2016', '10.10.2016', '11.10.2016', '12.10.2016', '13.10.2016', '14.10.2016',
              '15.10.2016', '16.10.2016', '17.10.2016', '18.10.2016', '19.10.2016', '20.10.2016', '21.10.2016',
              '22.10.2016', '23.10.2016', '24.10.2016', '25.10.2016', '26.10.2016', '27.10.2016', '28.10.2016',
              '29.10.2016', '30.10.2016', '31.10.2016']
    ws.row(0).height = 2500
    ws.col(0).width = 3500
    ws.col(1).width = 1000
    ws.col(2).width = 1000
    for i in range(4,50):
        ws.col(i).width = 3000
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[3]
        a = int(sh[:2])
        b = int(sh[3:5])
        c = int(sh[6:10])
        see = ''
        one = datetime.date(c, b, a)
        oneone = one + datetime.timedelta(1)
        onetrue=oneone.strftime('%d.%m.%Y')
        if str(sheet.row_values(i)[5]) == 'Есть объединяющая заявка':
            if str(sheet.row_values(i)[3]) == x:
                closed+=1
                if str(sheet.row_values(i)[6]) == "E-MAIL":
                    closed1+=1
                elif str(sheet.row_values(i)[9]) == x:
                    closed1+=1
                elif str(sheet.row_values(i)[9]) == onetrue:
                    closed1+=1
                elif str(sheet.row_values(i)[5]) == 'Есть объединяющая заявка':
                    closed-=1
                else:
                    uncl += 1
                    see = see + str(sheet.row_values(i)[0])+ ' ' + str(sheet.row_values(i)[2]) + ' ' + str(sheet.row_values(i)[9])
                    ws.write(uncl, int(sh[:2])+3, see)
    sum1=[x, closed, closed1, uncl]
    for i in range(len(listclosed)):
        if sum1[0] == listclosed[i]:
            ws.write(i, 1, sum1[1], style0)
            ws.write(i, 2, sum1[2], style0)
            ws.write(i, 0, listclosed[i], style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/closedres.xls')

#9 indicator
def indicator(x):
    closed=closed1=uncl=k=0
    listclosed = ['01.10.2016', '02.10.2016', '03.10.2016', '04.10.2016', '05.10.2016', '06.10.2016', '07.10.2016',
              '08.10.2016', '09.10.2016', '10.10.2016', '11.10.2016', '12.10.2016', '13.10.2016', '14.10.2016',
              '15.10.2016', '16.10.2016', '17.10.2016', '18.10.2016', '19.10.2016', '20.10.2016', '21.10.2016',
              '22.10.2016', '23.10.2016', '24.10.2016', '25.10.2016', '26.10.2016', '27.10.2016', '28.10.2016',
              '29.10.2016', '30.10.2016', '31.10.2016']
    list1 = ['Александра Моисеенко', 'Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова',
            'Вадим Писарев', 'Виталий БУРЫЙ', 'Андрей Арапов', 'Владимир Михаловский', 'Дмитрий Гуменников',
            'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз', 'Виктор Журавель', 'Никита Зайченко',
            'Павел Жаворонок', 'Николай Волынец', 'Семен Дубовский', 'Никита Пастухов', 'Дмитрий Бородин',
            'Илья Хрисанфов', 'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    ws.row(0).height = 2500
    ws.col(0).width = 3500
    ws.col(1).width = 1400
    ws.col(2).width = 1400
    for i in range(1,sheet.nrows):
# esli closed to 9, esli created - 3
        if str(sheet.row_values(i)[9]) == x:
            closed+=1
    sum1=[x, closed]
    for i in range(len(listclosed)):
        if sum1[0] == listclosed[i]:
            ws.write(i, 1, sum1[1], style0)
            ws.write(i, 0, listclosed[i], style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/indicator.xls')





#10
def dublicat(x):
    dubl = []
    count=0
    count_dubl=0
    for i in range(1,sheet.nrows):
        sh = sheet.row_values(i)[1]
        for j in range(1,sheet.nrows):
            if j != i:
                sh1 = sheet.row_values(j)[1]
                if sh == sh1:
                    if sh != '':
                        dubl.insert(count,sh)
                        count+=1
                        dubl.sort()

    for i in range(len(dubl)):
        if dubl[i]!=dubl[i-1]:
            print(dubl[i])
            count_dubl+=1
    for i in range(len(dubl)):
        if dubl[i]!=dubl[i-1] and dubl[i]!='':
            ws.write(i, 0, dubl[i], style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/dublicat.xls')


#11
def dynamic(x):
    table31=table30=table29=table28=table27=table26=table25=table24=table23=table22=table21=table20=table19=table18=table17=table16=table1=table2=table3=table4=table5=table6=table7=table8=table9=table10=table11=table12=table13=table14=table15=0
    ws.row(0).height = 2500
    ws.col(0).width = 3000
    for i in range(50):
        ws.col(i+1).width = 1000
    list1 = ['Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова', 'Виталий БУРЫЙ', 'Андрей Арапов',
    'Владимир Михаловский', 'Дмитрий Гуменников', 'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз',
    'Виктор Журавель', 'Никита Зайченко', 'Павел Жаворонок', 'Николай Волынец', 'Семен Дубовский',
    'Никита Пастухов', 'Дмитрий Бородин', 'Александра Моисеенко', 'Вадим Писарев', 'Илья Хрисанфов',
    'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
    for i in range(1,sheet.nrows):
        if sheet.row_values(i)[14] == x:
            sh2 = sheet.row_values(i)[9]
            if sh2 == '01.10.2016':
                table1+=1
            elif sh2 == '02.10.2016':
                table2+=1
            elif sh2 == '03.10.2016':
                table3+=1
            elif sh2 == '04.10.2016':
                table4+=1
            elif sh2 == '05.10.2016':
                table5+=1
            elif sh2 == '06.10.2016':
                table6+=1
            elif sh2 == '07.10.2016':
                table7+=1
            elif sh2 == '08.10.2016':
                table8+=1
            elif sh2 == '09.10.2016':
                table9+=1
            elif sh2 == '10.10.2016':
                table10+=1
            elif sh2 == '11.10.2016':
                table11+=1
            elif sh2 == '12.10.2016':
                table12+=1
            elif sh2 == '13.10.2016':
                table13+=1
            elif sh2 == '14.10.2016':
                table14+=1
            elif sh2 == '15.10.2016':
                table15+=1
            elif sh2 == '16.10.2016':
                table16+=1
            elif sh2 == '17.10.2016':
                table17+=1
            elif sh2 == '18.10.2016':
                table18+=1
            elif sh2 == '19.10.2016':
                table19+=1
            elif sh2 == '20.10.2016':
                table20+=1
            elif sh2 == '21.10.2016':
                table21+=1
            elif sh2 == '22.10.2016':
                table22+=1
            elif sh2 == '23.10.2016':
                table23+=1
            elif sh2 == '24.10.2016':
                table24+=1
            elif sh2 == '25.10.2016':
                table25+=1
            elif sh2 == '26.10.2016':
                table26+=1
            elif sh2 == '27.10.2016':
                table27+=1
            elif sh2 == '28.10.2016':
                table28+=1
            elif sh2 == '29.10.2016':
                table29+=1
            elif sh2 == '30.10.2016':
                table30+=1
            elif sh2 == '31.10.2016':
                table31+=1
    sum1=[table1, table2, table3, table4, table5, table6, table7, table8, table9, table10, table11, table12, table13, table14, table15, table16, table17, table18, table19, table20, table21, table22, table23, table24, table25, table26, table27, table28, table29, table30, table31, x]
    for i in range(23):
        if sum1[31] == list1[i]:
            ws.write(0, i+1, sum1[31], style1)
            ws.write(1, i+1, table1, style0)
            ws.write(2, i+1, table2, style0)
            ws.write(3, i+1, table3, style0)
            ws.write(4, i+1, table4, style0)
            ws.write(5, i+1, table5, style0)
            ws.write(6, i+1, table6, style0)
            ws.write(7, i+1, table7, style0)
            ws.write(8, i+1, table8, style0)
            ws.write(9, i+1, table9, style0)
            ws.write(10, i+1, table10, style0)
            ws.write(11, i+1, table11, style0)
            ws.write(12, i+1, table12, style0)
            ws.write(13, i+1, table13, style0)
            ws.write(14, i+1, table14, style0)
            ws.write(15, i+1, table15, style0)
            ws.write(16, i+1, table16, style0)
            ws.write(17, i+1, table17, style0)
            ws.write(18, i+1, table18, style0)
            ws.write(19, i+1, table19, style0)
            ws.write(20, i+1, table20, style0)
            ws.write(21, i+1, table21, style0)
            ws.write(22, i+1, table22, style0)
            ws.write(23, i+1, table23, style0)
            ws.write(24, i+1, table24, style0)
            ws.write(25, i+1, table25, style0)
            ws.write(26, i+1, table26, style0)
            ws.write(27, i+1, table27, style0)
            ws.write(28, i+1, table28, style0)
            ws.write(29, i+1, table29, style0)
            ws.write(30, i+1, table30, style0)
            ws.write(31, i+1, table31, style0)
    wb.save(r'C:/Users/yusmedlev/OneDrive/stat/dynamic.xls')


#список сотрудников
list = ['Павел Демиденко', 'Алина Гафиятуллина', 'Екатерина Гончарова', 'Виталий БУРЫЙ', 'Андрей Арапов',
'Владимир Михаловский', 'Дмитрий Гуменников', 'Александр Дашкевич', 'Евгений Горбачев', 'Андрей Мороз',
'Виктор Журавель', 'Никита Зайченко', 'Павел Жаворонок', 'Николай Волынец', 'Никита Пастухов', 'Дмитрий Бородин',
'Александра Моисеенко', 'Вадим Писарев', 'Илья Хрисанфов', 'Евгений Клейнер', 'Евгений Михалёнок', 'Юрий Медлев', '']
#месяц
listclosed = ['01.10.2016', '02.10.2016', '03.10.2016', '04.10.2016', '05.10.2016', '06.10.2016', '07.10.2016',
              '08.10.2016', '09.10.2016', '10.10.2016', '11.10.2016', '12.10.2016', '13.10.2016', '14.10.2016',
              '15.10.2016', '16.10.2016', '17.10.2016', '18.10.2016', '19.10.2016', '20.10.2016', '21.10.2016',
              '22.10.2016', '23.10.2016', '24.10.2016', '25.10.2016', '26.10.2016', '27.10.2016', '28.10.2016',
              '29.10.2016', '30.10.2016', '31.10.2016']
#А что выбираем?!
choise = int(input('''
#1 - всего заявок
#2 - табель по заявкам
#3 - без изменений больше 2 суток по кол-ву заявок
#4 - без изменений больше 2 суток по ID
#5 - без изменений больше 14 суток по кол-ву заявок
#6 - без изменений больше 14 суток по ID
#7 - закрытие в течении 24 часов
#8 - без левых заявок закрытие в 24 часа
#9 - indicator
#10 - дубликаты
#11 - indicator
'''))

if choise == 1:
    ws = wb.add_sheet('общие')
    ws.write(1, 0, 'email')
    ws.write(2, 0, 'отзвон')
    ws.write(3, 0, 'очередь')
    ws.write(4, 0, 'всего')
    ws.write(5, 1, 'без владельца')
    ws.write(5, 6, 'на инженерах')
    ws.write(5, 11, 'на группе')
    ws.write(6, 0, 'email')
    ws.write(7, 0, 'отзвон')
    ws.write(8, 0, 'очередь')
    ws.write(9, 0, 'всего')
    for i in range(len(list)):
        oldtable(list[i])
if choise == 2:
    ws = wb.add_sheet('new')
    ws.write(1, 0, 'Срочнопорт')
    ws.write(2, 0, 'Консультация')
    ws.write(3, 0, 'Другое')
    ws.write(4, 0, 'NoLink')
    ws.write(5, 0, 'Диагностика')
    ws.write(6, 0, 'Не работает интернет')
    ws.write(7, 0, 'Низкая скорость')
    ws.write(8, 0, 'Дисконнекты (обрывается интернет)')
    ws.write(9, 0, 'SPAM')
    ws.write(10, 0, 'IPTV')
    ws.write(11, 0, 'Внутренние ресурсы')
    ws.write(12, 0, 'Деталька')
    ws.write(13, 0, 'Настройка')
    ws.write(14, 0, 'wi-fi')
    ws.write(15, 0, 'Видео')
    ws.write(16, 0, 'письма-спам')
    ws.write(17, 0, 'Не верная категория')
    for i in range(len(list)):
        newtable(list[i])
if choise == 3:
    ws = wb.add_sheet('>2')
    for i in range(len(list)):
        onemore(list[i])
if choise == 4:
    ws = wb.add_sheet('>2-ID')
    ws.write(1, 0, 'email')
    ws.write(2, 0, 'отзвон')
    ws.write(3, 0, 'очередь')
    ws.write(4, 0, 'всего')
    for i in range(len(list)):
        twoday(list[i])
if choise == 5:
    ws = wb.add_sheet('>14')
    for i in range(len(list)):
        fo14(list[i])
if choise == 6:
    ws = wb.add_sheet('>14id')
    fo14id('')
if choise == 7:
    ws = wb.add_sheet('<1')
    for i in range(len(listclosed)):
        oneday(listclosed[i])
        ws.write(0, i+4, listclosed[i])
if choise == 8:
    ws = wb.add_sheet('<1')
    for i in range(len(listclosed)):
        oneday(listclosed[i])
        ws.write(0, i+4, listclosed[i])
if choise == 9:
    ws = wb.add_sheet('<1')
    for i in range(len(listclosed)):
        indicator(listclosed[i])
        ws.write(0, i+5, listclosed[i])
if choise == 10:
    ws = wb.add_sheet('<1')
    dublicat(0)
if choise == 11:
    ws = wb.add_sheet('Dynamic')
    for i in range(len(list)):
        dynamic(list[i])
        ws.write(i+1, 0, listclosed[i])
