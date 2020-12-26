import pyautogui as pg
import time
import keyboard

long_sleep = 12
time.sleep(1)
pg.PAUSE = 1
pg.FAILSAFE = True

img_account = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\счет.png'
img_form = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\сформировать.png'
img_stop_form = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\F1.png'
img_filemame = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\имя файла.png'
img_filetype = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\тип файла.png'
img_save_button = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\сохранить.png'
img_xlsx = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\лист excel2007(xlsx).png'
img_alt_date = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\произвольный интервал.png'
#img_menu_file = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\меню файл.png'
#img_save_as = r'C:\Users\andrew_sapozhinsky\Desktop\прог\clicker\venv\Lib\image_raritek\пиктограмма сохранить.png'
#pg.alert(text='Переведите в русскую и нажмите ОК', title='Проверьте пожалуйста раскладку клавиатуры', button='OK')
#a = pg.prompt(text='4 цифры через пробел (н-р, 6201 5003)', title='Введите перечень счетов в формате' , default='')
# accounts = ['0101', '0103', '0104', '0105', '0109', '0201', '0203', '0204', '0205', '0401', '05', '0803', '0804', '0805', '09', '1003', '1006', '1007', '1009', '1010', '1011', '1012', '1901', '1902', '1903', '1904', '1905', '1909', '1910', '20011', '2601', '4101', '4105', '43', '44011', '4501', '5001', '5003', '51', '52', '5501', '5711', '5722', '58011', '5803', '6001', '6002', '6021', '6022', '6201', '6202', '6221', '6222', '6601', '6602', '6603', '6604', '6702', '6801', '6802', '68041', '68042', '6807', '6808', '6810', '6814', '6832', '6842', '69', '70', '7101', '7301', '7303', '7602', '7603', '7604', '7605', '7606', '7609', '7622', '7641', '76АВ', '76НА', '77', '8009', '8309', '8401', '90011', '90021', '9003', '90071', '90081', '9009', '9101', '91021', '9109', '96', '9721', '99011', '99021', '99022', '99023', '9909']
# accounts = ['01', '02', '0401', '05', '08', '0805', '09', '10', '19', '20011', '2601', '41', '43', '44011', '4501', '50', '51', '52', '5501', '57', '58011', '5803', '6001', '6002', '6021', '6022', '6201', '6202', '6221', '6222', '66', '67', '6801', '6802', '6804', '6807', '6808', '6810', '6814', '6832', '6842', '69', '70', '7101', '73', '7602', '7603', '7604', '7605', '7606', '7609', '7622', '7641', '76АВ', '76НА', '77', '8009', '8309', '8401', '90011', '90021', '9003', '90071', '90081', '9009', '9101', '91021', '9109', '96', '9721', '99011', '99021', '99022', '99023', '9909']
accounts = ['01', '02', '04', '05', '08', '0805', '09', '10', '19', '20011', '2601', '41', '43', '44011', '4501', '50', '51', '52', '5501', '57', '58', '60', '62', '66', '67', '68', '69', '70', '71', '73', '7602', '7603', '7604', '7605', '7606', '7609', '7622', '7641', '76АВ', '76НА', '77', '80', '83', '84', '90011', '90021', '9003', '90071', '90081', '9009', '9101', '91021', '9109', '96', '9721', '99']
#if a:
#    accounts.extend(a.split(' '))
#p = pg.prompt(text='Имя папки на рабочем столе', title='Введите путь для сохранения файлов' , default='')
#path = r'C:\Users\audit2\Desktop\дополнительно'.format(p) + r'\ '
path = r'C:\Users\audit2\Desktop\РМЗ 1 этап 2020' + r'\ '
y = '20'
month_date = [(f'010120{y}', f'310120{y}'), (f'010220{y}', f'280220{y}'), (f'010320{y}', f'310320{y}'),
              (f'010420{y}', f'300420{y}'), (f'010520{y}', f'310520{y}'), (f'010620{y}', f'300620{y}'),
              (f'010720{y}', f'310720{y}'), (f'010820{y}', f'310820{y}'), (f'010920{y}', f'300920{y}'),
              (f'011020{y}', f'311020{y}'), (f'011120{y}', f'301120{y}'), (f'011220{y}', f'311220{y}')]
quater_date = [(f'010120{y}', f'310320{y}'), (f'010420{y}', f'300620{y}'),
               (f'010720{y}', f'300920{y}'), (f'011020{y}', f'311220{y}')]
year = (f'010120{y}', f'311220{y}')
dates = 0  #если нужно, чтобы дата проставлялась в начале работы цикла - установить значение 1
report_type = pg.confirm(text='Выберите тип отчета', title='Тип отчета', buttons=['анализ счета', 'карточка счета', 'осв по счету'])

def coordinates():
    '''функция для определения координат некоторых кнопок'''
    d = {'account': 'счет', 'form': 'сформировать', 'change_date': 'выбор даты'}
    for i in ('account', 'form', 'change_date'):
        flag = pg.alert(text=f'Наведите курсор на кнопку {d.get(i)} и нажмите enter', title='Определение координат', button='OK')
        x, y = pg.position()
        globals()[i + '_x'] = x
        globals()[i + '_y'] = y


def change_account(a, period=year):
    '''функция для изменения номера счета и формирования отчета'''
    global dates
    timer = 0  #обнуляем таймер
    if dates:  #если перед этим использовались другие даты
        pg.moveTo(change_date_x, change_date_y, duration=0.25)  # (дата)переходим по координатам, нужно определить заранее
        pg.click()
        time.sleep(0.5)
        pg.moveTo(pg.locateOnScreen(img_alt_date), duration=0.25)    # (дата)переходим к смене даты
        pg.click()
        time.sleep(0.5)
        keyboard.press_and_release('tab')  # активное окно - дата начала
        keyboard.write(period[0], delay=0.1)  # печатаем период[0]
        time.sleep(0.5)
        keyboard.press_and_release('tab') # активное окно - дата окончания
        keyboard.write(period[1], delay=0.1)  # печатаем период[1]
        keyboard.press_and_release('enter')
        time.sleep(0.5)
        dates = 0  # обнуляем даты
    pg.moveTo(account_x, account_y, duration=0.25)  #(счет)переходим по координатам, нужно определить заранее
    pg.moveRel(50, 0, duration=0.25)  #перемещаемся чуть правей
    pg.click()  #клик
    keyboard.write(a, delay=0.1)
    keyboard.press_and_release('enter')
    keyboard.press_and_release('enter')

    pg.moveTo(form_x, form_y, duration=0.25)  # (сформировать)переходим по координатам, нужно определить заранее
    pg.click()  #клик
    pg.moveRel(200, 200, duration=0.25)  # перемещаемся чуть правей
    while not pg.locateOnScreen(img_stop_form):
        timer += 0.2  #пока не будет стоп-слова прибавляем 0.2 минуты
        time.sleep(long_sleep)  #пока на экране не будет стоп-слова ждем 12 сек
    if 4 < timer < 15:   #если отчет формировался от 4 до 15 минут
        globals()['dates'] =  4  #будем формировать поквартально
    elif timer > 15:
        globals()['dates'] = 12  #будем формировать помесячно


def change_date_save(a):
    '''функция для изменения дат, если отчет формировался слишком долго'''
    if dates:  #если есть дата начала и окончания
        time.sleep(0.5)
        for p in range(dates):  #для каждого периода в дате (4 или 12)
            if dates == 12:  #если дата == 12
                globals()['start'] = month_date[p][0]  #пользуемся словарем месяцы
                globals()['stop'] = month_date[p][1]
            elif dates == 4:  #если дата == 4
                globals()['start'] = quater_date[p][0]  #пользуемся словарем кварталы
                globals()['stop'] = quater_date[p][1]
            pg.moveTo(change_date_x, change_date_y, duration=0.25)  #(счет)переходим по координатам, нужно определить заранее
            pg.click()
            time.sleep(0.5)
            pg.moveTo(pg.locateOnScreen(img_alt_date), duration=0.25)  #(дата)переходим к смене даты
            pg.click()
            time.sleep(0.5)
            keyboard.press_and_release('tab')
            keyboard.write(start, delay=0.1)  #вводим начальную дату
            time.sleep(0.5)
            keyboard.press_and_release('tab')
            keyboard.write(stop, delay=0.1)  #вводим конечную дату
            keyboard.press_and_release('enter')
            time.sleep(0.5)
            pg.moveTo(form_x, form_y, duration=0.25)  # (сформировать)переходим по координатам, нужно определить заранее
            pg.click()  # клик
            pg.moveRel(200, 200, duration=0.25)  # перемещаемся чуть правей
            while not pg.locateOnScreen(img_stop_form):
                time.sleep(long_sleep)
            save_file(a)



def save_file(a, period=year):
    keyboard.press_and_release('ctrl+s')  #нажимаем ctrl+s
    time.sleep(1)
    pg.moveTo(pg.locateOnScreen(img_filemame), duration=0.25)  #ищем имя файла
    time.sleep(1)
    pg.moveRel(100, 0)  #переходим чуть правей
    time.sleep(1)
    pg.click()  #клик
    time.sleep(1)
    keyboard.press_and_release('ctrl+a')  #выделяем имя файла
    time.sleep(1)
    keyboard.press_and_release('backspace')  #удаляем все что было
    time.sleep(1)
    if dates:
        keyboard.write(f'{path[:-1]}{report_type} {a} с {start} по {stop}', delay=0.1)  # печатаем новое имя с датами
    else:
        keyboard.write(f'{path[:-1]}{report_type} {a}', delay=0.1)  #печатаем новое имя
    time.sleep(1)
    if not pg.locateOnScreen(img_xlsx):  #если текущий файл не xlsx
        pg.moveRel(0, 20, duration=0.25)  # переходим чуть ниже
        time.sleep(1)
        pg.click()
        time.sleep(1)
        pg.moveTo(pg.locateOnScreen(img_xlsx), duration=0.25)  # ищем xlsx
        time.sleep(1)
        pg.click()
        time.sleep(1)
    keyboard.press_and_release('enter')  #enter для сохранения
    while pg.locateOnScreen(img_save_button):
        time.sleep(long_sleep)




coordinates()
for n in accounts:
    change_account(n)
    if dates:
        change_date_save(n)
        continue
    save_file(n)
