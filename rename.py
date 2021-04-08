import os #библиотека работы с проводником
import psutil #для работы с процессами
import openpyxl #для работы с таблицами
#основная функция
def find_and_kill_procs(name): #поиск и завершение процесса
    for p in psutil.process_iter(['pid','name']):
        try:
            if p.info['name'] == name:
                psutil.Process(p.info['pid']).kill()
                print(f"Процесс {p.info['name']} завершен")
        except psutil.NoSuchProcess:
            pass

def rename_list_xlsx(name):
    xls = openpyxl.load_workbook(name) #открываем файл
    xls_list = xls[xls.sheetnames[0]] #смотрим название листа и открываем
    xls_list.title = f"{num:g}" #переименовываем
    xls.save(name) #сохраняем
    return name

old_name_first = input("Введите полное название файла без расширения: ") #имя создаваемого файла
print("OK")
old_name_last = input("Введите расширение файла с точкой: ") #расширение файла
print("OK")
old_name = f"{old_name_first}{old_name_last}"
const = input("Введите текстовую часть названия нового файла: ") #постоянный текст в новом название
num = float(input("Введите начальное значение: ")) #числовая составляющая
dif = float(input("Введите шаг: ")) #шаг изменения числовой составляющей
flag = True #флаг ошибки
while flag: #бесконечный цикл
    try: #проверяем исключения
        new_name = f"{const} {num:g}{old_name_last}" #собираем новое название файла
        #for fail in os.listdir():
        #    if fail == new_name:
        #        print('Обнружены совпадающие имена')
        #        flag = False #флаг ошибки
        os.rename(old_name, new_name) #переименовываем
        find_and_kill_procs('EXCEL.EXE') #отключаем процесс excel
        print("Файл переименован на ", new_name)
        rename_list_xlsx(new_name)
        print(f"Лист в файле {rename_list_xlsx(new_name)} переименован на {num:g}")
        num += dif #прибавляем шаг
    except:
        continue