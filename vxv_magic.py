# Генератор списков с условием "if"
lst =  [1,5,3,7,3,4,8,3,9,3,1,2,8,6,7,4,9]
lst2 = [i for i in lst if i >= 5]
# Генератор списков с условием "if / else"
lst2 = [i + 5 if i < 5 else i for i in lst]

# Сортировка вложенных списков по 1у и 2м ключам
# 1 ключ
response = sorted(response, key = lambda i: i['StageNumber'])
# 2 ключа
# Создаем новый список или перезаписываем старый
response = sorted(response, key = lambda i: (i['StageNumber'], i['GenplanNumber']))
# Сортируем существующий список, изменяя его
response.sort(key = lambda item: (item['StageNumber'], item['GenplanNumber']))

'''
Функция enumerate() вернет кортеж, содержащий отсчет от start и значение
'''
for i, val in enumerate(lst):
    print(f'№ {i} => {val}')

for i, val in enumerate(lst, start=1):
    print(f'№ {i} => {val}')

'''
Получение списка парных кортежей (number, value) 
(порядковый номер в последовательности, значение последовательности)
'''
seasons = ['Spring', 'Summer', 'Fall', 'Winter']
list(enumerate(seasons))
[(0, 'Spring'), (1, 'Summer'), (2, 'Fall'), (3, 'Winter')]

# можно указать с какой цифры начинать считать
list(enumerate(seasons, start=1))
[(1, 'Spring'), (2, 'Summer'), (3, 'Fall'), (4, 'Winter')]


'''----------------------------------------------------------------------
Использование enumerate() для нахождения индексов 
минимального и максимального значений в числовой последовательности:
'''
lst = [5, 3, 1, 0, 9, 7]
# пронумеруем список 
lst_num = list(enumerate(lst, 0))
# получился список кортежей, в которых 
# первый элемент - это индекс значения списка, 
# а второй элемент - само значение списка
lst_num
# [(0, 5), (1, 3), (2, 1), (3, 0), (4, 9), (5, 7)]

# найдем максимум (из второго значения кортежей)
tup_max = max(lst_num, key=lambda i : i[1])
tup_max
# (4, 9)
f'Индекс максимума: {tup_max[0]}, Max число {tup_max[1]}'
# 'Индекс максимума: 4, Max число 9'

# найдем минимум (из второго значения кортежей)
tup_min = min(lst_num, key=lambda i : i[1])
tup_min
# (3, 0)
f'Индекс минимума: {tup_min[0]}, Min число {tup_min[1]}'
# 'Индекс минимума: 3, Min число 0'

'''----------------------------------------------------------------------'''
'''Провеврка на тип данных'''
if isinstance(item, str):
    pass

'''----------------------------------------------------------------------'''


ui.tableWidget.setColumnWidth(0, 100)
ui.tableWidget.setColumnWidth(2, 170)
# header = ui.tableWidget.horizontalHeader()
# header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
# header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)     
# header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)

# ui.plainTextEdit_1 = QtWidgets.QPlainTextEdit(Form)
# ui.plainTextEdit_1.setPlaceholderText(_translate("Form", "XXX.XXX.XXX"))
# ui.plainTextEdit_1.setFrameShape(QtWidgets.QFrame.NoFrame)
# ui.tableWidget.setCellWidget(1, 2, ui.plainTextEdit_1)

ui.textEdit = QtWidgets.QTextEdit(Form)
ui.textEdit.setPlaceholderText(_translate("Form", "                 XXX.XXX.XXX"))
ui.textEdit.setFrameShape(QtWidgets.QFrame.NoFrame)
ui.textEdit.setAlignment(QtCore.Qt.AlignCenter)

ui.tableWidget.setCellWidget(1, 2, ui.textEdit)

'''----------------------------------------------------------------------'''
'''Как найти все повторяющиеся элементы в списке и количество повторов?'''
from collections import Counter
A = [111, 222, 111, 333, 222, 111]
counter = Counter(A)

'''----------------------------------------------------------------------'''
'''----------------------------------------------------------------------'''
'''Как работает Pathlib?'''
import pathlib
Path() — это дочерний узел PurePath(), он обеспечивает операции обработки с возможностью выполнения процесса записи.

# Когда вы инстанцируете Path(), он создает два класса для работы с путями Windows и отличных от Windows. 
# Как и PurePath(), Path() также создает общий объект пути "agnostic path", 
# независимо от операционной системы, в которой вы работаете.
# свойства PurePath() аналогичны с Path()

# Path().iterdir() возвращает содержимое каталога. Допустим, 
# у нас есть следующая папка, содержащая следующие файлы:
p = pathlib.Path('/data')

for child in p.iterdir():
    print(child)

# Полный путь текущего файла
p = pathlib.Path(__file__)
__file__    # встроеная переменная, служит ссылкой на путь к файлу, в котором мы ее прописываем
# тоже самое, но для работы в разных операционных системах
Pure = pathlib.PurePath(__file__)
# PurePath().parents[] выводит предков пути:
# чем больше индекс, тем дальше родитель
p = pathlib.PurePath('/src/goo/scripts/main.py')
# '/src/goo/scripts'
p.parents[0]
# '/src/goo'
p.parents[1]

PurePath().name     # предоставляет имя последнего компонента вашего пути:
# main.py
pathlib.PurePath('/src/goo/scripts/main.py').name 
# В свою очередь, PurePath().suffix предоставляет расширение файла 
# последнего компонента вашего пути:
pathlib.PurePath('/src/goo/scripts/main.py').suffix  # '.py'

# выводит только имя конечного компонента вашего пути без суффикса:
In [*]: pathlib.PurePath('/src/goo/scripts/main.py').stem                      
Out[*]: 'main'

PurePath().is_relative() проверяет, принадлежит ли данный путь другому заданному пути или нет:
In [*]: p = pathlib.PurePath('/src/goo/scripts/main.py')
        p.is_relative_to('/src')

PurePath().joinpath() конкатенирует путь с заданными аргументами (дочерними путями):

In [*]: p = pathlib.PurePath('/src/goo')
        p.joinpath('scripts', 'main.py')

Out[*]: PurePosixPath('/src/goo/scripts/main.py')

PurePath().match() проверяет, соответствует ли путь заданному шаблону:

In [*]: pathlib.PurePath('/src/goo/scripts/main.py').match('*.py')
Out[*]: True

In [*]: pathlib.PurePath('/src/goo/scripts/main.py').match('goo/*.py')
Out[*]: True

In [*]: pathlib.PurePath('src/goo/scripts/main.py').match('/*.py')
Out[*]: False

PurePath().with_stem() изменяет только имя последнего компонента пути:

In [*]: p = pathlib.PurePath('/src/goo/scripts/main.py')
        p.with_stem('app.py')

PurePath().with_suffix() временно изменяет суффикс или расширение последнего компонента пути:  

In [*]: p = pathlib.PurePath('/src/goo/scripts/main.py')
        p.with_suffix('.js')
Out[*]: PurePosixPath('/src/goo/scripts/main.js')        

Если имя заданного пути не содержит суффикса, метод .with_suffix() добавляет суффикс за вас:

In [*]: p = pathlib.PurePath('/src/goo/scripts/main')
        p.with_suffix('.py')
Out[*]: PurePosixPath('/src/goo/scripts/main.py')
'''----------------------------------------------------------------------'''