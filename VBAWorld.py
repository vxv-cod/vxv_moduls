

'''https:#club.directum.ru/post/778?ysclid=l6djdl35ao769763968'''

import os
import win32com.client
import win32com.client.gencache
from rich import print
from rich import inspect

fail = os.getcwd() + "\\Шаблон_печати_А4_альбом.dotx"

# Word = CreateObject("Word.Application") 
Word = win32com.client.Dispatch("Word.Application")
Word = win32com.client.gencache.EnsureDispatch("Word.Application")
Doc = Word.Documents.Open(fail)
Doc = Word.ActiveDocument

'''Закрыть документ'''
Doc.Close()
NameFaileDoc = 'test_result.docx'
for Docum in Word.Documents:
    if Docum.Name == NameFaileDoc:
        Docum.Close()
Word.Documents.Open(os.getcwd() + f"\\{NameFaileDoc}")

'''Создать по шаблону по шаблону'''
Doc = Word.Documents.Add(fail)

# Word.DisplayAlerts = -1     # Отображаются все поля сообщений и оповещения; ошибки возвращаются в макрос

'''Добавить текст'''
Doc.Paragraphs[2].Range.text = "12456"

'''Точек на сантиметр'''
PT = 28.34646  # количество "пт" в см

'''Установка единицы измерения размера таблицы, 
где 
    1 - Автоматический (сбрсмывает длину таблицы)
    2 - проценты
    3 - точки'''
Doc.Tables(1).PreferredWidthType = 1    # авто
Doc.Tables(1).PreferredWidthType = 2    # % от общей ширина между полями страницы
Doc.Tables(1).PreferredWidthType = 3    # CM

'''Установка ширины 1ой колонки таблицы, если строчки не объеденены (не рекомендуется)'''
Doc.Tables(1).Columns(1).PreferredWidth = 20
'''Установка ширины 1ой колонки таблицы, если строчки объеденены'''
Doc.Tables(1).Cell(1, 1).Range.Columns.PreferredWidth = 20
'''Процентное соотношение всей таблицы к стронице между полями'''
tabWord.PreferredWidth = 20
'''Установка ширины колонок всей таблице'''
tabWord.Range.Columns.PreferredWidthType = 2


'''Автоподбор размера ячейки таблицы
Фиксированная ширина = 0
По содержимому = 1
По ширине окна = 2
'''
Doc.Tables(1).AutoFitBehavior(0)    # фиксированный размер независимо от содержимого и не имеет автоматического размера
Doc.Tables(1).AutoFitBehavior(1)    # размер таблицы по содержимому
Doc.Tables(1).AutoFitBehavior(2)    # по ширине активного окна


'''=================================================================='''
'''------------------------------------------------------------------'''
'''При работе с таблицами в сантиметрах'''
PT = 28.34646  # количество "пт" в см
'''Установка единицы измерения размера таблицы''' 
tabWord = Doc.Tables(1)
tabWord.PreferredWidthType = 3
oriWidthCol_CM = [1.0, 6.5, 3.5, 3.2, 4.0, 2.8, 3.0, 2.2, 2.6]
WidthList = oriWidthCol_CM
'''Задаем общую ширину таблицы для более точного определения'''
tabWord.PreferredWidth = sum(WidthList) * PT
'''Проходим по всем колонкам для установки размеров из списка'''
for i in range(1, len(WidthList) + 1):
    '''Номер строки принимать последнюю строчку шапки, если ячейки в шапке объединенные'''
    col = tabWord.Cell(1, i).Range.Columns
    col.PreferredWidthType = 3
    col.PreferredWidth = WidthList[i-1] * PT
'''------------------------------------------------------------------'''
'''При работе с таблицами в %'''
# Запись размера таблицы в 80% от размера страницы без учета полей страницы
tabWord.PreferredWidthType = 2
tabWord.PreferredWidth = 80

'''Работа % стобцов таблицы'''
'''Запись ширин столбцов [40, 10, 10] эквивалентна - [4, 1, 1]'''
'''Смысл задания % заключен в отношении всех столбцов к наименьшему, 
из вышеописанного списка 1ый столбец в 4 раза больше 2го и 3го'''
# oriWidthCol_proc = [4, 1, 2] = [57.1, 14.2, 28.5]
# [400/9 =44.5, 300/9 =33.3, 400/9 = 22.2]
oriWidthCol_proc = [4, 3, 2]
WidthList = oriWidthCol_proc
for i in range(1, len(WidthList) + 1):
    col = tabWord.Cell(1, i).Range.Columns
    col.PreferredWidthType = 2
    col.PreferredWidth = WidthList[i-1]
'''------------------------------------------------------------------'''
'''=================================================================='''


'''Поля в ячейках таблицы'''
tabWord = Doc.Tables(1)
tabWord.TopPadding = 0
tabWord.BottomPadding = 0
tabWord.LeftPadding = 0.05
tabWord.RightPadding = 0
tabWord.Spacing = 0
tabWord.AllowPageBreaks = True
tabWord.AllowAutoFit = True

'''Установка высоты ячеек'''
# Значение константы HeightRule
# Размер, указанный в параметре RowHeigh, является точным = 2
# Размер, указанный в параметре RowHeigh, является минимальным = 1
# Автоматический подбор высоты строк (параметр RowHeigh игнорируется) = 0

Doc.Tables(1).Rows.HeightRule = HeightRule   # указывает на способ изменения высоты
Doc.Tables(1).Rows.Height = RowHeigh         # RowHeight указывает на новую высоту строки в пунктах.

'''Установка интервала перед и после абзаца в таблице
Единица измерения интервала пт.
'''
Doc.Tables(1).Range.ParagraphFormat.SpaceBefore = 6  #интервал перед
Doc.Tables(1).Range.ParagraphFormat.SpaceAfter = 6   #интервал после

'''Абзац отступ слева'''
Doc.Tables(1).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646
tabWord.Cell(4, 2).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646
'''Отступ первой строки'''
Doc.Tables(1).Range.ParagraphFormat.FirstLineIndent = 1.2 * 28.34646
tabWord.Cell(4, 2).Range.ParagraphFormat.FirstLineIndent = 1.2 * 28.34646

'''Установка интервалов перед и после абзаца
'''
Doc.Paragraphs(1).Format.SpaceBefore = 12 # интервал перед
Doc.Paragraphs(1).Format.SpaceAfter = 12  # интервал после

'''Межстрочный интервал
0.5 - одинарный интервал
1 – полуторный интервал и 1.5 – двойной интервал.
'''
Doc.Paragraphs(1).Format.LineSpacingRule = 0.5 # одинарный интервал

'''Установка стиля таблицы'''
Doc.Tables(1).Style = ИмяСтиляТаблицы

'''Установка отступа слева '''
Doc.Tables(1).Rows.LeftIndent = 0

'''Абзацный отступ (красная строка) абзаца'''
Doc.Paragraphs(1).Format.FirstLineIndent = Значение в пунктах

'''Выравниваем по вертикали все ячейки в таблице'''
Doc.Tables(1).Range.Cells.VerticalAlignment = 1
'''Удаляем интервал после абзаца во всей таблице'''
Doc.Tables(1).Range.ParagraphFormat.SpaceBefore = 0  # интервал перед
Doc.Tables(1).Range.ParagraphFormat.SpaceAfter = 0  # интервал после
'''Добавляем отступ в ячейке слева'''
Doc.Tables(1).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646

'''Выделяем колонку и делаем отступ в ячейках'''
tabWord.Columns(2).Select()
col = Doc.Application.Selection
col.Font.Color = 255
col.ParagraphFormat.LeftIndent = 0.1 * 28.34646

'''Выбор параграфа целиком'''
Doc.Paragraphs(Doc.Paragraphs.Count - 1).Range.Select()
Selection = Doc.Application.Selection
# Selection.MoveLeft()      # Сдвигаем курсор влево
Selection.Collapse(1)     # '''Схлапываем выделение параграффа в его начало'''
# Selection.Collapse(0)     # '''Схлапываем выделение параграффа в его конец'''
# Selection.Delete()        № Удаляем выбранное
Selection.TypeBackspace()


'''Для перевода сантиметров в пункты можно воспользоваться функцией CentimetersToPoints, 
тогда абзацный отступ в 1,5 см можно задать следующим образом:'''
Doc.Paragraphs(1).Format.FirstLineIndent = Word.CentimetersToPoints(1.5)

'''Установка левой и правой границ текста абзаца:'''
Doc.Paragraphs(1).Format.LeftIndent = 10    # отступ слева
Doc.Paragraphs(1).Format.RightIndent = 10   # отступ справа

'''Установка значений полей ячеек по умолчанию'''
Doc.Tables(1).TopPadding = 0     # верхнее
Doc.Tables(1).BottomPadding = 0  # нижнее
Doc.Tables(1).LeftPadding = 0    # левое      
Doc.Tables(1).RightPadding = 0   # правое

'''Выравнивание текста в таблице по горизонтали
по левому краю = 0
по центру = 1
по правому краю = 2
по ширине = 3
'''
Doc.Tables(1).Range.ParagraphFormat.Alignment = 3  # по ширине
# в ячейке
Doc.Tables(1).Cell(2, 2).Range.ParagraphFormat.Alignment = 2
'''Выравниевание колонки 0 - левый край'''
SelectCell = Doc.Tables(1).Columns(2).Select()
Selection = Doc.Application.Selection
Selection.ParagraphFormat.Alignment = 0

'''Выравнивание текстового абзаца:
Значения констант выравнивания такие же как в предыдущем пункте.
'''
Doc.Paragraphs(1).Format.Alignment = 0  # выравнивание по левому краю


'''Выравнивание текста в ячейке по вертикали
по верхнему краю = 0
по центру = 1
по нижнему краю = 3
'''
'''Выравниваем по вертикали ячейку'''
Doc.Tables(1).Rows(4).Cells(5).VerticalAlignment = 1
'''Выравниваем по вертикали строку'''
Doc.Tables(1).Rows(4).Cells.VerticalAlignment = 1
'''Выравниваем по вертикали столбец'''
Doc.Tables(1).Columns(4).Cells.VerticalAlignment = 1
'''Выравниваем по вертикали все ячейки в таблице'''
Doc.Tables(1).Range.Cells.VerticalAlignment = 1

'''Выбираем ячейку по номеру колонки и номера строки'''
fff = tabWord.Columns(2).Cells(4).Range
fff.Font.Color = 255



'''Установка размера шрифта таблицы'''
Doc.Tables(1).Range.Font.Size = 7

'''Установка цвета текста в ячейке'''
Doc.Tables(1).Cell(НомерСтроки, НомерСтолбца).Range.Font.Color = 255 

'''Выделение всего текста таблицы (жирным, курсивом, подчеркиванием)'''
Doc.Tables(1).Range.Font.Bold = True          # жирным
Doc.Tables(1).Range.Font.Italic = True        # курсивом
Doc.Tables(1).Range.Font.Underline  = True    # подчеркивание

'''Выделение или работа со строчкой по ячейке в ней'''
sheet.Cells(StartRow, col1).EntireRow.Delete()


'''Установка цвета подчеркивания'''
Doc.Tables(1).Cell(НомерСтроки, НомерСтолбца).Range.Font.UnderlineColor = 255

'''Установка темы шрифта таблицы'''
Doc.Tables(1).Range.Font.Name = "Arial"
# tabWord.Range.Font.Name = "Times New Roman"


'''Объединение ячеек'''
# # объединение первой и второй ячеек первой строки
Cell = Doc.Tables(1).Cell(1, 1)
Cell.Merge(Doc.Tables(1).Cell(1, 2))
'''Ячейка Текст жирный'''
cell3 = Doc.Tables(1).Cell(6, 3).Range.Font.Bold = True
'''Ячейка Текст курсив'''
cell3 = Doc.Tables(1).Cell(6, 3).Range.Font.Italic = False
'''Ячейка Текст масштаб по горизонтали (расстояние между буквами в слове)'''
Doc.Tables(1).Cell(6, 3).Range.Font.Scaling
'''Ячейка Цвет шрифта'''
Doc.Tables(1).Cell(6, 3).Range.Font.TextColor 
'''Ячейка Цвет подчеркнутой линии шрифта'''
Doc.Tables(1).Cell(6, 3).Range.Font.UnderlineColor  = 255
'''Ячейка подчеркнутая линии шрифта'''
Doc.Tables(1).Cell(6, 3).Range.Font.Underline = True


'''Вставка Excel таблицы в Word
Paragraph – номер параграфа, куда будет вставлена таблица из Excel.
'''
SelectionWord = Doc.Paragraphs(Paragraph).Range
SelectionWord.PasteExcelTable(True, False, False)
'''Вставляем в существующую таблицу с объедидением таблиц'''
# После шапки добавить в шаблон две пустые строчки, вставить данные в последнюю и предыдущую удалить, 
# а то форматирование шапки перетянется на таблицу
tabWord = Doc.Tables(1)
Selection = tabWord.Rows(4).Range
Selection.PasteAppendTable()
tabWord.Rows(3).Delete()

'''Перемещаемся на последний параграф'''
CountP = Doc.Paragraphs.Count
myRange = Doc.Paragraphs(CountP).Range
'''Добавить параграф'''
myRange.Paragraphs.Add()
'''Разрыв страницы'''
myRange.InsertBreak()
# без агрумента - разрыв страницы
# 8	  Разрыв столбца в точке вставки.
# 6	  Разрыв строки.
# 9	  Разрыв строки.
# 10  Разрыв строки.
# 7	  Разрыв страницы в точке вставки.
# 3	  Новый раздел без соответствующего разрыва страницы.
# 4	  Разрыв раздела, при котором следующий раздел начинается на следующей четной странице. Если разрыв раздела попадает на четную страницу, Word оставляет следующую нечетную страницу пустой.
# 2	  Разрыв раздела на следующей странице.
# 5	  Разрыв раздела, при котором следующий раздел начинается на следующей нечетной странице. Если разрыв раздела попадает на нечетную страницу, Word оставляет следующую четную страницу пустой.
# 11  Заканчивает текущую строку и переносит оставшийся текст под рисунок, таблицу или другой элемент. Текст продолжается со следующей пустой строки, которая не содержит таблицу, выровненную по левому или правому полю.


'''Вставляет указанный текст'''
myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
myRange.InsertAfter('text')     #   в конце диапазона или выбора
myRange.InsertBefore('text')    #   в начале диапазона или выбора

'''Разрыв связь'''
Doc.Fields.Unlink  

'''Удаление абзаца
Paragraph – номер параграфа, который нужно удалить.
'''
Doc.Paragraphs(Paragraph).Range.Delete

'''Удаляем строку в таблице'''
Doc.Tables(1).Rows(НомерСтроки).Select
Doc.Tables(1).Rows(НомерСтроки).Delete

'''Установка границ таблицы'''
# WdBorderType(
# Линия, обрамляющая диапазон сверху = -1
# Линия, обрамляющая диапазон слева = -2
# Линия, обрамляющая диапазон снизу = -3
# Линия, обрамляющая диапазон справа = -4
# Все горизонтальные линии внутри диапазона = -5
# Все вертикальные линии внутри диапазона = -6
# Линия по диагонали сверху – вниз = -7
# Линия по диагонали снизу – вверх = -8
# )
'''Границы ячеек таблицы'''
Table = Doc.Tables(1)
Table.Borders(WdBorderType).LineStyle = 4

'''Закрасить всю таблицу цветом'''
Doc.Tables(1).Shading.BackgroundPatternColor = 255 # заливка красным цветом

'''Закрасить ячейку цветом'''
Cell = Doc.Tables(1).Cell(1, 1)
Cell.Shading.BackgroundPatternColor = -687800525 # заливка желтым цветом

'''Установка ориентации страницы'''
Doc.Application.Selection.PageSetup.Orientation = 1 # альбомная
Doc.Application.Selection.PageSetup.Orientation = 0 # книжная

'''Установка полей страницы'''
Word.Application.Selection.PageSetup.LeftMargin = Word.CentimetersToPoints(2)    # левое поле
Word.Application.Selection.PageSetup.RightMargin = Word.CentimetersToPoints(2)   # правое поле
Word.Application.Selection.PageSetup.TopMargin = Word.CentimetersToPoints(2)     # верхнее поле
Word.Application.Selection.PageSetup.BottomMargin = Word.CentimetersToPoints(2)  # нижнее поле 

'''Добавление таблицы в документ'''
Doc.Tables.Add(Doc.Paragraphs(1).Range, 3, 5) # добавление таблицы из 5 столбцов и 3 строк в 1 абзац

'''Добавление строки в таблицу'''
Doc.Tables(1).Rows.Add

'''Добавление колонки в таблицу'''
Doc.Tables(1).Columns.Add

'''Добавление текста в ячейку'''
Doc.Tables(1).Cell(1,3).Range.Text = 'Текст, который добавляется в ячейку'

'''--------------------------------------------------------------------------'''
'''Пример вставки данных в таблицу'''
'''--------------------------------------------------------------------------'''
'''Коллекция всех нижних колонтитулов'''
Footers = Doc.Sections(1).Footers
'''Подключаемся к 1-ой таблице нижнего колонтитула на 1-ом листе'''
FootersTables_1 = Footers(1).Range.Tables(1)
'''Все ячейки в таблице выравниваем по вертикали по центру'''
FootersTables_1.Range.Cells.VerticalAlignment = 1
'''Подключаемся к 1-ой таблице нижнего колонтитула на 2-ом листе'''
FootersTables_2 = Footers(2).Range.Tables(1)
'''Все ячейки в таблице выравниваем по вертикали по центру'''
FootersTables_2.Range.Cells.VerticalAlignment = 1

def insertCellText(table, row, col, text):
    """Отправляем текст в ячейку таблицы"""
    table.Cell(row, col).Range.Text = text
    '''Microsoft Word визуально уменьшает размер текста, впечатаемого в ячейку, 
    чтобы он вписывался в ширину столбца. Для чтения и записи, Boolean.'''
    table.Cell(row, col).FitText = True
'''Отправляем данные в штамп на 1-ом листе'''
for i in range(len(doljList)):
    insertCellText(FootersTables_2, rowList[i], 2, doljList[i])
    if doljList[i] != '':
        insertCellText(FootersTables_2, rowList[i], 3, UserList[i])
        if len(UserList[i]) > 10:
            FootersTables_2.Cell(rowList[i], 3).FitText = True
        insertCellText(FootersTables_2, rowList[i], 5, now)

# '''Отправляем данные в штамп на 2-ом листе'''
# text = nameobjextproekt
# '''ШИФР_ОБЪЕКТА'''
# insertCellText(FootersTables_1, 1, 8, text)
# '''Отправляем данные в штамп на 1-ом листе'''
# '''ШИФР_ОБЪЕКТА'''
# insertCellText(FootersTables_2, 1, 8, text)
# '''НАИМЕНОВАНИЕ_ОБЪЕКТА'''
# text = ui.plainTextEdit_5.toPlainText()
# insertCellText(FootersTables_2, 3, 8, text)
# '''НАИМЕНОВАНИЕ_РАЗДЕЛА'''
# text = ui.plainTextEdit_6.toPlainText()
# insertCellText(FootersTables_2, 6, 6, text)
# '''Спецификация'''
# text = ui.plainTextEdit_9.toPlainText()
# insertCellText(FootersTables_2, 9, 6, text)
# '''Стадия'''
# text = ui.plainTextEdit_8.toPlainText()
# insertCellText(FootersTables_2, 7, 7, text)
'''--------------------------------------------------------------------------'''



'''Сохранение документа в pdf формат:
где:
    Path - полный путь и имя нового файла формата PDF,
    17 - значение Microsoft.Office.Interop.Word.WdExportFormat, указывающие, что сохранять документ в формате PDF,
    openAfterExport - Значение True используется, чтобы автоматически открыть новый файл, 
        в противном случае используется значение False.
    CreateBookmarks - значение указывает, следует ли экспортировать закладки и тип закладки. 
        Значение константы WdExportCreateBookmarks:
    wdExportCreateHeadingBookmarks = 1 - Создание закладки в экспортируемом документе для всех заголовком, 
        которые включают только заголовки внутри основного документа и текстовые поля не в пределах колонтитулов, 
        концевых сносок, сносок или комментариев.
    wdExportCreateNoBookmarks = 0 - Не создавать закладки в экспортируемом документе.
    wdExportCreateWordBookmarks = 2 - Создание закладки в экспортируемом документе для каждой закладки, 
        которая включает все закладки кроме тех, которые содержатся в верхнем и нижнем колонтитулах.
'''
Doc.ExportAsFixedFormat(Path, 17, openAfterExport, CreateBookmarks) 

'''Для выделения определенного текста в документе можно воспользоваться, примерно, следующим кодом:'''

Word = win32com.client.CreateObject("Word.Application")  
Doc = Word.Documents.Open('Путь к документу', True, True)
myRange = Doc.Content 
# Поиск текста для выделения
myRange.Find.Execute("Выделяемый текст", True)     
isFind = myRange.Find.Found 
while isFind:
    # Выделения текста цветом
    myRange.Font.ColorIndex = 3
    myRange.Find.Execute("Выделяемый текст", True)     
    isFind = myRange.Find.Found
# endwhile           
Word.Visible = True 


'''Свойство таблицы "Повторять как заголовок на каждой странице"'''
Table.Rows.HeadingFormat = False
Table.Rows(1).HeadingFormat = True

'''Повторять как заголовок на каждой странице'''
Doc.Tables(1).Cell(StartRow, 1).Range.Rows.HeadingFormat = True

'''Пример для удаления после разрыва страницы (до разрыва находится 
Закладка "СписокРассылки") страницы без содержания:'''
Word = win32com.client.CreateObject("Word.Application")  
Doc = Word.Documents.Open('Путь к документу', True, True)
CountTabl = Doc.Tables.Count           
# Удаление двух таблиц 6-ой и 5-ой
Doc.Tables.Item(CountTabl - 1).Delete()
Doc.Tables.Item(CountTabl - 1).Delete()                     
# Удаляем последний параграф с данными
CountP = Doc.Paragraphs.Count           
Doc.Paragraphs(CountP).Range.Delete()
# Удаляем строки после Закладки
if Doc.Bookmarks.Exists("СписокРассылки"):         
    Doc.Bookmarks("СписокРассылки").Range.Delete
    Doc.Bookmarks("СписокРассылки").Range.Delete
    Doc.Bookmarks("СписокРассылки").Range.Delete                           
#Doc.Bookmarks("\Line").Range.Delete           
#Doc.Paragraphs(CountP - 2).Range.Delete

'''Сохранение файла по имени'''
Doc.save('table.docx')

'''Сохранить как'''
Doc.SaveAs(FileName = "generated.docx")

'''Если файл изменен, то сохранить его'''
if Doc.Saved == False: Doc.Save()



'''Выбор объектов'''
tabWord = Doc.Tables(1)
Selection = tabWord.Select()
Selection = Word.Selection.SelectCell

'''Обтекание таблиц'''
tabWord = Doc.Tables(1)
tabWord.Rows.WrapAroundText = True


'''Нижний колонтитул'''
Footers = Doc.Sections(1).Footers
FootersCount = Footers.Count
Footers_2_Tables = Footers(2).Range.Tables
FootersTabCount = Footers_2_Tables.Count
FootersTables_2 = Footers_2_Tables(1)
FootersTables_2.Range.Cells.VerticalAlignment = 1

'''Верхний колонтитул'''
Headers = Doc.Sections(1).Headers
HeadersCount = Headers.Count
HeadersTables = Headers(1).Range.Tables
HeadersTabCount = HeadersTables.Count
HeadersTab_1 = HeadersTables(1)



def handle_updateRequest(rect=QtCore.QRect(), dy=0):
    '''Изменение высоты plainTextEdit и окна'''
    for widgetX in widgetList:
        doc = widgetX.document()
        tb = doc.findBlockByNumber(doc.blockCount() - 1)
        h = widgetX.blockBoundingGeometry(tb).bottom() + 2 * doc.documentMargin()
        widgetX.setFixedHeight(h)

    eee = sum([i.height() for i in widgetList])
    ''' если бы было 4 элемента, то они бы были со следующими размерами: 25 + 25 + 25 + 60 = 135; 60 - высота удаленного 4го элемента'''
    xxx = 0 if eee <= 135 else eee - 135
    Form.resize(Form.minimumWidth(), Form.minimumHeight() + xxx)
    

widgetList = [ui.plainTextEdit_4, ui.plainTextEdit_5, ui.plainTextEdit_6]
for widget in widgetList:
    widget.updateRequest.connect(handle_updateRequest)


'''Вставляем новое пользовательское свойство документа 
(нельзя вставить свойство с таким же именем, будет ошибка)'''
name = "vxv_7"
value = "Value"
try:
    Doc.CustomDocumentProperties.Add(name, False, 4, value)
except:
    pass

UserProp = Doc.CustomDocumentProperties(name)
'''Назначить новое имя и значение поля'''
NameUserProp = UserProp.Name = "EEEEEEEEEEE"
NameUserProp = UserProp.Value = "6666666"
'''Прочитать имя и значение поля'''
NameUserProp = UserProp.Name
NameUserProp = UserProp.Value
'''Обновить поля'''
Doc.Fields.Update()
'''Удаляем все пользовательские поля свойств'''
for i in Doc.CustomDocumentProperties:
    i.Delete()

def ListCDPPrint():
    '''Вылеляем весь контент, по сути все что есть в документе'''
    myRange = Doc.Content
    for i in Doc.CustomDocumentProperties:
        '''Вставляем параграфф'''
        myRange.InsertParagraphAfter()
        '''Вставляем текст'''
        myRange.InsertAfter(f'{i.Name} & = ')
        '''Выбираем последний параграфф'''
        Selection = Doc.Paragraphs[Doc.Paragraphs.Count - 1].Range
        '''Схлапываем выделение параграффа в его конец'''
        Selection.Collapse(0)
        '''Вставляем кастомное поле свойства документа пользователя'''
        Selection.Fields.Add(Selection, -1, f"DOCPROPERTY  {i.Name}", True)
    '''Обновляем весь документ, в том числе и поля свойств'''
    Doc.Fields.Update()

def setCDP(cdpListName):
    '''Добавляем поля из списка'''
    for i in range(len(cdpListName)):
        try:
            Doc.CustomDocumentProperties.Add(cdpListName[i], False, 4, cdpListName[i])
        except:
            pass

def getCDP():
    '''Читаем поля и их значения'''
    cdpName = [i.Name for i in Doc.CustomDocumentProperties]
    cdpValue = [i.Value for i in Doc.CustomDocumentProperties]
    return cdpName, cdpValue

def DelCDP():
    '''Удаляем все поля свойст документа'''
    for i in Doc.CustomDocumentProperties:
        i.Delete()


def UpdateCDP():
    '''Обновляем поля свойств в основной области с текстом'''
    Doc.Fields.Update()
    '''Обновляем поля свойств в нижнем колонтитуле'''
    Footers = Doc.Sections(1).Footers
    for i in range(1, Footers.Count + 1):
        Footers(i).Range.Fields.Update()

'''Повторять как заголовок на каждой странице'''
tabWord.Rows(StartRow).HeadingFormat = True

def OriForm():
    '''Определение ориентации и формата (размера) листа'''
    # OrientationLista, FormatLista = OriForm()
    OrientationLista = Doc.PageSetup.Orientation    # 1 - альбомная, 0 - книжная
    FormatLista = Doc.PageSetup.PaperSize           # 6 - формат А3, 7 - формат А4
    return OrientationLista, FormatLista