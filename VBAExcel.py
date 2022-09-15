import os
from time import sleep
import win32com.client
import threading
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
from PyQt5 import QtCore, QtWidgets
# from vxv_tnnc_SQL_Pyton import Sql
# import traceback

# from rich import print
# from rich import inspect
# inspect(xxx, methods=True)
# inspect(xxx, all =True)
# from prettytable import PrettyTable
# os.system('CLS')

class Signals(QtCore.QObject):
    '''
    sig.signal_Probar.emit(ui.progressBar_1, 100)
    sig.signal_label.emit(ui.label, "Выполнено")
    sig.signal_err.emit(f"Ошибка работы, повторите попытку \n\n{traceback.format_exc()}")
    sig.signal_color.emit(ui.progressBar_1, 0)
    sig.signal_color.emit(ui.progressBar_1, 1)
    sig.signal_bool.emit(ui.pushButton, True)
    sig.signal_bool.emit(ui.pushButton, False)
    '''
    signal_Probar = QtCore.pyqtSignal(QtWidgets.QWidget, int)
    signal_label = QtCore.pyqtSignal(QtWidgets.QWidget, str)
    signal_err = QtCore.pyqtSignal(QtWidgets.QWidget, str)
    signal_bool = QtCore.pyqtSignal(QtWidgets.QWidget, bool)
    signal_color = QtCore.pyqtSignal(QtWidgets.QWidget, int)

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.signal_Probar.connect(self.on_change_Probar,QtCore.Qt.QueuedConnection)
        self.signal_label.connect(self.on_change_label,QtCore.Qt.QueuedConnection)
        self.signal_err.connect(self.on_change_err,QtCore.Qt.QueuedConnection)
        self.signal_bool.connect(self.on_change_bool,QtCore.Qt.QueuedConnection)
        self.signal_color.connect(self.on_change_color,QtCore.Qt.QueuedConnection)

    '''Отправляем сигналы в элементы окна'''
    def on_change_Probar(self, s1, s2):
        '''Значение процента в прогресбаре'''
        s1.setValue(s2)
    def on_change_label(self, s1, s2):
        '''Отправляем текст в label'''
        s1.setText(s2)
    def on_change_err(self, s1, s2):
        '''Сообщение об ошибке'''
        QtWidgets.QMessageBox.information(s1, 'Сбой программы...', s2)
    def on_change_color(self, s1, s2):
        '''Устанавливаем цвет прогресбара'''
        if s2 == 1:
            color = "170, 170, 170"
        else:
            color = "100, 150, 150"
        s1.setStyleSheet("QProgressBar::chunk {background-color: rgb("f"{color}); margin: 2px;""}")
    def on_change_bool(self, s1, s2):
        s1.setDisabled(s2)

# sig = Signals()

import ctypes
def Allobject():
    '''Выясняем сколько объкетов Excel во всех экземлярах открыто'''
    EnumWindows = ctypes.windll.user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    GetWindowText = ctypes.windll.user32.GetWindowTextW
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible
    titles = []
    countExelList = []
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buff, length + 1)
            titles.append((hwnd, buff.value))
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    for i in range(len(titles)):
        if "- Excel" in  titles[i][1]:
            countExelList.append(1)
    countfail = sum(countExelList)
    return countfail

def thread(my_func):
    '''Обертка функции в потопк (декоратор)'''
    def wrapper():
        global thr
        # threading.Thread(target=my_func, daemon=True).start()
        thr = threading.Thread(target=my_func, daemon=True).start()
    return wrapper


def ExcelInstances():
    '''Поиск всех процессов EXCEL.EXE'''
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(".", "root\cimv2")
    colExcelInstances = objSWbemServices.ExecQuery(
            f"SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
    return colExcelInstances

def decorExcel(my_func):
    '''Обертка функции в экземпляр Excel (декоратор)'''
    def wrapper():
        global objWorkbook, countfail
        countfail = Allobject()
        colExcelInstances = ExcelInstances()
        for objInstancei in colExcelInstances:
            objExcel = win32com.client.Dispatch("Excel.Application")
            for objWorkbook in objExcel.Workbooks:
                my_func()
            objExcel.Quit()
            objInstancei.Terminate
            sleep(2)
    return wrapper


'''Пример как исмользовать декоратор decorExcel'''
@decorExcel
def myfunc():
    WbName = objWorkbook.Name
    print(f"WbName = {WbName}")



'''-------------------------------------------------------------------------------------------------------------------------'''
'''-------------------------------------------------------------------------------------------------------------------------'''

def Book(fail=None, sheetName=None, ExcelVisible=1):
    '''Подключаемся к Excel'''
    pythoncomCoInitializeEx(0)
    try:
        Excel = win32com.client.GetActiveObject('Excel.Application')
    except:
        Excel = win32com.client.Dispatch("Excel.Application")
        '''Статическое подключение (работатет при копировании леистов из файла в файл)'''
        # Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    Excel.Visible = ExcelVisible
    
    CountBook = Excel.Workbooks.Count
    # print(f"CountBook = {CountBook}")
    if CountBook == 0:
        print("Нет открытых файлов Excel")
        wb, sheet, Namebook = '', '', ''
    else:
        if fail == None:
            '''Получаем доступ к активной книге'''
            wb = Excel.ActiveWorkbook
        else:
            '''Получаем доступ к определенному файлу'''
            wb = Excel.Workbooks.Open(fail)
        if sheetName == None:
            """Получаем доступ к активному листу"""
            sheet = wb.ActiveSheet
        else:
            """Получаем доступ к определенному листу"""
            sheet = wb.Worksheets(sheetName)
            sheet.Activate()
        Namebook = wb.Name
    return wb, sheet


# def printTabconsole(dataAll:list = [[1,2,3], [4,5,6]], TitleTab = '', align = "c", column = 0, add_column = False):
#     '''Печатаем список из списков в таблице в консоле
#     printTabconsole(dataAll = aaa, TitleTab = 'bbb', align = 'lrc', column = ccc)
#     '''
#     from prettytable import PrettyTable
#     mytable = PrettyTable()

#     if add_column == False:
#         if TitleTab == '':
#             TitleTab = ["-- " + str(i) + " --" for i in range(1, len(dataAll[0]) + 1)]
#         else:
#             TitleTab = ["-- " + str(i) + " --" for i in TitleTab]
#         mytable.field_names = TitleTab
#         mytable.add_rows(dataAll)

#     if add_column == True:
#         TitleTab = ["-- " + str(i) + " --" for i in range(1, len(dataAll) + 1)]
#         for i in range(len(dataAll)):
#             mytable.add_column(TitleTab[i], dataAll[i])

#     if column != 0:
#         mytable.align[TitleTab[column - 1]] = align
#     else:
#         mytable.align = align

#     return print(mytable)


def ifErr(formula):
    '''Убираем ошибку при пустом значении'''
    iferror = f"IFERROR({formula},\"\")"
    text = f"=IF({iferror}=0,\"\",{iferror})"
    # print(text)
    return text

'''Формулы (при вставке значений по вертикали использовать кортежи с пустым 2ым значением)
пример: ("=формула", )"'''
# formula1 = ifErr("RC[-2]-RC[-1]")

def EndIndexRowCol(sheet):
    # EndRow, EndCol = EndIndexRowCol(sheet)
    '''Определяем позиции первой и последней ячейки'''
    UsedRange = sheet.UsedRange
    # '''Количество занимаемых таблицей строк'''
    count_row = UsedRange.Rows.Count
    # '''Количество занимаемых таблицей колонок'''
    count_col = UsedRange.Columns.Count
    # '''Номер первой занимаемой строчки'''
    StartRow = UsedRange.Row
    # '''Номер первой занимаемой колонки'''
    StartCol = UsedRange.Column
    # '''Номер последней занимаемой строчки'''
    EndRow = StartRow + count_row - 1
    # '''Номер последней занимаемой колонки'''
    EndCol = StartCol + count_col - 1
    return EndRow, EndCol



def NameEndCell(sheet):
    # NameEndColumn, NameEndRow = NameEndCell(sheet)
    '''Адресс последней ячейки'''
    UsedRange = sheet.UsedRange
    Address = UsedRange.Address.split(":")[1]
    NameEndColumn = Address[1 : Address.rfind("$")]
    NameEndRow = Address[Address.rfind("$") + 1 :]
    # NameEndCell = NameEndColumn + NameEndRow
    return NameEndColumn, NameEndRow

def RangeCells(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Выделяем диапозон ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    return cel

def importdata(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Собираем данные из диапозона ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    vals = cel.Formula
    if StartCol == EndCol:
        vals = [vals[i][x] for i in range(len(vals)) for x in range(len(vals[i]))]
    return vals

def exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol):
    '''Отправляем данные в диапозон ячеек'''
    if StartCol == EndCol:
        data = [(i, None) for i in data]
    sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol)).Formula = data

def grani(cel):
    '''Все грани тонкие в диапазоне'''
    cel.Borders.Weight = 2
    # cel.Borders.ColorIndex = 0
    # cel.Borders.Weight = 3          # Все границы
    # cel.Borders(1).Weight = 3       # Левая граница
    # cel.Borders(2).Weight = 3       # Правая граница
    # cel.Borders(3).Weight = 3       # Верхняя граница
    # cel.Borders(4).Weight = 3       # Нижняя граница
    # cel.Borders(5).Weight = 3       # Диагональ граница
    # cel.Borders(6).Weight = 3       # Диагональ граница
    # cel.Borders(7).Weight = 3       # Левая граница крайних ячеек в диапозоне
    # cel.Borders(8).Weight = 3       # Верхняя граница крайних ячеек в диапозоне
    # cel.Borders(9).Weight = 3       # Нижняя граница крайних ячеек в диапозоне
    # cel.Borders(10).Weight = 3      # Правая граница крайних ячеек в диапозоне
    # cel.Borders(11).Weight = 3      # Вертикальные внутренние границы ячеек в диапозоне
    # cel.Borders(12).Weight = 3      # Горизонтальные внутренние границы ячеек в диапозоне

def PatchFail(widgetText):
    strPath = str(widgetText)
    if "file:///" in strPath:
        strPath = strPath[8:]
    if strPath == '':
        sig.signal_err.emit(f"Не указана папка для сохранения файлов")
        return
    return strPath.replace("/", "\\")

def exportPDF(widgetText, objWorkbook):
    '''Экспорт в PDF'''
    strPath = PatchFail(widgetText)
    pdfName = objWorkbook.Name if ".xls" not in objWorkbook.Name else objWorkbook.Name.split(".xls")[0]
    # sheet.PrintOut(Copies=1, ActivePrinter="Microsoft Print to PDF", PrintToFile=True, PrToFileName = f"{strPath}\\{pdfName}.pdf")
    OutputFile = f"{strPath}\\{pdfName}.pdf"
    objWorkbook.ExportAsFixedFormat(0, OutputFile)

# fileName = "Автомобильные дороги.xltx"
def redactExcel(fileName):
    '''Редактировать шаблон Excel *.xltx '''
    ZeroExcel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    ZeroExcel.Workbooks.Open(os.getcwd() + f"\{fileName}", Editable=True)
    ZeroExcel.Visible = 1

'''Округление'''
def NFt(cells, okrug):
    try:
        cells.NumberFormat = okrug
    except:
        cells.NumberFormat = okrug.replace('.', ',')


def infoCellMy(cel):
    dirCell = dir(cel)
    for i in dirCell:
        try:
            xxx = f"cel.{i}()"
            sss = eval(xxx)
            print(f"{xxx} = {sss}")
        except:
            try:
                xxx = f"cel.{i}"
                sss = eval(xxx)
                print(f"{xxx} = {sss}")
                pass
            except:
                # print(f'///cel.{i} {type((i))} - не обработано................')
                pass


'''================================================================================================'''
'''================================================================================================'''

# wb, sheet = Book()
# NameEndColumn = NameEndCell(sheet)[0]
# celX = f"=ПКЗМ!{NameEndColumn}8"
# print(celX)
# cel = sheet.Range("D30")
# cel.Formula = f"=ПКЗМ!G8"
# cel.Formula = "=ПКЗМ!R[-22]C[3]"



# inspect(sheet.Columns)

'''Сортировка таблицы'''
# sheet.Sort.SortFields.Clear
# sheet.Sort.SortFields.Add(Key=sheet.Range("E3"))
# sor = sheet.Sort
# sor.SetRange(sheet.Range(sheet.Cells(3, 1), sheet.Cells(EndRow, EndCol)))
# sor.Apply()

'''Зачеркиваем текст в ячейке'''
# cel.Font.Strikethrough = True

'''Цвет текста в ячейке'''
# cel.Font.Color = 255

'''Сохранить как'''
# objWorkbook.SaveAs(f"{strPath}\\{objWorkbook.Name}{strFileExtension}", FileFormat=objWorkbook.FileFormat, CreateBackup=0)

'''Выбрать несколько строчки'''
# RowsSelect = sheet.Rows(f"{StartRow}:{EndRow}")
# sheet.Range(sheet.Rows(7), sheet.Rows(18)).Select()
# cel = Excel.Selection.Interior
# cel.Color = 65535     # желтый
# cel.Color = 255       # красный
# cel.Pattern = 1
'''Выбор несколько колонок'''
# ColSelect = sheet.Columns("A:D")
# sheet.Range(sheet.Columns(2), sheet.Columns(5)).Activate()
# sheet.Range(sheet.Columns(2), sheet.Columns(5)).Select()


'''Объединение ячеек'''
# cel.Merge()
# cel.MergeCells = True
# cel.UnMerge

'''Перенести текст'''
# cel.WrapText = True

'''Задание ширины ячейки (колонки)'''
# cel.ColumnWidth = 45
'''Задание высоты ячейки (строки)'''
# cel.RowHeight = 45

'''Выравниваем текст в ячейке'''
'''Выравниваем по горизонтали (вертикали) центр'''
# cel.HorizontalAlignment = 3
# cel.VerticalAlignment = 2
'''Выравниваем по горизонтали влево'''
# cel.HorizontalAlignment = 1

'''Удаляем строки '''
# sheet.Rows(11).Delete()
# sheet.Rows("11").Delete()
# sheet.Rows("11:12").Delete()
# sheet.Rows(f"{StartRow}:{EndRow}").Delete()

'''Выделение или работа со строчкой по ячейке в ней'''
'''Удалить строчку по ячейке'''
# sheet.Cells(1, 1).EntireRow.Delete()

'''Удаляем колонки'''
# sheet.Columns(7).Delete()
# sheet.Columns("AV").Delete()
# sheet.Columns("AV:AW").Delete()
# sheet.Range(sheet.Columns(2), sheet.Columns(5)).Delete()

'''Очистить содержимое строчек'''
# sheet.Rows(f"{StartRow}:{EndRow}").ClearContents()

'''Копируем лист'''
# sheet.Copy(After=wb.Worksheets[wb.Worksheets.Count])

'''Копируем ячейки'''
# cel = RangeCells(sheet, StartRow, StartCol, EndRow, EndCol)
# cel.Copy()

'''Вставить ячейки'''
# cel = RangeCells(sheet, StartRow, StartCol, EndRow, EndCol)
# cel.Activate()
# sheet.Paste()

'''Подчеркнутый текст в ячейке'''
# cel.Font.Underline = 2

"""Отключение уведомлений с ответом по умолчанию для сохранения без подтверждения"""
# Excel.DisplayAlerts = False

'''Закрыть файл без сохранения'''
# wb.Close(False)
'''Закрыть файл с сохранением'''
# wb.Close()
'''Закрыть экземпляр Excel'''
# Excel.Quit()

'''
Сору    	                    Копирует объект Shape в буфер обмена
Cut	                            Копирует объект Shape в буфер обмена с удалением
Delete	                        Удаляет объект Shape
    img.Delete()
Paste	                        Вставляет объект Shape из буфера обмена
IncrementLeft, IncrementTop	    Сдвигает объект Shape по горизонтали и вертикали соответственно на заданное в аргументе количество пунктов. Синтаксис:
  IncrementLeft (Increment) 
  IncrementTop (Increment)
IncrementRotation	            Поворачивает объект Shape на заданный в аргументе угол. Синтаксис:
  IncrementRotation (Increment)
  '''
  
'''Сдвигает объект Shape по горизонтали и вертикали соответственно на заданное в аргументе количество пунктов'''
# img.IncrementLeft(100)
# img.IncrementTop(100)
'''
Left, Top — координаты левого верхнего угла объекта;
Width, Height — ширина и высота объекта.
'''
'''Привязываемся к координатам ячейки'''
# cell = sheet.Cells(5, 3)
# img.Left = cell.Left                                                                                                                                                               
# img.Top = cell.Top



'''True, False = 1, 0'''

'''Масштаб'''
# img.ScaleHeight(2, 1, 0)

# bbb = img.ScaleWidth(1, 1, 0)
'''Сохранить пропорции рисунка'''
# img.LockAspectRatio = True
'''НЕ охранять пропорции рисунка'''
# img.LockAspectRatio = False
# img.Height = 50
# img.Width = 100

# dirCell = dir(cel)
# for i in dirCell:
#     try:
#         xxx = f"cel.{i}()"
#         sss = eval(xxx)
#         print(f"{xxx} = {sss}")
#     except:
#         try:
#             xxx = f"cel.{i}"
#             sss = eval(xxx)
#             print(f"{xxx} = {sss}")
#             pass
#         except:
#             # print(f'///cel.{i} {type((i))} - не обработано................')
#             pass

# '''Удаляем колонки'''
# sleep(2)
# NameEndColumn, NameEndRow = NameEndCell(sheet)
# col = sheet.Columns(f"D:{NameEndColumn}")
# col.Delete()
# # print(EndCol)
# # for i in range(4, EndCol + 1):
# #     sheet.Columns(4).Delete()
# '''Удаляем строки со сдвигом вверх'''
# sheet.Rows(f"11:{NameEndRow}").Delete(1)