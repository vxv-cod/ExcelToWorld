import os
import sys
from time import sleep
import win32com.client
import win32com.client.gencache
import threading
# from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
import traceback
from PyQt5 import QtCore, QtWidgets
import pickle
import VXVtranslittext
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx

from okno_ui import Ui_Form
from vxv_tnnc_SQL_Pyton import Sql
from version import ver
from VBAExcel import *
import PrintMSW
# import PrintMSWShablon

# from rich import print
# from rich import inspect
# inspect(xxx, methods=True)
# inspect(xxx, all =True)
# from prettytable import PrettyTable
# os.system('CLS')

app = QtWidgets.QApplication(sys.argv)
Form = QtWidgets.QWidget()
ui = Ui_Form()
ui.setupUi(Form)
Form.show()

_translate = QtCore.QCoreApplication.translate
Title = 'ExcelToWorld v. 1.0' + str(ver)
Form.setWindowTitle(_translate("Form", Title))


def startFun(my_func):
    """Обертка функции (декоратор)"""
    def wrapper():
        Sql("ExcelToWorld")
        pushButtonList = [ui.pushButton_4]
        progressBar = ui.progressBar_1
        label = ui.label
        try:
            for i in pushButtonList:
                sig.signal_bool.emit(i, True)
            sig.signal_label.emit(label, "Обработка данных . . .")
            sig.signal_Probar.emit(progressBar, 0)
            sig.signal_color.emit(progressBar, 0)
            my_func()
        except:
            errortext = traceback.format_exc()
            print(errortext)
            text = f"Ошибка работы, повторите попытку \n\n{errortext}"
            sig.signal_err.emit(Form, text)
        for i in pushButtonList:
            sig.signal_bool.emit(i, False)
        sig.signal_Probar.emit(progressBar, 0)
        sig.signal_color.emit(progressBar, 100)
        # sig.signal_label.emit(label, "Выполнено: таблица вставлена в документ Word . . .")

    return wrapper

# ui.tabWidget.setEnabled(False)


'''--------------------------------------------------------------------'''
'''Ручная простановка высоты формы'''
# Form.setMinimumHeight(500)

def handle_updateRequest(rect=QtCore.QRect(), dy=0):
    '''Изменение высоты plainTextEdit и окна'''
    for widgetX in widgetList:
        doc = widgetX.document()
        tb = doc.findBlockByNumber(doc.blockCount() - 1)
        h = widgetX.blockBoundingGeometry(tb).bottom() + 2 * doc.documentMargin()
        widgetX.setFixedHeight(h)

        eee = sum([widgetList[i].height() for i in range(len(widgetList) - 1)])
        ''' (25 * 4 = 100) + 100 = 300; 100 - максимальная высота последнего элемента с подставленным Vertical Spacer'''
        hhh = 4 * 25 + 100
        xxx = 0 if eee <= hhh else eee - hhh
        '''Сравниваем максимальные значения на разных вкладках'''
        rrr = widgetList[-1].height() - 25
        xxx = max(xxx, rrr)
    '''Корректируем высоту формы, если размеры widgetX больше допустимых'''
    Form.resize(Form.minimumWidth(), Form.minimumHeight() + xxx)

widgetList = [ui.plainTextEdit_4, ui.plainTextEdit_5, ui.plainTextEdit_6, ui.plainTextEdit_8, ui.plainTextEdit_9, ui.plainTextEdit_3]
for widget in widgetList:
    widget.updateRequest.connect(handle_updateRequest)

'''--------------------------------------------------------------------'''
'''Чистим "plainTextEdit" для отображения текста по умолчанию'''
# ui.plainTextEdit.clear()
ui.plainTextEdit_3.clear()
ui.plainTextEdit_4.clear()
ui.plainTextEdit_5.clear()
ui.plainTextEdit_6.clear()
ui.plainTextEdit_8.clear()
ui.plainTextEdit_9.clear()
ui.plainTextEdit_10.clear()

'''Отслеживаем сигнал закрытия окна и сохраняем все из окна перед закрытием'''
def AppQuit():
    doljList = []
    UserList = []
    for i in range(ui.tableWidget_1.rowCount()):
        dolj = eval(f'ui.tableWidget_1.item({i}, 0).text()')
        xxx = eval(f'ui.tableWidget_1.item({i}, 1).text()')
        doljList.append(dolj)
        UserList.append(xxx)
    saveData = [
                ui.plainTextEdit_3.toPlainText(),
                ui.plainTextEdit_8.toPlainText(),
                ui.plainTextEdit_9.toPlainText(),           
                doljList,
                UserList,
                ui.checkBox_3.isChecked(),
                ui.checkBox_2.isChecked(),
                ui.checkBox.isChecked(),
                ui.tabWidget.isEnabled(),
                ui.tableWidget_1.isEnabled(),
                ui.plainTextEdit_3.isEnabled(),
                ui.plainTextEdit_10.isEnabled(),
                ui.frame.isEnabled(),
                ui.checkBox.isEnabled()
                ]
    with open("saveData.ini", "wb") as f:
        pickle.dump(saveData, f) # помещаем объект в файл

app.aboutToQuit.connect(AppQuit)

"""Если файл с данными НЕ существует"""
savePathFile = os.getcwd() + "\saveData.ini"
if os.path.exists(savePathFile) == False:
    with open("saveData.ini", "wb") as file:
        pass

'''--------------------------------------------------------------------'''
'''Заполняем значения с данными в форму из файла после запуска программы'''
with open("saveData.ini", "rb") as f:
    try:
        loadx = pickle.load(f) # извлекаем ообъект из файла
        # print('loadx = ', loadx)
        ui.plainTextEdit_3.setPlainText(f"{loadx[0]}")
        ui.plainTextEdit_8.setPlainText(f"{loadx[1]}")
        ui.plainTextEdit_9.setPlainText(f"{loadx[2]}")
        for i in range(ui.tableWidget_1.rowCount()):
            eval (f'ui.tableWidget_1.item({i}, {0}).setText(_translate("Form", "{loadx[3][i]}"))')
            eval (f'ui.tableWidget_1.item({i}, {1}).setText(_translate("Form", "{loadx[4][i]}"))')
        ui.checkBox_3.setChecked(loadx[5])
        ui.checkBox_2.setChecked(loadx[6])
        ui.checkBox.setChecked(loadx[7])
        ui.tabWidget.setEnabled(loadx[8])
        ui.tableWidget_1.setEnabled(loadx[9])
        ui.plainTextEdit_3.setEnabled(loadx[10])
        ui.plainTextEdit_10.setEnabled(loadx[11])
        ui.frame.setEnabled(loadx[12])
        ui.checkBox.setEnabled(loadx[13])
    except:
        pass

'''Отслеживаем сигнал в plainTextEdit на изменение данных и удаляем не нужный текст'''
def ChangedPT(plainTextEdit):
    '''Удаления ненужного текста в plainTextEdit_3'''
    directory = plainTextEdit.toPlainText()
    if "file:///" in directory:
        xxx = directory.rfind("file:///")
        directory = directory[xxx + 8:]
        try:
            directory = directory.replace("/", "\\")
        except:
            pass
        plainTextEdit.setPlainText(rf"{directory}")
ui.plainTextEdit_3.textChanged.connect(lambda : ChangedPT(ui.plainTextEdit_3))
ui.plainTextEdit_10.textChanged.connect(lambda : ChangedPT(ui.plainTextEdit_10))

sig = Signals()

@thread
@startFun
def startWorld():
    PrintMSW.GO(ui, Form, sig)

# @thread
# def redactShablonWorld(fileName):
#     fail = os.getcwd() + f"\\{fileName}"
#     Word = win32com.client.gencache.EnsureDispatch("Word.Application")
#     Word.Documents.Open(fail)
# ui.pushButton_5.clicked.connect(redactShablonWorld)

ui.pushButton_4.clicked.connect(startWorld)

if __name__ == "__main__":
    # start()
    sys.exit(app.exec_())
