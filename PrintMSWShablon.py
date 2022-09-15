import win32com.client
import win32com.client.gencache
from VBAExcel import *
import VXVtranslittext
from time import sleep
# os.system("CLS")


def GO(ui, Form, sig, fail=None):
    print('---------------------------------------------------------')
    progressBar = ui.progressBar_1
    '''Создаем COM объект Excel'''
    try:
        Excel = win32com.client.GetActiveObject('Excel.Application')
    except:
        Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = 1
    wb = Excel.ActiveWorkbook
    sheet = wb.ActiveSheet

    sig.signal_Probar.emit(progressBar, 10)
    '''--------------------------------------------'''
    '''Находим номера крайней строки и столбца в таблице Excel'''
    EndRow, EndCol = EndIndexRowCol(sheet)
    '''Номер первой занимаемой строчки и столбца'''
    StartRow, StartCol = sheet.UsedRange.Row + 1, sheet.UsedRange.Column
    '''Выбираем таблицу в Excel'''
    print(f"StartRow = {StartRow}")
    tabEx = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    '''Копируем выбранный диапозон из Excel'''
    tabEx.Copy()
    sleep(1)
    sig.signal_Probar.emit(progressBar, 20)
    '''---------------------------------------------------------------'''
    """Создаем COM объект Word"""
    Word = win32com.client.Dispatch("Word.Application")
    # Word = win32com.client.gencache.EnsureDispatch("Word.Application")
    Word.Visible = 1
    StartRow, StartCol = 1, 1
    # sleep(1)
    strPath = ui.plainTextEdit_10.toPlainText()
    if strPath != '':
        strPath = strPath
    if strPath == '':
        # strPath = os.getcwd()
        strPath = wb.FullName.split(wb.Name)[0][:-1]
    nameobjextproekt = ui.plainTextEdit_4.toPlainText()
    if nameobjextproekt == '':
        text = f'Укажите ШИФР_ОБЪЕКТА'
        sig.signal_err.emit(Form, text)
        return

    """Добавляем документ по шаблону"""
    
    if ui.comboBox.currentIndex() == 1:
        fail = os.getcwd() + "\\Шаблон_ВОР.dotx"
    if ui.comboBox.currentIndex() == 2:
        # fail = os.getcwd() + "\\Шаблон_печати_А4_альбом.dotx"
        fail = os.getcwd() + "\\Шаблон_Спецификация оборудования, изделий и материалов.dotx"

    '''Сохранить как'''
    saveAsNameobjextproekt = VXVtranslittext.GO(nameobjextproekt)
    FileName = f"{strPath}\\{saveAsNameobjextproekt}.docx"

    '''Проверяем открыт ли одноименный файл *.docx и подключаемся к нему,
    иначе создаем новый файл'''
    prov = False
    for i in Word.Documents:
        if i.FullName == FileName:
            prov = True
            Doc = i
    if prov == False:
        Doc = Word.Documents.Add(fail)

    Doc.Activate()
    Doc.SaveAs(FileName)
    sleep(1)

    '''В последний параграф, идущий после таблицы 
    вставляем таблицу и объединяем ее с существующей'''
    CountP = Doc.Paragraphs.Count
    Selection = Doc.Paragraphs(CountP).Range
    # Selection.PasteAppendTable()

    # Selection.PasteAndFormat(16)
    Selection.Paste()

    '''Вторая строка отформатирована в ручную для передачи формата 
    вставленным строкам, ткперь ее можно удалить'''
    Doc.Tables(1).Rows(2).Delete()
    '''---------------------------------------------------------------'''
    sig.signal_Probar.emit(progressBar, 40)

    if Doc.Saved == False:
        try:
            Doc.Save()
        except:
            pass
    print("---END---")
    '''---------------------------------------------------------------'''


if __name__ == "__main__":
    import sys
    from ExcelToWorld import app, ui, Form, sig

    fail = os.getcwd() + "\\Шаблон_Спецификация оборудования, изделий и материалов.dotx"
    GO(ui, Form, sig, fail)
    sys.exit(app.exec_())
