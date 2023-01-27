import win32com.client
import win32com.client.gencache
import imageZeroFon
from VBAExcel import *
import time
import VXVtranslittext
from time import sleep
# os.system("CLS")

def importdata(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Собираем данные из диапозона ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    vals = cel.Formula
    if StartCol == EndCol:
        vals = [vals[i][x] for i in range(len(vals)) for x in range(len(vals[i]))]
    return vals


def GO(ui, Form, sig):
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
    StartRow, StartCol = sheet.UsedRange.Row, sheet.UsedRange.Column
    count_col = sheet.UsedRange.Columns.Count
    '''Выбираем таблицу в Excel'''
    tabEx = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    '''Копируем выбранный диапозон из Excel'''
    tabEx.Copy()
    StartRow, StartCol = 1, 1
    sleep(1)
    sig.signal_Probar.emit(progressBar, 20)
    '''---------------------------------------------------------------'''
    """Создаем COM объект Word"""
    Word = win32com.client.Dispatch("Word.Application")
    # Word = win32com.client.gencache.EnsureDispatch("Word.Application")
    Word.Visible = 1
    Word.DisplayAlerts = False
    """Добавляем документ по шаблону"""
    sleep(1)

    strPath = ui.plainTextEdit_10.toPlainText()

    if strPath != '':
        strPath = strPath
    if strPath == '':
        # strPath = os.getcwd()
        strPath = wb.FullName.split(wb.Name)[0][:-1]

    """Добавляем документ по шаблону"""
    if ui.comboBox.currentIndex() == 0:
        fail = os.getcwd() + "\\Формат А4 (альбомный).dotx"
    if ui.comboBox.currentIndex() == 1:
        fail = os.getcwd() + "\\Формат А4 (книжный).dotx"
    if ui.comboBox.currentIndex() == 2:
        fail = os.getcwd() + "\\Формат А3 (альбомный).dotx"
    if ui.comboBox.currentIndex() == 3:
        fail = os.getcwd() + "\\Формат А3 (книжный).dotx"
    if ui.comboBox.currentIndex() == 4:
        # fail = os.getcwd() + "\\Шаблон_Спецификация оборудования, изделий и материалов.dotx"
        fail = os.getcwd() + "\\Формат А3 (альбомный).dotx"
        if count_col != 9:
            text = f"Спецификация оборудования, изделий и материалов должна содержать 9 столбцов"
            sig.signal_err.emit(Form, text)
            sig.signal_label.emit(ui.label, "Ошибка")
            return
    '''Сохранить Excel как'''
    # # 3129/5069/2-Р-001.004.390-НВК-01-С-001
    # nameobjextproekt = VXVtranslittext.GO(nameobjextproekt)
    # failPath = f"{strPath}\\{nameobjextproekt}.xlsx"
    # wb.SaveAs(failPath, CreateBackup=0)

    nameobjextproekt = ui.plainTextEdit_4.toPlainText()
    if ui.checkBox_3.isChecked() == True:
        if nameobjextproekt == '':
            text = f'Укажите ШИФР_ОБЪЕКТА'
            sig.signal_err.emit(Form, text)
            return
    saveAsNameobjextproekt = VXVtranslittext.GO(nameobjextproekt)
    FileName = f"{strPath}\\{saveAsNameobjextproekt}.docx"
    if '\n' in FileName:
        FileName = FileName.replace('\n', '')
    
    def OriForm():
        '''Определение ориентации и формата (размера) листа'''
        # OrientationLista, FormatLista = OriForm()
        OrientationLista = Doc.PageSetup.Orientation    # 1 - альбомная, 0 - книжная
        FormatLista = Doc.PageSetup.PaperSize           # 6 - формат А3, 7 - формат А4
        return OrientationLista, FormatLista

    '''Проверяем открыт ли одноименный файл *.docx и подключаемся к нему,
    иначе создаем новый файл'''
    prov = False
    for i in Word.Documents:
        if i.FullName == FileName:
            prov = True
            Doc = i
    if prov == False:
        Doc = Word.Documents.Add(fail)

    '''Проверяем соответствие ориентации листа согласно выбору'''
    sleep(1)
    OrientationLista, FormatLista = OriForm()
    if ui.comboBox.currentIndex() == 0:
        if OrientationLista != 1 or FormatLista != 7:
            Doc.Close(False)
            sleep(1)
            Doc = Word.Documents.Add(fail)

    if ui.comboBox.currentIndex() == 1:
        if OrientationLista == 1 or FormatLista != 7:
            Doc.Close(False)
            sleep(1)
            Doc = Word.Documents.Add(fail)

    if ui.comboBox.currentIndex() == 2:
        if OrientationLista != 1 or FormatLista == 7:
            Doc.Close(False)
            sleep(1)
            Doc = Word.Documents.Add(fail)

    if ui.comboBox.currentIndex() == 3:
        if OrientationLista == 1 or FormatLista == 7:
            Doc.Close(False)
            sleep(1)
            Doc = Word.Documents.Add(fail)

    if ui.comboBox.currentIndex() == 4:
        if OrientationLista != 1 or FormatLista == 7:
            Doc.Close(False)
            sleep(1)
            Doc = Word.Documents.Add(fail)
    
    Doc.Activate()

    if ui.checkBox_3.isChecked() == True:
        Doc.SaveAs(FileName)
    sleep(1)

    '''Выбираем все документе'''
    myRange = Doc.Range()
    sleep(1)
    '''Вставляем скопированную таблицу в World'''
    myRange.PasteExcelTable(False, False, False)
    '''Подключаемся к таблице'''
    tabWord = Doc.Tables(1)
    '''Автоподбор размера таблицы по содержимому'''
    tabWord.AutoFitBehavior(1)
    sleep(0.1)
    tabWord.AutoFitBehavior(2)

    '''Поля в ячейках таблицы'''
    tabWord.TopPadding = 0
    tabWord.BottomPadding = 0
    tabWord.LeftPadding = 0.05
    tabWord.RightPadding = 0
    tabWord.Spacing = 0
    tabWord.AllowPageBreaks = True
    tabWord.AllowAutoFit = True

    # '''Ширина столбцов'''
    # if ui.comboBox.currentIndex() == 4:
    #     WidthList = [2, 13, 6, 3.5, 4.5, 2, 2, 2.5, 4]
    #     tabWord.PreferredWidthType = 2
    #     for i in range(1, len(WidthList) + 1):
    #         tabWord.Columns(i).PreferredWidth = WidthList[i-1]


    '''Ширина столбцов'''
    if ui.comboBox.currentIndex() == 4:
        PT = 28.34646
        '''Установка единицы измерения размера таблицы''' 
        tabWord.PreferredWidthType = 3    # CM        
        WidthList = [2, 13, 6, 3.5, 4.5, 2, 2, 2.5, 4]
        '''Задаем общую ширину таблицы для более точного определения'''
        tabWord.PreferredWidth = sum(WidthList) * PT
        '''Проходим по всем колонкам для установки размеров из списка'''
        for i in range(1, len(WidthList) + 1):
            col = tabWord.Cell(1, i).Range.Columns
            col.PreferredWidthType = 3    # CM
            col.PreferredWidth = WidthList[i-1] * PT







    '''Высота строк в таблице'''
    PT = 28.34646  # количество "пт" в см
    Doc.Tables(1).Rows.HeightRule = 1  # указывает на способ изменения высоты: минимальный
    Doc.Tables(1).Rows.Height = 0.8 * PT  # RowHeight указывает на новую высоту строки в пунктах.

    try:
        if ui.comboBox.currentIndex() == 4:
            Doc.Tables(1).Rows(StartRow).Height = 3.2 * PT  # RowHeight указывает на новую высоту строки в пунктах.
    except:
        pass
    '''Обтекание таблиц'''
    # tabWord.Rows.WrapAroundText = True
    '''Выравниваем по вертикали все ячейки в таблице'''
    Doc.Tables(1).Range.Cells.VerticalAlignment = 1
    '''Удаляем интервал после абзаца во всей таблице'''
    Doc.Tables(1).Range.ParagraphFormat.SpaceBefore = 0  # интервал перед
    Doc.Tables(1).Range.ParagraphFormat.SpaceAfter = 0  # интервал после
    '''Добавляем отступ в ячейке слева'''
    Doc.Tables(1).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646

    '''Повторять как заголовок на каждой странице'''
    Doc.Tables(1).Cell(StartRow, 1).Range.Rows.HeadingFormat = True

    '''Выделяем колонку и делаем отступ в ячейках'''
    # tabWord.Columns(2).Select()
    # col = Doc.Application.Selection
    # col.ParagraphFormat.LeftIndent = 0.1 * 28.34646
    # Doc.Range(0, 0).Select()

    sig.signal_Probar.emit(progressBar, 40)
    # sleep(1)

    '''---------------------------------------------------------------'''
    # '''Заменяем имена полей пользовательских свойств на имена по умолчанию, которые понимает Сапсан'''
    # if ui.checkBox_3.isChecked() == False:
    #     NameCustPropSapsan = [
    #                     'Razrab',
    #                     'Фамилия проверяющего',
    #                     'Фамилия Гл. спец.',
    #                     'Фамилия согласующего',
    #                     'Фамилия нормоконтроллёра',
    #                     'GipFamily',
    #                     'Должность Гл. спец.',
    #                     'Должность согласующего',
    #                     'Дата нормоконтроля',
    #                     'Дата проверки',
    #                     'Дата согласования',
    #                     'Шифр_документа',
    #                     'Название проекта',
    #                     'ObjectName',
    #                     'StageKey',
    #                     'Название спецификации',
    #                     'Шифр_ревизии']

    #     for name in Doc.CustomDocumentProperties:
    #         for default_Name in NameCustPropSapsan:
    #             if default_Name in name:
    #                 Doc.CustomDocumentProperties(name).Name = default_Name


    if ui.checkBox_3.isChecked() == True:
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
        
        '''Заполняем штамп фамилиями'''
        if ui.checkBox_2.isChecked() == True:
            '''Собираем фамилии из таблицы'''
            doljList = []
            UserList = []
            for i in range(ui.tableWidget_1.rowCount()):
                dolj = eval(f'ui.tableWidget_1.item({i}, 0).text()')
                xxx = eval(f'ui.tableWidget_1.item({i}, 1).text()')
                doljList.append(dolj)
                if dolj == '':
                    UserList.append('')
                else:
                    UserList.append(xxx)

            # rowList = 6, 7, 8, 9, 10, 11
            rowList = 7, 8, 9, 10, 11, 12
            sec = time.localtime(time.time())
            now = f'{str(sec.tm_mday).rjust(2, "0")}.{str(sec.tm_mon).rjust(2, "0")}.{str(sec.tm_year)[-2:]}'

            sig.signal_Probar.emit(progressBar, 45)
            
            '''Отправляем данные в штамп на 1-ом листе'''
            cdpdict = {
                    'Razrab' : UserList[0],
                    'Фамилия проверяющего' : UserList[1],
                    'Фамилия Гл. спец.' : UserList[2],
                    'Фамилия согласующего' : UserList[3],
                    'Фамилия нормоконтроллёра' : UserList[4],
                    'GipFamily' : UserList[5],

                    'Должность Гл. спец.' : doljList[2],
                    'Должность согласующего' : doljList[3],
                    
                    'Дата нормоконтроля' : now,
                    'Дата проверки' : now,
                    'Дата согласования' : now,
                    
                    'Шифр_документа' : ui.plainTextEdit_4.toPlainText(),
                    'Название проекта' : ui.plainTextEdit_5.toPlainText(),
                    'ObjectName' : ui.plainTextEdit_6.toPlainText(),
                    'StageKey' : ui.plainTextEdit_8.toPlainText(),
                    'Название спецификации' : ui.plainTextEdit_9.toPlainText(),
                    'Шифр_ревизии' : 'C01'
                    }

            '''Заменяем значение полей пользовательских свойств'''
            for key, value in cdpdict.items():
                Doc.CustomDocumentProperties(key).Value = value

            '''Обновляем поля свойств в основной области с текстом'''
            Doc.Fields.Update()
            '''Обновляем поля свойств в нижнем колонтитуле'''
            Footers = Doc.Sections(1).Footers
            for i in range(1, Footers.Count + 1):
                Footers(i).Range.Fields.Update()

 
        '''---------------------------------------------------------------'''
        '''Работа с подписями'''
        directory = str(ui.plainTextEdit_3.toPlainText())
        if ui.checkBox.isChecked() == True:
            if directory == '':
                sig.signal_err.emit(Form, "Подписи не были вставлены в штамп.\nУкажите папку с подписями в формате *.jpg , *.png")
                if Doc.Saved == False: Doc.Save()
                return
            try:
                direct = os.listdir(directory)
            except FileNotFoundError:
                sig.signal_err.emit(Form, "Папка с подписями не найдена")
                return

            '''Собираем список с полным именем файлов в папке с подписями'''
            PatchFileList = []
            for filename in direct:
                FullName = os.path.join(directory, filename)
                if os.path.isfile(FullName):
                    if ".png" in FullName:
                        PatchFileList.append(FullName)
                        continue
                    if ".jpg" in FullName:
                        PatchFileList.append(FullName)

            patchPod = []
            userErr = []
            '''Для каждого значения фамилии из таблицы'''
            for User in UserList:
                UserTrue = False
                '''если оно не равно '' '''
                if User != '':
                    '''перебираем все полные пути файлов в папке'''
                    for patchP in PatchFileList:
                        '''если фамилия есть в адресе файла'''
                        if User in patchP:
                            '''обозначаем наличие фамилии в названиях файлов в папке'''
                            UserTrue = True
                            '''добавляем адрес файла для фамилии из таблицы'''
                            patchPod.append(patchP)
                            '''Производит переход за пределы объемлющего цикла (всей инструкции цикла 
                            на уровень "for User in UserList") при нахождении фамилии'''
                            break

                    '''Если наличие файла в папке не подтвердилось'''
                    if UserTrue == False:
                        '''Cписок не найденных фамилий'''
                        userErr.append(f'"{User}"')
                        patchPod.append('')
                if User == '':
                    patchPod.append('')

            '''Работа над ошибками'''
            if userErr != []:
                UserErrList = ', '.join(userErr)
                text = f"Не найдены картинки с фамилией {UserErrList} в папке: \n{directory}"
                sig.signal_err.emit(Form, text)
            # printTabconsole([UserList, rowList, patchPod], add_column = True)

            '''Вставляем картинки с подписями в штамп и делаем их перед текстом'''
            xxx = 50
            for i in range(len(patchPod)):
                xxx += 8
                sig.signal_Probar.emit(progressBar, xxx)
                if patchPod[i] != '':
                    if '..png' in patchPod[i]:
                        FileName = patchPod[i]
                    else:
                        FileName = imageZeroFon.GO(patchPod[i])
                    img = FootersTables_2.Cell(rowList[i], 4).Range.InlineShapes.AddPicture(FileName=FileName, LinkToFile=False, SaveWithDocument=True)
                    img.ConvertToShape().WrapFormat.Type = 3
                    sleep(0.5)

        if Doc.Saved == False:
            try:
                Doc.Save()
            except:
                pass
    sig.signal_label.emit(ui.label, "Выполнено: таблица вставлена в документ Word . . .")

    '''---------------------------------------------------------------'''
if __name__ == "__main__":
    # import sys
    # from ExcelToWorld import app, ui, Form, sig

    # GO(ui, Form, sig)
    # sys.exit(app.exec_())
    
    os.system(r'call C:\vxvproj\tnnc-ExcelToWorld\tnnc-ExcelToWorld\ExcelToWorld.py')
