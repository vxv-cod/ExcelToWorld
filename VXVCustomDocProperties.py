import os
from time import sleep
import win32com.client
import win32com.client.gencache

from VBAExcel import EndIndexRowCol

Word = win32com.client.Dispatch("Word.Application")
Doc = Word.ActiveDocument
    
cdpListName = [
        'Razrab',
        'Фамилия проверяющего',
        'Фамилия Гл. спец.',
        'Фамилия согласующего',
        'Фамилия нормоконтроллёра',
        'GipFamily',
        'Шифр_документа',
        'Название проекта',
        'Название спецификации',
        'StageKey',
        'ObjectName',
        'Шифр_ревизии',
        'Дата нормоконтроля',
        'Дата проверки',
        'Дата согласования',
        'Должность согласующего',
        'Должность Гл. спец.'
        ]
    
cdpListValue = [
        'Разраб.',
        'Пров.',
        'Гл. спец.',
        'Соглас.',
        'Н. контр.',
        'ГИП',
        'Шифр_документа',
        'Название проекта',
        'Спецификация оборудования, изделий и материалов',
        'ПД',
        'ObjectName',
        'CO1',
        '88.88.88',
        '88.88.88',
        '88.88.88',
        'Нач. отд.',
        'Гл. спец.'
        ]


def DelCDP():
    '''Удаляем все поля свойст документа'''
    for i in Doc.CustomDocumentProperties:
        i.Delete()


def getCDP():
    cdpName = [i.Name for i in Doc.CustomDocumentProperties]
    cdpValue = [i.Value for i in Doc.CustomDocumentProperties]
    return cdpName, cdpValue

def setCDP(cdpListName, cdpListValue):
    for i in range(len(cdpListName)):
        try:
            Doc.CustomDocumentProperties.Add(cdpListName[i], False, 4, cdpListValue[i])
        except:
            pass

DelCDP()
setCDP(cdpListName, cdpListValue)


# def NewCDPValue(cdpListName, cdpListValue):
#     '''Заменяем значения в полях по их именам'''
#     for i in range(len(cdpListName)):
#         Doc.CustomDocumentProperties(cdpListName[i]).Value = cdpListValue[i]
    
#     '''Коллекция всех нижних колонтитулов'''
#     Footers = Doc.Sections(1).Footers
#     '''Подключаемся к 1-ой таблице нижнего колонтитула на 1-ом листе'''
#     Footers = Doc.Sections(1).Footers
#     Footers(1).Range.Tables(1).Select()
    
#     Selection = Doc.Application.Selection
#     Selection.Fields.Update()




def UpdateCDP():
    '''Обновляем поля свойств в основной области с текстом'''
    Doc.Fields.Update()
    '''Обновляем поля свойств в нижнем колонтитуле'''
    Footers = Doc.Sections(1).Footers
    for i in range(1, Footers.Count + 1):
        Footers(i).Range.Fields.Update()

        
# UpdateCDP()


def perenosZagolovka():
    # tabWord = Doc.Tables(1)
    # # tabWord.Rows(2).Select()
    # # Selection = Doc.Application.Selection
    # # Selection.Tables(1).Rows.HeadingFormat = True
    # tabWord.Rows(2).HeadingFormat = True
    
    # Word = win32com.client.Dispatch("Word.Application")
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.ActiveWorkbook
    sheet = wb.ActiveSheet
    EndRow, EndCol = EndIndexRowCol(sheet)
    tabEx = sheet.Range(sheet.Cells(1, 1), sheet.Cells(EndRow, EndCol))
    sleep(1)
    tabEx.Copy()
    sleep(1)    

    


    Word.Visible = 1
    fail = os.getcwd() + "\\Шаблон_печати_А4_альбом.dotx"
    Doc = Word.Documents.Add(fail)

    myRange = Doc.Range()
    myRange.Delete()
    sleep(0.5)
    myRange.InsertParagraphAfter()
    myRange = Doc.Paragraphs(2).Range
    myRange.PasteExcelTable(False, False, False)
    # myRange.Paste()
    sleep(2)

    # Word.DisplayAlerts = False
    StartRow, StartCol = 1, 1


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

    '''Высота строк в таблице'''
    PT = 28.34646  # количество "пт" в см
    Doc.Tables(1).Rows.HeightRule = 1  # указывает на способ изменения высоты: минимальный
    Doc.Tables(1).Rows.Height = 0.8 * PT  # RowHeight указывает на новую высоту строки в пунктах.
    try:
        Doc.Tables(1).Rows(StartRow).Height = 1.0 * PT  # RowHeight указывает на новую высоту строки в пунктах.
        Doc.Tables(1).Rows(StartRow + 1).Height = 1.2 * PT  # RowHeight указывает на новую высоту строки в пунктах.
    except:
        pass
    # '''Обтекание таблиц'''
    # tabWord.Rows.WrapAroundText = True

    # Doc.Tables(1).Rows.HeadingFormat = True
    '''Повторять как заголовок на каждой странице'''
    tabWord.Rows(StartRow).HeadingFormat = True

    '''Выравниваем по вертикали все ячейки в таблице'''
    Doc.Tables(1).Range.Cells.VerticalAlignment = 1
    '''Удаляем интервал после абзаца во всей таблице'''
    Doc.Tables(1).Range.ParagraphFormat.SpaceBefore = 0  # интервал перед
    Doc.Tables(1).Range.ParagraphFormat.SpaceAfter = 0  # интервал после
    # Doc.Tables(1).Range.ParagraphFormat.LeftIndent = 0.1 * 28.34646

    '''---------------------------------------------------------------'''






    print("---END---")




# perenosZagolovka()




print("--------")









def NewCDPValue111():
    '''Заменяем значения в полях по их именам'''
    for i in Doc.CustomDocumentProperties:
        i.Value = ''

# NewCDPValue111()



def NewCDPName(cdpListNameOld, cdpListNameNew):
    '''Заменяет имена полей в староом списке из нового '''
    for i in range(len(cdpListNameOld)):
        Doc.CustomDocumentProperties(cdpListNameOld[i]).Name = cdpListNameNew[i]
    Doc.Fields.Update()

# NewCDPValue(cdpListName, cdpListValue)


# setCDP(cdpListName, cdpListValue)



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

# ListCDPPrint()




def newCDP():
    '''Вставляем новое пользовательское свойство документа 
    (нельзя вставить свойство с таким же именем, будет ошибка)'''
    name = "vxv_7"
    value = "Value"
    try:
        Doc.CustomDocumentProperties.Add(name, False, 4, value)
    except:
        pass
    UserProp = Doc.CustomDocumentProperties(name)
    NameUserProp = UserProp.Value = "6666666"
    # VameUserProp = UserProp.Name = "EEEEEEEEEEE"
    NameUserProp = UserProp.Name
    NameUserProp = UserProp.Value
    Doc.Fields.Update()

# cdpName, cdpValue = getCDP()
# print("---")
# print(f'cdpName = {cdpName}')
# print("---")
# print(f'cdpValue = {cdpValue}')
# print("---")



sapsanCDP = [
        'CompanyName',
        'COMPANYNAME1-B8B9-BD780F24413A',
        'DisciplineId',
        'DISCIPLINEID1-B8B9-BD780F24413A',
        'DocPackageId',
        'DOCPACKAGEID1-B8B9-BD780F24413A',
        'DOCUMENTPARTNAME',
        'NameDraft',
        'NAMEDRAFT1-B8B9-BD780F24413A',
        'ObjectName',
        'OBJECTNAME1-B8B9-BD780F24413A',
        'Page',
        'PAGE1-B8B9-BD780F24413A',
        'PartName',
        'PARTNAME1-B8B9-BD780F24413A',
        'PrintJobId',
        'Stadiya',
        'STADIYA1-B8B9-BD780F24413A',
        'TemplateId',
        'TEMPLATEID1-B8B9-BD780F24413A',
        'ГИП',
        'ГИП1-B8B9-BD780F24413A',
        'ГИПтип',
        'Дата_изма',
        'ДАТА_ИЗМА1-B8B9-BD780F24413A',
        'Кол_уч',
        'КОЛ_УЧ1-B8B9-BD780F24413A',
        'Листов',
        'ЛИСТОВ1-B8B9-BD780F24413A',
        'Масса',
        'Масштаб',
        'Материал',
        'Название проекта',
        'НАЗВАНИЕ ПРОЕКТА1-B8B9-BD780F24413A',
        'Название_листа_разреш',
        'НАЗВАНИЕ_ЛИСТА_РАЗРЕШ1-B8B9-BD780F24413A',
        'Номер_изма',
        'НОМЕР_ИЗМА1-B8B9-BD780F24413A',
        'Номер_листа_разреш',
        'НОМЕР_ЛИСТА_РАЗРЕШ1-B8B9-BD780F24413A',
        'Подписант_1',
        'ПОДПИСАНТ_11-B8B9-BD780F24413A',
        'Подписант_2',
        'ПОДПИСАНТ_21-B8B9-BD780F24413A',
        'Подписант_3',
        'ПОДПИСАНТ_31-B8B9-BD780F24413A',
        'Подписант_4',
        'ПОДПИСАНТ_41-B8B9-BD780F24413A',
        'Подписант_5',
        'ПОДПИСАНТ_51-B8B9-BD780F24413A',
        'Подписант_6',
        'ПОДПИСАНТ_61-B8B9-BD780F24413A',
        'Подписант1_Дата',
        'ПОДПИСАНТ1_ДАТА1-B8B9-BD780F24413A',
        'Подписант2_Дата',
        'ПОДПИСАНТ2_ДАТА1-B8B9-BD780F24413A',
        'Подписант3_Дата',
        'ПОДПИСАНТ3_ДАТА1-B8B9-BD780F24413A',
        'Подписант4_Дата',
        'ПОДПИСАНТ4_ДАТА1-B8B9-BD780F24413A',
        'Подписант5_Дата',
        'ПОДПИСАНТ5_ДАТА1-B8B9-BD780F24413A',
        'Подписант6_Дата',
        'ПОДПИСАНТ6_ДАТА1-B8B9-BD780F24413A',
        'ПодписантТип_1',
        'ПОДПИСАНТТИП_11-B8B9-BD780F24413A',
        'ПодписантТип_2',
        'ПОДПИСАНТТИП_21-B8B9-BD780F24413A',
        'ПодписантТип_3',
        'ПОДПИСАНТТИП_31-B8B9-BD780F24413A',
        'ПодписантТип_4',
        'ПОДПИСАНТТИП_41-B8B9-BD780F24413A',
        'ПодписантТип_5',
        'ПОДПИСАНТТИП_51-B8B9-BD780F24413A',
        'ПодписантТип_6',
        'ПОДПИСАНТТИП_61-B8B9-BD780F24413A',
        'Шифр_документа',
        'ШИФР_ДОКУМЕНТА1-B8B9-BD780F24413A',
        'Шифр_проекта',
        'ШИФР_ПРОЕКТА1-B8B9-BD780F24413	'
        ]


def create(cdpName:list):
    value = ''
    for name in cdpName:
        try:
            Doc.CustomDocumentProperties.Add(name, False, 4, value)
        except:
            pass
        # Doc.CustomDocumentProperties(name).Value = cdpValue[cdpName.index(name)]

# create(cdpListName)
# DelCDP()
# create(sapsanCDP)
# ListCDPPrint()

print("--------")


