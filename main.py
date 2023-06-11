import openpyxl
from openpyxl.utils import get_column_letter
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

FILENAME = ''
TABLE = {"number": [],
         "code": [],
         'name': [],
         'count': [],
         'price': [],
         'price_without_nds': [],
         'price_with_nds': [],
         'nds': [],
         'GTD': [],
         'numberUPD': '',
         'dateUPD': ''}
def openfile(): # функция для открытия даиалогового окна выбора файла
    global FILENAME
    FILENAME = filedialog.askopenfilename()
    short_filename = FILENAME[FILENAME.rfind('/')+1:]
    label_filename = Label(window, text=f'Файл: {short_filename}') # создается текст, где прописан путь к
    # выбранному
    # файлу
    label_filename.grid(column=1, row=0)#label_filename.grid(column=1, row=0, padx=(50,0),pady=(50,0))


def read_book():
    workbook = openpyxl.load_workbook(FILENAME, data_only=True)  # загрузили книгу
    sheets_list = workbook.sheetnames  # выгрузили имена листов
    sheet_active = workbook[sheets_list[0]]  # выбрали первый лист
    row_max = sheet_active.max_row
    column_max = sheet_active.max_column
    keys = list(TABLE.keys())  # вывод всех ключей в список
    count = 0
    for col in [3, 4, 5, 6, 7, 8, 10, 11, 13]:
        key = keys[count]  # выбор ключа
        print(key)
        for row in range(2, row_max + 1):
            cell_letter = str(get_column_letter(col)) + str(row)  # ячейка
            val = sheet_active[cell_letter].internal_value  # значение ячейки
            TABLE[key].append(val)  # добавили значение ячейки
        print(TABLE[key])
        count += 1
    TABLE['numberUPD'] = str(sheet_active['A2'].internal_value)
    s = str(sheet_active['B2'].internal_value)
    TABLE['dateUPD'] = s[8:10] + '.' + s[5:7] + '.' + s[0:4]



def table_in_line():
    lengt = TABLE['number'][-1]
    line ='<ТаблСчФакт>'
    for i in range(lengt):
        line += f'<СведТов НомСтр="{TABLE["number"][i]}" ' \
                f'НаимТов="{TABLE["name"][i]}"' \
                f' ОКЕИ_Тов="796" КолТов="{TABLE["count"][i]}" ЦенаТов="{TABLE["price"][i]}" СтТовБезНДС="' \
                f'{TABLE["price_without_nds"][i]}" ' \
                f'НалСт="20%" ' \
                f'СтТовУчНал="{TABLE["price_with_nds"][i]}"><Акциз><БезАкциз>без ' \
                f'акциза</БезАкциз></Акциз><СумНал><СумНал>{TABLE["nds"][i]}</СумНал></СумНал>' \
                f'<СвТД КодПроисх="156" НомерТД="{TABLE["GTD"][i]}" />' \
                f'<ДопСведТов ПрТовРаб="1" КодТов="{TABLE["code"][i]}" НаимЕдИзм="шт" КрНаимСтрПр="КИТАЙ" НадлОтп="0" /></СведТов>'
    line += '</ТаблСчФакт>'
    return line

def change_line(text_editor1,text_editor2):
    input_line = text_editor1.get(1.0, END)

    num1_start = input_line.find('ВерсФорм')-6
    num1_end = input_line.find('ВерсФорм')-2
    num1 = str(int(input_line[num1_start:num1_end]) + 1)
    output_line = input_line[:num1_start] + num1

    upd1_start = input_line.find('НомерСчФ=')
    codeOKV = input_line[input_line.find('КодОКВ') + 8:input_line.find(' ДатаСчФ') - 1]
    output_line += input_line[num1_end:upd1_start] + 'НомерСчФ="' + TABLE['numberUPD'] + '" КодОКВ="' + codeOKV + \
                   '" ДатаСчФ="' + TABLE['dateUPD'] + '">'

    upd2_start = input_line.find('НомДокОтгр') + 12
    output_line += input_line[input_line.find('<ИспрСчФ ДефНомИспрСчФ'):upd2_start] + TABLE['numberUPD'] + \
        '" ДатаДокОтгр="' + TABLE['dateUPD'] + '" /><ИнфПолФХЖ1></ИнфПолФХЖ1></СвСчФакт>'
    output_line += table_in_line()
    #print(num1_start,num1_end,upd1_start,codeOKV,upd2_start,input_line.find('<ИспрСчФ ДефНомИспрСчФ'))
    output_line += '<СвПродПер><СвПер СодОпер="Товары переданы"><ОснПер ДатаОсн="' + TABLE['dateUPD'] + \
        '" НаимОсн="Уведомление о выкупе" НомОсн="' + TABLE['numberUPD'] + '" />'
    output_line += input_line[input_line.find('<СвПерВещи'):]
    text_editor2.insert(1.0, output_line)

def start(text_editor1,text_editor2):
    global TABLE
    label_progress = Label(window, text='Пошел процесс')
    label_progress.grid(column=1, row=1)#label_progress.grid(column=0, row=2, padx=(50, 0), pady=(50, 0))
    read_book()
    change_line(text_editor1, text_editor2)


window = Tk() # создается окно интрефейса
window.title("Данные из екселя в текст")
window.geometry("800x1300")

for c in range(10): window.columnconfigure(index=c, weight=10)
for r in range(10): window.rowconfigure(index=r, weight=10)


button_chose = Button(window,text='Выбрать файл', command=openfile) #создается кнопка с функциее откртия файла
button_chose.grid(column=0, row=0)  #button_chose.grid(column=0, row=0, padx=(50,0), pady=(50,0))


text_editor1 = Text(width=40, height=10, wrap=WORD)
text_editor1.grid(column=0, row=3, columnspan=2)
text_editor2 = Text(width=40, height=10, wrap=WORD)
text_editor2.grid(column=0, row=4, columnspan=2)


button_start = Button(window, text='Выполнить', command=lambda: start(text_editor1,text_editor2))
button_start.grid(column=0, row=1)  #button_start.grid(column=0, row=1, padx=(10,0), pady=(50,0))

window.mainloop()
