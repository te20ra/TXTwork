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
         'GTD': []}
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
    print(keys)
    print(column_max, row_max)
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

def stroka():
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
                f'<ДопСведТов ПрТовРаб="1" КодТов="{TABLE["code"][i]}" НаимЕдИзм="шт" КрНаимСтрПр="КИТАЙ" НадлОтп="0" /></СведТов> '
    line += '</ТаблСчФакт>'
    print(line)
    return line
def start(text_editor1,text_editor2):
    global TABLE
    label_progress = Label(window, text='Пошел процесс')
    label_progress.grid(column=1, row=1)#label_progress.grid(column=0, row=2, padx=(50, 0), pady=(50, 0))
    read_book()
    line = stroka()
    text_editor2.insert(1.0, line)


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
path_to_file = 'Копия Модуль 4. Урок 17.xlsx'
