import openpyxl
from openpyxl.utils import get_column_letter
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

FILENAME = ''
TABLE = {"number": [],
         'name': [],
         'count': [],
         'price': [],
         'price_without_nds': [],
         'price_with_nds': [],
         'nds': [],
         'GTD':[],
             }
def openfile(): # функция для открытия даиалогового окна выбора файла
    global FILENAME
    FILENAME =filedialog.askopenfilename()
    label_filename = Label(window, text=f'Выбранный файл: {FILENAME}') # создается текст, где прописан путь к выбранному файлу
    label_filename.grid(column=1, row=0, padx=(50,0),pady=(50,0))




def start():
    global TABLE
    label_progress = Label(window, text='Пошел процесс')
    label_progress.grid(column=0, row=2, padx=(50, 0), pady=(50, 0))
    workbook = openpyxl.load_workbook(FILENAME, data_only=True) # загрузили книгу
    sheets_list = workbook.sheetnames # выгрузили имена листов
    sheet_active = workbook[sheets_list[0]] # выбрали первый лист
    row_max = sheet_active.max_row
    column_max = sheet_active.max_column
    keys = list(TABLE.keys()) # вывод всех ключей в список
    print(keys)
    print(column_max, row_max)
    count = 0
    for col in [1,2,4,5,6,8,9,11]:
        key = keys[count] # выбор ключа
        print(key)
        for row in range(2, row_max+1):
            cell_letter = str(get_column_letter(col)) + str(row) # ячейка
            #print(cell_letter)
            val = sheet_active[cell_letter].internal_value # значение ячейки
            #print(val)
            TABLE[key].append(val) # добавили значение ячейки
        print(TABLE[key])
        count += 1

window = Tk() # создается окно интрефейса
window.title("Данные из екселя в текст")
window.geometry("1920x1080")



#label_1 = Label(window,text="Выберите Excel файл:")
#label_1.grid(column=0,row=0, padx=(50,0), pady=(50,0))

button_chose = Button(window,text='Выбрать файл', command=openfile) # создается кнопка с функциее откртия файла
button_chose.grid(column=0, row=0, padx=(50,0), pady=(50,0))

button_start = Button(window, text='Выполнить',command=start)
button_start.grid(column=0, row=1, padx=(10,0), pady=(50,0))
#workbook = openpyxl.load_workbook(FILENAME,data_only=True)
#sheets_list = workbook.sheetnames
#sheet_active = workbook[sheets_list[0]]



window.mainloop()
path_to_file = 'Копия Модуль 4. Урок 17.xlsx'
