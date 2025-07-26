# -*- coding: utf–8 -*-
import os
import re
import random
import time
import copy
import re

# import aspose.pdf as pdf
from tkinter import *
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter import messagebox
from tkinter import colorchooser
from tkinter import font

from PIL import ImageGrab, ImageEnhance
from PIL import ImageTk

class Main(Tk):

    def __init__(self):
        super().__init__()

        self.cellfont = "Century"
        self.font16 = 'Arial 16 bold'
        self.font13 = 'Arial 13 bold'
        self.font10 = 'Arial 10 bold'
        self.font9 = 'Arial 9 bold'
        # main color scheme
        self.majorcolor = "#047a00"
        self.minorcolor = "#035c00"
        self.rightframecolor = "white"
        self.font_color1 = "black"
        self.font_color2 = "black"
        # button colors
        self.buttonfgcolor = "black"
        self.buttoncolor = "#cfcfcf"
        self.activebuttoncolor = "#ffffff"
        # label colors
        self.labelfgcolor = "white"
        self.approvecolor = "#14cc00"
        self.deniedcolor = "red"
        self.waitcolor = "#ffe600"
        #entry colors
        self.disabledentrycolor = "#242424"
        self.cellcolor = "white"
        self.fixedcellcolor = "#57ff47"

        self.eng_alphabet = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 
                             'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
        self.ENG_alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
                             'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        self.rus_alphabet = ['а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж', 'з', 'и', 'й', 
                             'к', 'л', 'м', 'н', 'о', 'п', 'р', 'с', 'т', 'у', 'ф', 
                             'х', 'ц', 'ч', 'ш', 'щ', 'ь', 'ы', 'ъ', 'э', 'ю', 'я']
        self.RUS_alphabet = ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ё', 'Ж', 'З', 'И', 'Й', 
                             'К', 'Л', 'М', 'Н', 'О', 'П', 'Р', 'С', 'Т', 'У', 'Ф', 
                             'Х', 'Ц', 'Ч', 'Ш', 'Щ', 'Ь', 'Ы', 'Ъ', 'Э', 'Ю', 'Я', '_']
        self.tatar_extension = {'1': 'Ә',  '2': 'Җ',  '3': 'Ң', '4': 'Ө',  '5': 'Ү',  '6': 'Һ'}
        self.digits = [str(i) for i in range(10)]
        self.global_alphabet = (self.eng_alphabet + self.ENG_alphabet + self.rus_alphabet + 
                                    self.RUS_alphabet + self.digits)

        self.char_check = (self.register(self.char_valid), "%P")

        # Стили виджетов ttk
        buttonstyle = ttk.Style()
        buttonstyle.theme_use('alt')
        buttonstyle.configure("TButton",  color=self.buttoncolor, relief='ridge',
                              foreground=self.buttonfgcolor, focusthickness=0, focuscolor='none', 
                                font=self.font13, justify="center", highlightcolor='cyan')
        buttonstyle.map('TButton', background=[('active', self.activebuttoncolor)])

        infolabel = ttk.Style()
        infolabel.configure("infolabel.TLabel", foreground=self.labelfgcolor, 
                                    background=self.minorcolor, font=self.font10, padding=[0, 0])
        notificationlabel = ttk.Style()
        notificationlabel.configure("notificationlabel.TLabel", 
                                    background=self.minorcolor, font=self.font13,
                                    padding=[0, 0], anchor=TOP)

        # Левая и правая части интерфейса
        # self.mainleftframe = Frame(self, bg=self.majorcolor)
        # self.mainleftframe.grid(row=0, column=0, sticky="nsew")
        # self.canvas = Canvas(self.mainleftframe, borderwidth=0, background="#ffffff")
        # self.leftframe = Frame(self.canvas, bg=self.majorcolor)
        # self.leftframe.pack(side="left", fill="y")
        # self.vsb = Scrollbar(self.mainleftframe, orient="vertical", command=self.canvas.yview)
        # self.canvas.configure(yscrollcommand=self.vsb.set)
        # self.vsb.pack(side="right", fill="y")
        # self.canvas.pack(side="left", fill="both", expand=True)
        # self.canvas.create_window((4, 4), window=self.leftframe, anchor="nw")
        
        # self.canvas.update_idletasks() 
        # self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.leftframe = Frame(self, bg=self.majorcolor)
        self.leftframe.grid(row=0, column=0, sticky="nsew")

        self.rightframe = Frame(self, bg=self.rightframecolor)
        self.rightframe.grid(row=0, column=1, sticky="nsew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Наполнение левой части
        self.opendirectorybutton = ttk.Button(self.leftframe, text='Выбрать папку', 
                                              width=22, command=self.open_directory)
        self.opendirectorybutton.grid(row=0, column=0, columnspan=4, pady=5, padx=10, sticky="nsew")
        
        self.dictlistbox = Listbox(self.leftframe, border=3, selectmode=MULTIPLE, height=3, font=self.font10)
        self.dictlistbox.grid(row=1, column=0, columnspan=4, pady=0, padx=10, ipadx=5, ipady=3, sticky="nsew")

        self.infolabel = ttk.Label(self.leftframe, style="infolabel.TLabel", text=2*'\n')
        self.infolabel.grid(row=2, column=0, columnspan=4, pady=0, padx=10, sticky="nsew")

        self.selectdictbutton = ttk.Button(self.leftframe, text='Загрузить словарь',
                                    command=lambda: self.notifiationlabel.config(
                                        text='Директория со\nсловарями не выбрана!', 
                                    style="notificationlabel.TLabel", foreground='red'))
        self.selectdictbutton.grid(row=3, column=0, columnspan=4, pady=3, padx=10, sticky="nsew")

        self.notifiationlabel = ttk.Label(self.leftframe, style="notificationlabel.TLabel", text=1*'\n')
        self.notifiationlabel.grid(row=4, column=0, columnspan=4, pady=0, padx=10, sticky="nsew")
        
        self.hintlabel = ttk.Label(self.leftframe, font=self.font9, foreground="yellow", 
                                background=self.majorcolor, text="1: Ә     2: Җ     3: Ң     4: Ө     5: Ү     6: Һ\n")
        self.hintlabel.grid(row=5, column=0, columnspan=4, pady=0, padx=10, sticky="nsew")

        self.sizelabel = ttk.Label(self.leftframe, style="infolabel.TLabel", 
                                   background=self.majorcolor, font=self.font13, text='Конфигурация сетки', anchor="center")
        self.sizelabel.grid(row=6, column=0, columnspan=4, pady=0, padx=10)

        self.check = False
        # self.checkbutton = Checkbutton(self.leftframe, bg=self.majorcolor, text='⮁', fg="white",
        #                 font=self.font13, activebackground=self.majorcolor, command=self.check_change)
        # self.checkbutton.grid(row=6, column=3, pady=0, padx=0, sticky='nsew')

        # 1 поле конфигурации
        self.word_length_label = ttk.Label(self.leftframe, background=self.majorcolor,  justify=LEFT,
                                   foreground=self.labelfgcolor, font=self.font10, text='Длина слова:')
        self.word_length_label.grid(row=7, column=0, padx=10, sticky="w")
        self.entry1 = Entry(self.leftframe, justify=CENTER, width=15)
        self.entry1.grid(row=7, column=1, padx=10, pady=3, columnspan=3, sticky="w")
        self.entry1.configure(state='normal', fg="#b8b8b8")
        self.entrybind1_in = self.entry1.bind('<Button-1>', lambda x: self.on_focus_in(self.entry1))

        # 2 поле конфигурации
        self.phrase = ttk.Label(self.leftframe, background=self.majorcolor,  justify=LEFT,
                                   foreground=self.labelfgcolor, font=self.font10, text='Фраза:')
        self.phrase.grid(row=8, column=0, padx=10, pady=0, sticky="w")
        self.phrasetext = Text(self.leftframe, width=28, height=2, wrap="word", validatecommand=self.char_check)
        self.phrasetext.grid(row=9, column=0, columnspan=4, padx=12, pady=3, sticky="nsew")

        self.phrase_position = 2
        # self.phrase_position_label = ttk.Label(self.leftframe, background=self.majorcolor,  justify=LEFT,
        #                            foreground=self.labelfgcolor, font=self.font10, text='Позиция фразы:')
        # self.phrase_position_label.grid(row=10, column=0, padx=10, sticky="w")
        # self.phrase_position_entry = Entry(self.leftframe, justify=CENTER, width=15)
        # self.phrase_position_entry.grid(row=10, column=1, padx=10, pady=5, columnspan=3, sticky="w")

        # 3 поле конфигурации
        self.bd_size_label = ttk.Label(self.leftframe, background=self.majorcolor,  justify=LEFT,
                                   foreground=self.labelfgcolor, font=self.font10, text='Толщина границ:')
        self.bd_size_label.grid(row=11, column=0, columnspan=2, padx=10, sticky="w")
        self.bd_scale_var = IntVar()
        self.bd_scale = Scale(self.leftframe, orient=HORIZONTAL, length=87,  from_=1, to=10, variable=self.bd_scale_var)
        self.bd_scale.set(1)
        self.bd_scale.grid(row=11, column=1, columnspan=2, padx=10, pady=0, sticky="nsew")
        self.bd_scale.configure(state='normal', bg=self.majorcolor, fg='white', troughcolor='white', highlightbackground=self.majorcolor)

        # 4 поле конфигурации
        self.font_label = ttk.Label(self.leftframe, background=self.majorcolor, justify=LEFT,
                                   foreground=self.labelfgcolor, font=self.font10, text='Шрифт:')
        self.font_label.grid(row=12, column=0, padx=10, sticky="w")
        # self.font_var = StringVar(value=font.families()[11])  
        self.font_combobox = ttk.Combobox(self.leftframe, values=font.families(), width=13, state="readonly")
        self.font_combobox.set(font.families()[11])
        self.font_combobox.grid(row=12, column=1, columnspan=2, padx=10, pady=3, sticky="w")

        # 5 поле конфигурации
        self.def_cell_color_label = ttk.Label(self.leftframe, background=self.majorcolor,  justify=LEFT,
                                   foreground=self.labelfgcolor, font=self.font10, text='Обычная клетка:')
        self.def_cell_color_label.grid(row=13, column=0, padx=10, sticky="w")
        self.select_color1_btn = Button(self.leftframe, relief="groove", justify=CENTER, width=4, command = lambda: self.choose_color(1))
        self.select_color1_btn.grid(row=13, column=1, columnspan=1, padx=10, pady=3, sticky="nsew")
        self.select_color1_btn.configure(state='normal', fg="#b8b8b8", bg=self.cellcolor)

        self.select_fontcolor1_btn = Button(self.leftframe, relief="groove", justify=CENTER, width=4, command = lambda: self.choose_color(3))
        self.select_fontcolor1_btn.grid(row=13, column=2, columnspan=1, padx=10, pady=3, sticky="nsew")
        self.select_fontcolor1_btn.configure(state='normal', fg="#b8b8b8", bg=self.font_color1)

        # 6 поле конфигурации
        self.fixed_color_label = ttk.Label(self.leftframe, background=self.majorcolor, justify=LEFT, 
                                   foreground=self.labelfgcolor, font=self.font10, text='Фиксиров. клетка:')
        self.fixed_color_label.grid(row=14, column=0, padx=10, sticky="w")
        self.select_color2_btn = Button(self.leftframe, relief="groove", justify=CENTER, width=4, command = lambda: self.choose_color(2))
        self.select_color2_btn.grid(row=14, column=1, columnspan=1, padx=10, pady=3, sticky="nsew")
        self.select_color2_btn.configure(state='normal', fg="#b8b8b8", bg=self.fixedcellcolor)

        self.select_fontcolor2_btn = Button(self.leftframe, relief="groove", justify=CENTER, width=4, command = lambda: self.choose_color(4))
        self.select_fontcolor2_btn.grid(row=14, column=2, columnspan=1, padx=10, pady=3, sticky="nsew")
        self.select_fontcolor2_btn.configure(state='normal', fg="#b8b8b8", bg=self.font_color2)


        self.leftframe.grid_columnconfigure(0, weight=3)
        self.leftframe.grid_columnconfigure(2, weight=3)

        self.gridbutton = ttk.Button(self.leftframe, text='Построить сетку', 
                        command= lambda: self.make_crossword_grid(self.entry1.get()))
        self.gridbutton.grid(row=15, column=0, columnspan=4, pady=3, padx=10, sticky="nsew")

        self.generatorbutton = ttk.Button(self.leftframe, text='Заполнить сетку', 
                                command=lambda: self.notifiationlabel.config(
                                    text='Словарь не загружен\nили не построена сетка', 
                                style="notificationlabel.TLabel", foreground='red'))
        self.generatorbutton.grid(row=16, column=0, columnspan=4, pady=3, padx=10, sticky="nsew")
        # self.leftframe.grid_rowconfigure(8, weight=1)

        self.maxlabel = ttk.Label(self.leftframe, style="infolabel.TLabel", 
                                   background=self.majorcolor, font=self.font10, text='Перемешать:')
        self.maxlabel.grid(row=17, column=0, columnspan=1, pady=0, padx=10, sticky="w")

        self.mix_check = False
        self.checkbutton = Checkbutton(self.leftframe, bg=self.majorcolor, text='⮁', fg="white",
                        font=self.font13, activebackground=self.majorcolor, command=self.mix_tails)
        self.checkbutton.grid(row=17, column=1, pady=0, padx=0, sticky='nsew')

        self.shiftlabel = ttk.Label(self.leftframe, style="infolabel.TLabel", 
                                   background=self.majorcolor, font=self.font10, text='Сдвиг строки:')
        self.shiftlabel.grid(row=18, column=0, columnspan=1, pady=0, padx=10, sticky="w")
  
        self.left_shift_button = Button(self.leftframe, command=self.left_shift_fixed_cells, text="←", font=self.font13)
        self.left_shift_button.grid(row=18, column=1, columnspan=1, pady=3, padx=10, sticky="nsew")
        
        self.right_shift_button = Button(self.leftframe, command=self.right_shift_fixed_cells, text="→", font=self.font13)
        self.right_shift_button.grid(row=18, column=2, columnspan=1, pady=3, padx=10, sticky="nsew")

        self.savebutton = ttk.Button(self.leftframe, text='Сохранить в PDF', 
                        command=lambda: self.notifiationlabel.config(
                        text='Сетка отсутствует\nили она пуста', foreground='red'))

        self.savebutton.grid(row=19, column=0, columnspan=4, pady=3, padx=10, sticky="nsew")
        # self.leftframe.grid_rowconfigure(9, weight=1)


        # Наполнение правой части
        self.crosswordframe = Frame(self.rightframe)
        self.crosswordframe.grid(row=0, column=0, padx=10, pady=3)
        self.rightframe.grid_rowconfigure(0, weight=1)
        self.rightframe.grid_columnconfigure(0, weight=1)
        

    # === СЛУЖЕБНЫЕ ФУНКЦИИ И ОБРАБОТЧИКИ СОБЫТИЙ =======================================
    # функции оформления entry 1/2
    def on_focus_in(self, entry):
        if entry.cget('state') == 'normal':
            entry.configure(state='normal', fg="black")
            entry.delete(0, 'end')
    def on_focus_out(self, entry, placeholder):
        if entry.get() == "":
            entry.insert(0, placeholder)
            entry.configure(state='normal', fg="#b8b8b8")
    def choose_color(self, num):
        # variable to store hexadecimal code of color
        if num == 1:
            self.cellcolor = colorchooser.askcolor(title ="Choose color")[1]
            self.select_color1_btn.configure(bg=self.cellcolor)
        if num == 2:
            self.fixedcellcolor = colorchooser.askcolor(title ="Choose color")[1]
            self.select_color2_btn.configure(bg=self.fixedcellcolor)
        if num == 3:
            self.font_color1 = colorchooser.askcolor(title ="Choose color")[1]
            self.select_fontcolor1_btn.configure(bg=self.font_color1)
        if num == 4:
            self.font_color2 = colorchooser.askcolor(title ="Choose color")[1]
            self.select_fontcolor2_btn.configure(bg=self.font_color2)    
        
    # функция валидации и ограничения ввода
    def char_valid(self, newval):
        return re.match(u"^\w{0,1}$" , newval, re.UNICODE) is not None # "^\w{0,1}$"
    # функция смены полярности сетки
    def check_change(self):
        if self.check: 
            self.check = False
            self.checkbutton.config(fg="white")
        else: 
            self.check = True
            self.checkbutton.config(fg="red")
        self.make_crossword_grid(self.entry1.get())
    # функция перемешивания хвостов
    def mix_tails(self):
        if self.mix_check: 
            self.mix_check = False
            self.checkbutton.config(fg="white")
        else: 
            self.mix_check = True
            self.checkbutton.config(fg="red")
        self.makeRightSide()
    # функция выбора ячеек
    def cell_picker(self, event, entry, i, j):
        if self.check == False:
            if self.enabledcell[i][j] == '0':
                entry.config(state=NORMAL, validate="key", fg='black', bg=self.cellcolor, readonlybackground=self.cellcolor,
                            validatecommand=self.char_check)
                self.enabledcell[i][j] = '1'
            else:
                entry.delete(0, END)
                entry.config(state=DISABLED, fg='black',
                            validate=None, validatecommand=None)
                self.enabledcell[i][j] = '0'
        else:
            if self.enabledcell[i][j] != '0':
                entry.delete(0, END)
                entry.config(state=DISABLED, fg='black',
                            validate=None, validatecommand=None)
                self.enabledcell[i][j] = '0'
            else:
                entry.config(state=NORMAL, validate="key", fg='black', bg=self.cellcolor, readonlybackground=self.cellcolor,
                            validatecommand=self.char_check)
                self.enabledcell[i][j] = '1'
    # функция фиксации ячеек, заполненных вручную
    def fixing_cell(self, entry, i, j):
        entry.config(validate="key", fg=self.font_color2, readonlybackground=self.fixedcellcolor, 
                    validatecommand=self.char_check)
        self.fixed_phrase[i] = entry.get()

    # функция расфиксации ячеек, заполненных вручную
    # def unfixing_cell(self, event, entry, i, j):
    #     if self.grid[i][j]['readonlybackground'] == self.fixedcellcolor:
    #         entry.delete(0, END)
    #         entry.config(state=NORMAL, validate="key", fg=self.font_color1, bg=self.cellcolor, 
    #                      validatecommand=self.char_check, readonlybackground=self.cellcolor)
    #         self.enabledcell[i][j] = '1'
    #         self.fixed_phrase[i] = '_'

    def phrase_analize(self):
        self.fixed_phrase = self.phrasetext.get(1.0, "end-1c")
        self.fixed_phrase = re.sub(r'\W+', '', self.fixed_phrase)
        self.fixed_phrase = re.sub(r'\d+', '', self.fixed_phrase)
        self.fixed_phrase = list(self.fixed_phrase.upper())
        self.h = len(self.fixed_phrase)
        print(self.fixed_phrase)
    
    def left_shift_fixed_cells(self):
        if self.phrase_position > 0:
            self.phrase_position -= 1
        self.make_crossword_grid(self.entry1.get())

        
    def right_shift_fixed_cells(self):
        if self.phrase_position < self.w-1:
            self.phrase_position += 1
        self.make_crossword_grid(self.entry1.get())
        # for x in range(0, self.h):
        #     for y in range(0, self.w):
        #         if self.grid[x][y]['bg'] == self.fixedcellcolor and y<self.w-1:
        #             temp_value = self.grid[x][y].get()
        #             self.grid[x][y].delete(0, END)
        #             self.grid[x][y]['bg'] = self.cellcolor
        #             self.grid[x][y+1].insert(0, temp_value)
        #             self.grid[x][y+1]['bg'] = self.fixedcellcolor
        #             break


    # === ОСНОВНЫЕ ФУНКЦИИ ПРИЛОЖЕНИЯ ==================================================================
    # функция открытия директории
    def open_directory(self):
        try:
            # выбор и открытие папки
            self.folderpath = fd.askdirectory()
            # self.folderpath = 'D:\my\Programming\crossword-architect'
            self.notifiationlabel.config(text=1*'\n', 
                            style="notificationlabel.TLabel", foreground='red')
            # Чтение данных и сохранение в массивы
            self.dictdir = os.listdir(self.folderpath)
            self.txts = []
            self.dictlistbox.delete(0, END)
            for filename in self.dictdir:
                if filename.endswith('.txt'):
                    self.dictlistbox.insert(END, '- '+filename)
                    self.txts.append(f'{self.folderpath}/{filename}')
            
                self.selectdictbutton.config(command=self.make_dictionary)

        except FileNotFoundError:
            self.selectdictbutton.config(command=None)
            self.notifiationlabel.config(text='Директория не выбрана\n', 
                            style="notificationlabel.TLabel", foreground='red')
    
    # функция загрузки слов из выбранных файлов в память и формирования словаря
    def make_dictionary(self):
        self.dictionary = {}
        self.wordsarray = []
        self.selected = self.dictlistbox.curselection()
        if len(self.selected) != 0:
            for fileindex in self.selected:
                with open(self.txts[fileindex], 'r', encoding='utf-8') as file:
                    for word in file:
                        word = word.replace('\n', '')
                        word = word.replace(' ', '-')
                        word = word.upper()
                        # проверка на наличие нежеланных символов (' ', '-')
                        if '-' in word:
                            splitedwords = word.split('-')
                        else:
                            splitedwords = [word]
                        for sw in splitedwords:
                            if sw != '' and sw != ' ' and len(sw)>1 and sw not in self.wordsarray:
                                l = len(sw)
                                self.wordsarray.append(sw)
                                if l not in self.dictionary.keys():
                                    self.dictionary[l] = []
                                temp = self.dictionary[l]
                                temp.append(sw)
                                self.dictionary[l] = temp
            print(self.dictionary.keys())
            #print(self.dictionary.keys())
            self.shortest = min(self.dictionary.keys())
            self.longest = max(self.dictionary.keys())
            
            try:
                if self.w > 0:
                    self.generatorbutton.config(command=self.generator)
            except:
                pass
            finally:
                self.notifiationlabel.config(text='Словарь загружен\n', 
                                style="notificationlabel.TLabel", foreground=self.approvecolor)
        else:
            self.shortest = '-'
            self.longest = '-'
            self.infolabel.config(text=2*'\n')  
            self.notifiationlabel.config(text='Словари не выбраны\n', 
                            style="notificationlabel.TLabel", foreground='red')
        
        self.infolabel.config(text=f' Общее кол-во слов: {len(self.wordsarray)}\n Мин. длина слова: {self.shortest}\n Макс. длина слова: {self.longest}')

    # функция создания сетки
    def make_crossword_grid(self, w):
        self.cellfont = self.font_combobox.get()
        if w.isdigit():
            try:
                for widget in self.crosswordframe.winfo_children():
                    widget.destroy()
            finally:
                self.w = int(w)
                self.phrase_analize()
                self.grid = []
                self.enabledcell = []
                #self.crosswordframe.config(width=self.w*2, height=self.h*5)
                self.fontcoeff = round( 450/max([self.w, self.h]) )
                
                for i in range(self.h):
                    self.grid.append([])
                    self.enabledcell.append([])
                    for j in range(self.w):
                        if not self.check:
                            tempobj = Entry(self.crosswordframe, justify=CENTER, font=(self.cellfont, self.fontcoeff, 'bold'), fg=self.font_color1, 
                                        relief='solid', width=2, state=NORMAL, validate="key", bg=self.cellcolor, borderwidth=self.bd_scale.get(),
                                        validatecommand=self.char_check, readonlybackground=self.cellcolor, disabledbackground=self.disabledentrycolor)
                            self.enabledcell[i].append('1')
                        else:
                            tempobj = Entry(self.crosswordframe, justify=CENTER, font=(self.cellfont, self.fontcoeff, 'bold'),
                                        relief='solid', width=2, state=DISABLED, readonlybackground=self.cellcolor, bg=self.cellcolor, validate=None, 
                                        validatecommand=None, disabledbackground=self.disabledentrycolor) 
                            self.enabledcell[i].append('0')
                        tempobj.grid(row=i, column=j, sticky="nsew")
                        self.grid[i].append(tempobj)
                        self.grid[i][j].insert(0, '')
                        # self.grid[i][j].bind('<Button-3>', lambda x, entry=self.grid[i][j], i=i, j=j: self.cell_picker(x, entry, i, j))
                        # self.grid[i][j].bind('<Key>', lambda x, entry=self.grid[i][j], i=i, j=j: self.fixing_cell(x, entry, i, j))
                        # self.grid[i][j].bind('<BackSpace>', lambda x, entry=self.grid[i][j], i=i, j=j: self.unfixing_cell(x, entry, i, j))
                        # self.grid[i][j].bind('<Delete>', lambda x, entry=self.grid[i][j], i=i, j=j: self.unfixing_cell(x, entry, i, j))
                        if j == self.phrase_position:
                            self.grid[i][j].insert(0, self.fixed_phrase[i])
                            self.fixing_cell(self.grid[i][j], i, j)
                        # self.grid[i][j].configure(state="readonly")
                        
                        self.crosswordframe.grid_rowconfigure(i, weight=1)
                        self.crosswordframe.grid_columnconfigure(j, weight=1)
                self.notifiationlabel.config(text='\n', style="notificationlabel.TLabel", foreground='red')

                try:
                    if len(self.dictionary) != 0:
                        self.generatorbutton.config(command=self.generator)
                except:
                    self.notifiationlabel.config(text='Словарь не загружен!\n',
                                    style="notificationlabel.TLabel", foreground='red')
        else:
            self.notifiationlabel.config(text='Укажите размер сетки\n', 
                            style="notificationlabel.TLabel", foreground='red')
            self.generatorbutton.config(command=None)
    
    # функция очистки сетки от букв
    def clear_grid(self):
        for x in range(0, self.h):
            for y in range(0, self.w):
                if self.enabledcell[x][y] != '0':
                    if self.grid[x][y]['readonlybackground'] != self.fixedcellcolor:
                        self.grid[x][y].delete(0, END)
    
    # функция очистки enabledcell
    def clear_enabledcell_list(self):
        for x in range(0, self.h):
            for y in range(0, self.w):
                if self.enabledcell[x][y] != '0':
                    self.enabledcell[x][y] = '1'

    # функция настройки паддингов для enabledcell
    def set_paddings(self, mode):
        if mode == 'set':
            # добавление паддингов
            self.enabledcell.insert(0, ['0' for i in range(self.w)])
            for x in range(self.h+1):
                self.enabledcell[x].insert(0, '0')
                self.enabledcell[x].append('0')
            self.enabledcell.append(['0' for i in range(self.w+2)])
        if mode == 'del':
            # удаление паддингов
            self.enabledcell.pop(0)
            for x in range(self.h+1):
                self.enabledcell[x].pop(0)
                self.enabledcell[x].pop()
            self.enabledcell.pop()
    
    # опциональная функция вывода результатов анализа сетки
    def show_analize_results(self, turn='on'):
        """turn: on/off (default "on")"""
        if turn == 'on':
            print('\n=== Результат анализа сетки ===')        
            for x in range(len(self.enabledcell)):
                for y in range(len(self.enabledcell[x])):
                    print(self.enabledcell[x][y].center(2), end=' ')
                print()
            try:
                print('h_words:', self.h_params)
                print('v_words:', self.v_params)
            except AttributeError:
                pass
            print('=== ======================= ===')

    # функция анализа сетки
    def analize_grid(self):
        """
        Aнализ 0 | 1 -> E, H, V, VH, vH, Vh, h, v, +
        Horizontal: H VH vH Vh h +
        Vertical: V VH vH Vh v +
        Other: E
        """
        self.h_params = []
        self.v_params = []
        self.fixed_words = []
        self.max_length = 0
        self.h_signs = ('1', 'h', 'H', 'VH', 'Vh', 'vH', '+')
        self.v_signs = ('1', 'v', 'V', 'VH', 'Vh', 'vH', '+')
        # горизонталь
        for x in range(1, len(self.enabledcell)-1):
            l = 0
            intersection = 0
            status = "c"
            for y in range(1, len(self.enabledcell[x])-1):
                if self.enabledcell[x][y] in self.h_signs:
                    upper = self.enabledcell[x-1][y]
                    lower = self.enabledcell[x+1][y]
                    left = self.enabledcell[x][y-1]
                    right = self.enabledcell[x][y+1]

                    # Empty 
                    if upper == '0' and lower == '0' and right == '0' and left == '0':
                        self.enabledcell[x][y] = 'E'
                    if left == '0'  and right in self.h_signs:
                        # Начало горизонтального
                        if upper == '0' and lower == '0':
                            self.enabledcell[x][y] = 'H'
                        # Начало горизонтального и вертикального Г
                        if upper == '0' and lower in self.v_signs:
                            self.enabledcell[x][y] = 'VH'
                            intersection += 1
                        # Начало горизонтального из буквы вертикального |-
                        if upper in self.v_signs:
                            self.enabledcell[x][y] = 'vH'
                            intersection += 1
                        self.h_params.append([])
                        self.h_params[-1].append(x-1)
                        self.h_params[-1].append(y-1)
                        l += 1
                    if left in self.h_signs:
                        # Начало вертикального из буквы горизонтального T
                        if upper == '0' and lower in self.v_signs:
                            self.enabledcell[x][y] = 'Vh'
                            intersection += 1
                            
                        # Продолжение горизонтального --
                        if upper == '0' and lower == '0':
                            self.enabledcell[x][y] = 'h'
                            if self.grid[x-1][y-1]['readonlybackground'] == self.fixedcellcolor:
                                status = "f"
                        # Пересечение слов в любом месте
                        if upper in self.v_signs:
                            self.enabledcell[x][y] = '+'
                            intersection += 1
                        l += 1
                        if right == '0':
                            self.h_params[-1].append(intersection)
                            self.h_params[-1].append(l)
                            self.h_params[-1].append('h')
                            self.h_params[-1].append(status)
                            if self.max_length < l:
                                self.max_length = l
                            l = 0
                            intersection = 0
                            status = "c"

    # функция установки ограничения на итерации
    def set_interation_limit(self, limit):
        return 10
    
    # функция настройки конфигов сетки
    def set_config(self, X, Y, length, position):
        for l in range(length):
            if position in ('h','H') and self.enabledcell[X][Y+l] in self.h_signs:
                if self.grid[X][Y+l]['readonlybackground'] != self.fixedcellcolor:
                    self.grid[X][Y+l].config(readonlybackground=self.deniedcolor)
            if position in ('v','V') and self.enabledcell[X+l][Y] in self.v_signs: 
                if self.grid[X+l][Y]['readonlybackground'] != self.fixedcellcolor:
                    self.grid[X+l][Y].config(readonlybackground=self.deniedcolor)
    
    # функция добавления полученных слов в массивы
    def word_adding(self, word, position, status='c'):
        if status == 'c':
            if position in ('h','H'):
                self.h_words.append(word)
            if position in ('v','V'):
                self.v_words.append(word)
        else:
            if position in ('h','H'):
                self.h_words.append(word+' (f)')
            if position in ('v','V'):
                self.v_words.append(word+' (f)')
        
    # функция подбора слов и заполнения сетки
    def word_randomizer(self, X, Y, length, position, status):
        if status == "c":
            pattern = ''
            words_with_fixed_len = '\n'.join(self.dictionary[length])
            for l in range(length):
                if position in ('h','H'):
                    char = self.grid[X][Y+l].get()
                    if char == '':
                        pattern += '.'
                    else:
                        pattern += char
                if position in ('v','V'):
                    char = self.grid[X+l][Y].get()
                    if char == '':
                        pattern += '.'
                    else:
                        pattern += char
            pattern = re.compile(pattern)
            result = pattern.findall(words_with_fixed_len)
            try:
                # генерация слова с учетом неправильных начальных букв (Ъ, Ь)
                is_not_okay = True
                while is_not_okay:
                    is_not_okay = False
                    word = result[random.randint(0, len(result)-1)]
                    if word in (self.h_words + self.v_words):
                        result.remove(word)
                        is_not_okay = True
                        continue
                        # if len(result) > 1 and repeat <= 10:
                        #     is_not_okay = True
                        #     repeat += 1
                        #     continue
                        # else:
                        #     raise 
                    if 'ь' in word or 'ъ' in word:
                        if position in ('h','H'):
                            for l in range(length):
                                if word[l] in self.wrong_letters and self.enabledcell[X][Y+l] == 'Vh':
                                    is_not_okay = True
                                    break
                        if position in ('v','V'):
                            for l in range(length):
                                if word[l] in self.wrong_letters and self.enabledcell[X+l][Y] == 'vH':
                                    is_not_okay = True
                                    break
                # побуквенная вставка в сетку
                for l in range(length):
                    if position in ('h','H'):
                        self.grid[X][Y+l].insert(0, word[l])
                    if position in ('v','V'):
                        self.grid[X+l][Y].insert(0, word[l])
                # добавление слов в H/V списки
                self.word_adding(word, position)
                #print('len', length, pattern, word, position)
            except:
                self.stop = False
                self.empty += 1
                if self.iteration == self.iteration_limit:
                    self.set_config(X, Y, length, position)
        # обработка фиксированных слов
        else:
            word = ''
            for l in range(length):
                if position in ('h','H'):
                    word = word + self.grid[X][Y+l].get()
                if position in ('v','V'):
                    word = word + self.grid[X+l][Y].get()
            self.word_adding(word, position, status)

    def makeRightSide(self):
        # перемешиваем хвосты
        new_tails_list = copy.deepcopy(self.word_tails)
        if self.mix_check:
            random.shuffle(new_tails_list)
        max_tail_length = 0
        for tail in new_tails_list:
            if len(tail) > max_tail_length:
                max_tail_length = len(tail)

        # создание и заполнение доп сетки
        for i in range(len(new_tails_list)):
            word = new_tails_list[i]
            while len(word) < max_tail_length:
                word+=' '
            for j in range(self.w):
                try:
                    print(self.grid[i][self.w+j+2].get(), j)
                    self.grid[i][-1].configure(state="readonly")
                except: pass
            for j in range(len(word)):
                tempobj = Entry(self.crosswordframe, justify=CENTER, font=(self.cellfont, self.fontcoeff, 'bold'), fg=self.font_color1, 
                            relief='solid', width=2, validate="key", bg=self.cellcolor, borderwidth=self.bd_scale.get(), state="normal",
                            validatecommand=self.char_check, readonlybackground=self.cellcolor, disabledbackground=self.disabledentrycolor)
                tempobj.grid(row=i, column=self.w+j+2, sticky="nsew")
                self.grid[i].append(tempobj)
                self.grid[i][-1].insert(0, word[j])
                self.grid[i][-1].configure(state="readonly")

                self.crosswordframe.grid_rowconfigure(i, weight=1)
                self.crosswordframe.grid_columnconfigure(self.w+j+2, weight=1)
            

    # основная функция генерации кроссворда
    def generator(self):
        # self.notifiationlabel.config(text='Ожидание...\n', foreground=self.waitcolor)
        self.clear_enabledcell_list()

        for i in range(self.h):
            for j in range(self.w):
                self.grid[i][j].configure(state="normal")

        # анализ сетки
        self.show_analize_results(turn='off') # on/off
        self.set_paddings('set')
        self.analize_grid()
        self.set_paddings('del')
        self.show_analize_results(turn='on') # on/off

        # сортировка
        self.sum_params = self.h_params + self.v_params
        self.sum_params.sort(key = lambda x: x[3], reverse=True) # по длине
        self.sum_params.sort(key = lambda x: x[2], reverse=True) # по кол-ву пересечений

        # настройка количества итераций
        self.iteration = 1
        self.stop = False
        self.iteration_limit = self.set_interation_limit(10)
        for x in range(len(self.enabledcell)):
            for y in range(len(self.enabledcell[x])):
                if self.enabledcell[x][y] in (self.h_signs + self.v_signs):
                    if self.grid[x][y]['readonlybackground'] != self.fixedcellcolor:
                        self.grid[x][y].config(readonlybackground=self.cellcolor)
        
        # алгоритм генерации и заполнения слов
        start_time = time.time()
        self.min_empty_count = 10
        best_iteration_count = 0
        self.wrong_letters = ('ь', 'ъ')
        if self.max_length <= self.longest:
            while not self.stop and self.iteration <= self.iteration_limit:
                self.clear_grid()
                self.h_words = []
                self.v_words = []
                self.empty = 0
                self.stop = True
                #print(f'=== Итерация #{self.iteration} ===')
                for word in (self.sum_params):
                    self.word_randomizer(word[0], word[1], word[3], word[4], word[5])
                if self.empty <= self.min_empty_count:
                    if self.min_empty_count == self.empty:
                        best_iteration_count += 1
                    else:
                        best_iteration_count = 1
                        self.min_empty_count = self.empty
                self.iteration += 1
            # self.show_messagebox(time.time()-start_time, best_iteration_count)
            test_set = set(self.h_words)
            print(f'\n{len(self.h_words)} ({len(test_set)}) horizontal {self.h_words}')
            test_set = set(self.v_words)
            print(f'{len(self.v_words)} ({len(test_set)}) vertical {self.v_words}')
            
            self.savebutton.config(command=self.save_in_file)
        else:
            self.h_words = []
            self.v_words = []
            self.notifiationlabel.config(
                                text='В словаре отсутствуют\nслова такой длины!', 
                                foreground=self.deniedcolor)
        
        # создание колонки справа
        self.word_tails = []
        for i in range(self.h):
            # интервал между колонками
            self.img_label = Entry(self.crosswordframe, justify=CENTER, font=(self.cellfont, self.fontcoeff, 'bold'), 
                                relief='solid', width=2, state=DISABLED, validate="key", readonlybackground=self.cellcolor, 
                                disabledbackground=self.disabledentrycolor,)
            self.img_label.grid(row=i, column=self.w+1, sticky="nsew")
            
            # img = ImageTk.PhotoImage(file="uzor.png")
            # img_label = Button(self.crosswordframe, border=1, image=img)
            # img_label.grid(row=i, column=self.w+1)
            
            self.fixed_letter_position = -1
            for j in range(self.w):
                if self.grid[i][j]['readonlybackground'] == self.fixedcellcolor:
                    self.fixed_letter_position = j
                    break
            try:
                if self.fixed_letter_position != -1:
                    self.word_tails.append(self.h_words[i][self.fixed_letter_position:])
            except: pass
            for j in range(self.w-self.fixed_letter_position):
                if self.fixed_letter_position != -1:
                    self.grid[i][self.fixed_letter_position+j].delete(0)
        
        self.makeRightSide()
        
        for i in range(self.h):
            for j in range(self.w):
                self.grid[i][j].configure(state="readonly")

    # функция сохранения в pdf файл
    def save_in_file(self):
        self.wm_geometry("+%d+%d" % (40, 5))

        self.savefolderpath = fd.askdirectory()

        x1 = self.winfo_x() + self.rightframe.winfo_x() + self.crosswordframe.winfo_x() + 4
        y1 = self.winfo_y() + self.rightframe.winfo_y() + self.crosswordframe.winfo_y() + 27
        x2 = x1 + self.crosswordframe.winfo_width() + 8
        y2 = y1 + self.crosswordframe.winfo_height() + 9

        snapshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        enhancer = ImageEnhance.Sharpness(snapshot)
        snapshot_enhanced = enhancer.enhance(2)
        
        counter = 1
        while counter < 100:
            file_name = f'{self.w}x{self.h}_crossword_[{len(self.h_words)}h, {len(self.v_words)}v]_{counter}.pdf'
            if not os.path.exists(f'{self.savefolderpath}/{file_name}'):
                snapshot_enhanced.save(f'{self.savefolderpath}/{file_name}', 
                        format='PDF', quality=200)
                # document.save(f'{self.savefolderpath}/{file_name}')
                # os.remove(tempscreenpath)
                break
            else:
                counter += 1



if __name__ == "__main__":
    # print(time.time())
    # if (time.time()) < 1716816043.604615 + 800000: # program blocker
    main = Main()
    main.geometry(f'{1024}x{685}') # main.winfo_screenheight()
    main.wm_geometry("+%d+%d" % (40, 5))
    main.title('CW Architect 2')
    main['bg'] = 'white'
    #main.attributes('-fullscreen', True)


    main.mainloop()

    # pyinstaller -F -w -i 'ico.png' script.py