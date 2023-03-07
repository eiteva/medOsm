from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import tkinter.font as tkFont

import xml.etree.ElementTree as ET 
import csv
import pandas as pd

class WmDialog:
        def __init__(self):
            self.root = Tk()
            w = self.root.winfo_screenwidth()#ширина экрана
            h = self.root.winfo_screenheight()#высота
            w = w//2
            h = h//2
            w = w - 325 # смещение от середины
            h = h - 346

            
            self.root.geometry('693x650+{}+{}'.format(w, h))
            self.root.resizable(False, False)
            self.root.update_idletasks() #принудительная обработка 
            self.root.overrideredirect(True) #откл стандартную панель!!
            self.root['bg'] = '#1A2C34'

                #фон-картинка bg.png
            self.bg_img = PhotoImage(file = 'bg.png')
            self.bg_label = Label(self.root, image=self.bg_img).grid(row=0, column=0, rowspan = 13, columnspan = 3, sticky=N+W+E+S)

            self.entlog = ttk.Entry(self.root, width=20)
            self.entlog.grid(row=5, column=0, sticky=N, pady=10)
            
            self.entpass = ttk.Entry(self.root, width=20)
            self.entpass.grid(row=6, column=0, sticky=N, pady=10)

            self.btnOk = ttk.Button(self.root,
                                    text="Войти",
                                    command = self.checklog)
            self.btnOk.grid(row=7, column=0, sticky=N)

            self.btnCancel = ttk.Button(self.root,
                                        text="Выход",
                                        command=self.quit)
            self.btnCancel.grid(row=7, column=1, sticky=N)

            self.inserttext()

            self.root.bind('<Return>', self.callback) #обрабатывает нажатие Enter

        def callback(self, event):
            """обрабатывает нажатие Enter"""
            self.checklog()


        def checklog(self):
            """Проверяет логин и пароль"""
            global login # глобальная, чтобы показать/скрыть админку в mainw
            login = self.entlog.get()
            password = self.entpass.get()
            print(login)
            print(password)

            d = {} #словарь {логин: пароль}
            with open('passwords.txt') as file:
                for line in file:
                    key, value = line.split()
                    d[key] = value
            a = login in d
            if a == True:
                if d[login] == password:
                    print('login succsess')
                    self.root.destroy()
            else:
                messagebox.showinfo('Ошибка!', "Введен неверный логин или пароль. Обратитесь к администратору")


        def quit(self):
            self.root.destroy()###заменить на exit

        def inserttext(self):
            """заполняет поля ввода """ 
            self.entlog.insert(0, 'admin')
            self.entpass.insert(0, '666ad')

        def liveWnd (self): 
            self.root.mainloop()

class wmain:

        def __init__(self):
            
            self.main = Tk()
            self.main.title('МедОсмотр')
            self.main.iconbitmap("icon.ico")
            self.main.state('zoomed')
            self.main['bg'] = 'WhiteSmoke' #'#1A2C34'#

            self.create_menu()

            fontStyle=tkFont.Font(family="Lucida Grande", size=14)

            #==========панель инструментов===============================================================
            self.tool_bar = Frame(self.main, bg="#A1A1A1", bd=1, relief=RAISED)
            #self.tool_bar.place(relx=0, rely=0)
            self.tool_bar.pack(fill=X)
            #self.tool_bar.grid(row=0, column=0)

            self.img_open = PhotoImage(file = 'open_excel.png')
            self.but_open = Button(self.tool_bar, image=self.img_open, command=self.open_xls_file)
            self.but_open.bind("<Enter>", lambda event: self.status_label.configure(text="Открыть XLS документ"))
            self.but_open.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_open.pack(side=LEFT)

            self.img_open_excel = PhotoImage(file = 'Open.png')
            self.but_open_excel = Button(self.tool_bar, image=self.img_open_excel, command=self.open_xml_file)
            self.but_open_excel.bind("<Enter>", lambda event: self.status_label.configure(text="Открыть XML документ - счет-реестр"))
            self.but_open_excel.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_open_excel.pack(side=LEFT)

            self.img_delete = PhotoImage(file='Delete.png')
            self.but_del = Button(self.tool_bar, image=self.img_delete, command=self.clear_workspace)
            self.but_del.bind("<Enter>", lambda event: self.status_label.configure(text="Очистить всё"))
            self.but_del.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_del.pack(side=LEFT)

            self.img_save = PhotoImage(file = 'Save.png')
            self.but_save = Button(self.tool_bar, image=self.img_save, command=self.save_xml)
            self.but_save.bind("<Enter>", lambda event: self.status_label.configure(text="Сохранить документ"))
            self.but_save.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_save.pack(side=LEFT)

            self.img_write = PhotoImage(file = 'write.png')
            self.but_write = Button(self.tool_bar, image=self.img_write, command=self.append_to_reestr)
            self.but_write.bind("<Enter>", lambda event: self.status_label.configure(text="Записать в реестр"))
            self.but_write.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_write.pack(side=LEFT)

            self.img_about = PhotoImage(file = 'About.png')
            self.but_about = Button(self.tool_bar, image=self.img_about, command=self.show_info)
            self.but_about.bind("<Enter>", lambda event: self.status_label.configure(text="Информация о документе"))
            self.but_about.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_about.pack(side=LEFT)

            self.img_help = PhotoImage(file='Help.png')
            self.but_help = Button(self.tool_bar, image=self.img_help, command=self.show_help)
            self.but_help.bind("<Enter>", lambda event: self.status_label.configure(text="Помощь"))
            self.but_help.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_help.pack(side=LEFT)

            self.img_exit = PhotoImage(file='Exit.png')
            self.but_exit = Button(self.tool_bar, image=self.img_exit, command=self.close)
            self.but_exit.bind("<Enter>", lambda event: self.status_label.configure(text="Выход"))
            self.but_exit.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_exit.pack(side=LEFT)

            self.img_theme = PhotoImage(file='day_night.png')
            self.but_theme = Button(self.tool_bar, image=self.img_theme, command=self.day_night)
            self.but_theme.bind("<Enter>", lambda event: self.status_label.configure(text="Изменить тему оформления"))
            self.but_theme.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_theme.pack(side=LEFT)



            #---------------верхний фрейм - настройки----------------------------------------------------------------
            self.f_top = LabelFrame(self.main, text="Настройки документа", bg="WhiteSmoke")
            #self.f_top.place(relx=0.05, rely=0.1) 
            self.f_top.pack(fill=X)#grid()#
            ##self.label_top = Label(self.f_top, text="hehehe")
            ##self.label_top.grid()

            #==========Статус-бар==================================================================================
            self.status_bar = Frame(self.main, bg="#A1A1A1", bd=1, relief=RAISED)
            #self.status_bar.place(relx=0, rely=1, anchor=SW)
            self.status_bar.pack(side=BOTTOM, fill=X)
            #self.status_bar.grid(row=11, column=0)

            self.status_label = Label(self.status_bar, font=7, bg='#A1A1A1', text='Начните работу с открытия файла')
            self.status_label.pack(side=LEFT, fill=X)
            #self.status_label.grid()            


            #-------нижний фрейм - вывод исходный документ-----------------------------------------------------------------------------------
            self.f_bot = ttk.LabelFrame(self.main, text="Содержимое документа", height=15, width=50)
            #self.f_bot.place(relx=0, rely=0.5)
            self.f_bot.pack(side=BOTTOM, fill=X)#grid()

            #---панель инструментов для таблицы------------------------------------
            self.tv_tool_bar = Frame(self.f_bot, bg="#A1A1A1", bd=1, relief=RAISED)
            self.tv_tool_bar.pack(side=TOP, fill=X)

            self.img_refresh = PhotoImage(file='Refresh.png')
            self.but_refresh_tv = Button(self.tv_tool_bar, image=self.img_refresh, command=self.tv_refresh)
            self.but_refresh_tv.bind("<Enter>", lambda event: self.status_label.configure(text="Показать исходные данные"))
            self.but_refresh_tv.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_refresh_tv.pack(side=LEFT)

            self.img_clean = PhotoImage(file='Delete.png')
            self.but_clean = Button(self.tv_tool_bar, image=self.img_clean, command=self.clear_tv)
            self.but_clean.bind("<Enter>", lambda event: self.status_label.configure(text="Очистить таблицу"))
            self.but_clean.bind("<Leave>", lambda event: self.status_label.configure(text="Наведите на элемент, чтобы получить справку"))
            self.but_clean.pack(side=LEFT)

            self.img_compare = PhotoImage(file='compare.png')
            self.but_compare = Button(self.tv_tool_bar, image=self.img_compare, command=self.compare)
            self.but_compare.bind("<Enter>", lambda event: self.status_label.configure(text='Сравнить с паспортом педиатрического участка'))
            self.but_compare.bind("<Leave>", lambda event: self.status_label.configure(text='Наведите на элемент, чтобы получить справку'))
            self.but_compare.pack(side=LEFT)

            self.img_find = PhotoImage(file='Find.png')
            self.but_find = Button(self.tv_tool_bar, image=self.img_find, command = self.find_window)
            self.but_find.bind("<Enter>", lambda event: self.status_label.configure(text='Найти элемент в таблице'))
            self.but_find.bind("<Leave>", lambda event: self.status_label.configure(text='Наведите на элемент, чтобы получить справку'))
            self.but_find.pack(side=LEFT)

            self.img_save_df = PhotoImage(file='save_df.png')
            self.but_save_df = Button(self.tv_tool_bar, image=self.img_save_df)#, command=self.save_df)
            self.but_save_df.bind("<Enter>", lambda event: self.status_label.configure(text='Сохранить текущий результат'))
            self.but_save_df.bind("<Leave>", lambda event: self.status_label.configure(text='Наведите на элемент, чтобы получить справку'))
            self.but_save_df.pack(side=LEFT)

            #----------------------------------------------------------------------

            self.scrlv = ttk.Scrollbar(self.f_bot, orient="vertical")
            self.scrlh = ttk.Scrollbar(self.f_bot, orient="horizontal")
                        
            self.treeview = ttk.Treeview(
                                    self.f_bot,
                                    yscrollcommand=self.scrlv.set, 
                                    xscrollcommand=self.scrlh.set,
                                    show="headings", 
                                    selectmode="browse",
                                    height=15
                                    )
            self.treeview.pack(side=RIGHT)#, fill=BOTH)#grid(row=0, column=0, rowspan=4, columnspan=1)
           
            self.scrlv.config(command=self.treeview.yview)
            self.scrlh.config(command=self.treeview.xview)

            self.scrlv.pack(side=LEFT, fill=Y)#grid(row=0, column=2, rowspan=4, sticky=N+S+E+W)
            self.scrlh.pack(side=TOP, fill=X)#grid(row=5, column=0, columnspan=2, sticky=N+S+E+W)


        def open_xml_file(self, event=None):
            """Открывает и парсит указанный .xml файл"""
            #global xmlfile
            xmlfile = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("","*.xml"), ("all files","*.*")))
            tree = ET.parse(xmlfile)
            global root 
            root = tree.getroot()   #получение корневого дерева тегов
            temp="temp.csv"
            self.file_to_csv(root, temp)


        def file_to_csv(self, root, temp):
            """парсит .xml, создает в корневой папке временный temp.csv файл с таблицей данных"""
            

            #=========================create a scv-file for writting
            list_data = open(temp, 'w') #, newline='', encoding='utf-8')
            csvwriter = csv.writer(list_data)

            title_head = []          #пустой список Заголовки столбцов

            count = 0
            count_zap = 0

            try:
                for element in root[1].findall('ZAP'):     ######### root.index = [1][5]       все элементы внутри ZAP
                    resident = []
                    usl_list = []
                    usl_f_list = []
                    usl_dat_list = []
                    sumv_usl_list = []
                    usl_krat_list = []
                    usl_osn_list = []
                    is_ext_usl_list = []
                    usl_prvs_list = []
                    usl_kodvr_list = []
                    usl_vr_snils_list = []
                    usl_pr_nep_list = []
                    commentu_list = []

                    if count == 0:
                        
                        CODE_STR = element.find('CODE_STR').tag
                        title_head.append(CODE_STR)
                        PR_NOV = element.find('PR_NOV').tag
                        title_head.append(PR_NOV)

                        #<PACIENT>
                            #<DPFS>
                        for child in root.find('.//DPFS'): ### все элементы внутри DPFS
                            #NPOLIS = child.find('.//NPOLIS').tag
                            #title_head.append(NPOLIS)
                            if child.tag == 'NPOLIS':
                                title_head.append(child.tag)
                            elif child.tag == 'ENP':
                                title_head.append(child.tag)
                                #<MTR>
                            elif child.tag == 'MTR':
                                for child in root.find('.//MTR'):
                                    title_head.append(child.tag)
                                #</MTR>
                            elif child.tag == 'COMMENTD':
                                title_head.append(child.tag) 
                            else:
                                continue
                            #</DPFS>

                            #<PERS>
                        for child in root.find('.//PERS'): ### все элементы внутри PERS
                            title_head.append(child.tag)
                            #</PERS>

                            #<PACIENT_TO>
                                #<CLS_PAC>
                        for child in root.find('.//CLS_PAC'):
                            title_head.append(child.tag)
                                #</CLS_PAC>
                            #</PACIENT_TO>
                        #</PACIENT>

                        #---------------------------------------------------------------
                        #<SLUCH>
                        ID_SL = element.find('.//ID_SL').tag
                        title_head.append(ID_SL)
               
                            #<SL_COM>
                        DATE_1 = element.find('.//DATE_1').tag
                        title_head.append(DATE_1)
                        DATE_2 = element.find('.//DATE_2').tag
                        title_head.append(DATE_2)
                        DLIT = element.find('.//DLIT').tag
                        title_head.append(DLIT)
                        NHISTORY = element.find('.//NHISTORY').tag
                        title_head.append(NHISTORY)
                        USL_OK = element.find('.//USL_OK').tag
                        title_head.append(USL_OK)
                        VID_MP = element.find('.//VID_MP').tag
                        title_head.append(VID_MP)
                        RSLT = element.find('.//RSLT').tag
                        title_head.append(RSLT)
                        KODVR = element.find('.//KODVR').tag
                        title_head.append(KODVR)
                        VR_SNILS = element.find('.//VR_SNILS').tag
                        title_head.append(VR_SNILS)
                        PR_REAB = element.find('.//PR_REAB').tag
                        title_head.append(PR_REAB)
                        IDSP = element.find('.//IDSP').tag
                        title_head.append(IDSP)
                        PR_DS_ONK = element.find('.//PR_DS_ONK').tag
                        title_head.append(PR_DS_ONK)

                                #<SLUCH_STM>
                        for child in root.find('.//SLUCH_STM'):
                            title_head.append(child.tag)
                                #</SLUCH_STM>
                                #<CLS_MKB>
                        for child in root.find('.//CLS_MKB'):
                            title_head.append(child.tag)
                                #</CLS_MKB>
                                #<CLS_MUSL>
                        for child in root.find('.//CLS_MUSL'):
                            title_head.append(child.tag)
                                #</CLS_MUSL>
                            #</SL_COM>
               
                            #<SLS_DD>
                        for child in root.find('.//SL_DD'):
                            title_head.append(child.tag)
                            #</SLS_DD>

                        #</SLUCH>
                    
                        count = count + 1 #################число записей=пациентов
                        csvwriter.writerow(title_head)   #запись списка заголовков в csv

                #================== заполнение таблицы .csv данными (.text)======================
                #NPOLIS,ENP,VPOLIS,TF_OKATO,SMO_OGRN,SMO_OKATO,SMO_NAME,DATE_N,COMMENTD,ID_PAC,NOVOR,FAM,IM,OT,W,DR,MR,DOCTYPE,DOCSER,DOCNUM,OKATOG,OKATOP,COMMENTPE,SPR_SOC,SPR_LGT,SPR_INVLD,PR_INV_PERV    
                    try:
                        CODE_STR = element.find('CODE_STR').text
                        resident.append(CODE_STR)
                        PR_NOV = element.find('PR_NOV').text
                        resident.append(PR_NOV)
                        #<PACIENT>
                            #<DPFS>
                        npolis = element.find('.//NPOLIS').text
                        resident.append(npolis)
                        enp = element.find('.//ENP').text
                        resident.append(enp)
                                #<MTR>
                        for child in element.find('.//MTR'):
                            resident.append(child.text)
                                #</MTR>
                        commentd = element.find('.//COMMENTD').text
                        resident.append(commentd) 
                            #</DPFS>

                            #<PERS>
                        for child in element.find('.//PERS'): ### все элементы внутри PERS
                            resident.append(child.text)
                            #</PERS>

                            #<PACIENT_TO>
                                #<CLS_PAC>
                        for child in element.find('.//CLS_PAC'):
                            resident.append(child.text)
                                #</CLS_PAC>
                            #</PACIENT_TO>
                        #</PACIENT>
                    except AttributeError:
                        resident.append('NONE')

                        #<SLUCH>
                    try:
                        id_sl = element.find('.//ID_SL').text
                        resident.append(id_sl)

                            #<SL_COM>
                        date_1 = element.find('.//DATE_1').text
                        resident.append(date_1)
                        date_2 = element.find('.//DATE_2').text
                        resident.append(date_2)
                        DLIT = element.find('.//DLIT').text
                        resident.append(DLIT)
                        NHISTORY = element.find('.//NHISTORY').text
                        resident.append(NHISTORY)
                        USL_OK = element.find('.//USL_OK').text
                        resident.append(USL_OK)
                        VID_MP = element.find('.//VID_MP').text
                        resident.append(VID_MP)
                        RSLT = element.find('.//RSLT').text
                        resident.append(RSLT)
                        KODVR = element.find('.//KODVR').text
                        resident.append(KODVR)
                        vr_snils = element.find('.//VR_SNILS').text
                        resident.append(vr_snils)
                    except AttributeError:
                        resident.append('NONE')

                    try:
                        PR_REAB = element.find('.//PR_REAB').text
                        resident.append(PR_REAB)
                    except AttributeError:
                        resident.append('NONE')
                    try:
                        IDSP = element.find('.//IDSP').text
                        resident.append(IDSP)
                    except AttributeError:
                        resident.append('NONE')
                    try:
                        PR_DS_ONK = element.find('.//PR_DS_ONK').text
                        resident.append(PR_DS_ONK)
                    except AttributeError:
                        resident.append('NONE')

                            #<SLUCH_STM>
                    try:
                        for child in element.find('.//SLUCH_STM'):
                            resident.append(child.text)
                    except AttributeError:
                        resident.append('NONE')
                                #</SLUCH_STM>
                                #<CLS_MKB>
                    try:
                        for child in element.find('.//CLS_MKB'):
                            resident.append(child.text)
                    except AttributeError:
                        resident.append('NONE')
                            #</CLS_MKB>

                                #<CLS_MUSL> (раздел повторяется, => формир. вложенный список) 

                    for child in element.findall('.//USL'):
                        usl_list.append(child.text)   
                    resident.append(usl_list)

                    for child in element.findall('.//USL_F'):
                        usl_f_list.append(child.text)
                    resident.append(usl_f_list)

                    for child in element.findall('.//USL_DAT'):
                        usl_dat_list.append(child.text)
                    resident.append(usl_dat_list)

                    for child in element.findall('.//SUMV_USL'):           
                        sumv_usl_list.append(child.text)
                    resident.append(sumv_usl_list)

                    for child in element.findall('.//USL_KRAT'):         
                        usl_krat_list.append(child.text)
                    resident.append(usl_krat_list)

                    for child in element.findall('.//USL_OSN'):         
                        usl_osn_list.append(child.text)
                    resident.append(usl_osn_list)

                    for child in element.findall('.//IS_EXT_USL'):          
                        is_ext_usl_list.append(child.text)
                    resident.append(is_ext_usl_list)

                    for child in element.findall('.//USL_PRVS'):
                        usl_prvs_list.append(child.text)
                    resident.append(usl_prvs_list)

                    for child in element.findall('.//USL_KODVR'):
                        usl_kodvr_list.append(child.text)
                    resident.append(usl_kodvr_list)

                    for child in element.findall('.//USL_VR_SNILS'):
                        usl_vr_snils_list.append(child.text)
                    resident.append(usl_vr_snils_list)

                    for child in element.findall('.//USL_PR_NEP'):
                        usl_pr_nep_list.append(child.text)
                    resident.append(usl_pr_nep_list)

                    for child in element.findall('.//COMMENTU'):
                        commentu_list.append(child.text)
                    resident.append(commentu_list)
                                #</CLS_MUSL>
                            #</SL_COM>

                            #<SLS_DD>
                    for child in element.find('.//SL_DD'):
                        resident.append(child.text)
                            #</SLS_DD>
                        #</SLUCH>
                        

                    csvwriter.writerow(resident) #запись данных (ZAP) в .csv
                
                    count_zap = count_zap + 1

                list_data.close

            except FileNotFoundError:
                messagebox.showinfo("Ошибка!", "Файл не распознан!") ###ЗАМЕНИТЬ НА ОКНО С ОШИБКОЙ!!!!!!
            list_data.close()
            self.main_settings()

        def open_xls_file(self, event=None):
            """открывает указанный xls файл и создает датафрейм"""
            xlsfile = filedialog.askopenfilename(initialdir = "/",title = "Выберите файл-таблицу", filetypes = (("","*.xls"), ("all files","*.*")))
            self.file_to_pandas(xlsfile)

        def file_to_pandas(self, file):
            #try:
            df_new=pd.read_excel(file, header=0, sep=',', encoding="Windows-1251")
            #except Exception:
                #messagebox.showerror(title="Ошибка!", message="Возможно, файл поврежден. Попробуйте выбрать другой файл")

            df_new.to_csv(r'temp.csv', index = False, header=True, encoding="Windows-1251")
            self.main_settings()

        def main_settings(self, event=None):
            """Создает виджеты настроек анализа в верхнем фрейфе главного окна"""

            #специалист, дата осмотра, диагноз, номера услуг, данные пациента
            ###суммировать по заданным признакам
            ###списки с нарастающим итогом           
            global df
            df = pd.read_csv('temp.csv', header=0, sep=',', encoding="Windows-1251")
            #try:#except Exception:
            #    messagebox.showinfo('Ошибка распознавания temp.csv','Не удалось сформировать таблицу. Попробуйте сначала открыть XML-документ')
            
            fontStyle=tkFont.Font(family="Lucida Grande", size=11)
            
            #значения radioB и CheckB отправляются в get_settings()
            global r_var
            r_var = IntVar()
            r_var.set(0)
            self.label_1 = Label(self.f_top, bg="WhiteSmoke", fg="black", font=fontStyle, text='Суммировать по признаку')
            self.label_1.grid(row=0, column=0, sticky=W, padx=10, pady=10)
            self.r1 = ttk.Radiobutton(
                                    self.f_top, 
                                    text="Найти количество:", 
                                    variable=r_var, 
                                    value=0
                                    )
            self.r1.grid(row=1, column=0, sticky=W, padx=10, pady=10)

            self.combox01 = ttk.Combobox(self.f_top)
            title_list = tuple(df.columns)
            self.combox01['values'] = title_list
            self.combox01.bind("<<ComboboxSelected>>", self.callbackFunc)
            self.combox01.grid(row=1, column=1, sticky=W, padx=10, pady=10)


            self.lab_2 = ttk.Label(self.f_top, text = "в группе: ")
            self.lab_2.grid(row=2, column=0)
            self.combox02 = ttk.Combobox(self.f_top)
            self.combox02.grid(row=2, column=1, sticky=W, padx=10, pady=10)


            self.label_2 = Label(self.f_top, bg="WhiteSmoke", fg="black", font=fontStyle, text='Сортировка')
            self.label_2.grid(row=0, column=2, sticky=W, padx=10, pady=10)

            self.check_sort = ttk.Radiobutton(self.f_top, text="Сортировать по: ", variable=r_var, value=2)
            self.check_sort.grid(row=1, column=2, sticky=W, padx=10, pady=10) 

            self.label_3 = Label(self.f_top, bg="WhiteSmoke", fg="black", font=fontStyle, text='Нарастающий итог')
            self.label_3.grid(row=0, column=3, sticky=W, padx=10, pady=10)

            self.r2 = ttk.Radiobutton(self.f_top, text="Нарастающий итог", variable=r_var, value=1)
            self.r2.grid(row=1, column=3, sticky=W, padx=10, pady=10)

            global date_sort
            date_sort = IntVar()
            date_sort.set(1)
            self.check_date_sort = ttk.Checkbutton(self.f_top, text='С сортировкой по дате', variable=date_sort, onvalue=1, offvalue=0)

            self.label_date = Label(self.f_top, fg="black", bg="WhiteSmoke", text="Выберите столбец для сортировки:")
            self.combo_date_for_cumsum = ttk.Combobox(self.f_top)
            self.combo_date_for_cumsum['values'] = title_list
            self.label_summ = Label(self.f_top, fg="black", bg="WhiteSmoke", text="Выберите столбец сумм:")
            self.combo_summ_for_cumsum = ttk.Combobox(self.f_top)
            self.combo_summ_for_cumsum['values'] = title_list

            self.check_date_sort.grid(row=2, column=3, sticky=W, padx=10, pady=10)
            self.label_date.grid(row=3, column=3, sticky=W, padx=10, pady=10)
            self.combo_date_for_cumsum.grid(row=3, column=4, sticky=W, padx=10, pady=10)
            self.label_summ.grid(row=4, column=3, sticky=W, padx=10, pady=10)
            self.combo_summ_for_cumsum.grid(row=4, column=4, sticky=W, padx=10, pady=10)           

            

            
            self.combox_sort = ttk.Combobox(self.f_top)
            self.combox_sort['values'] = title_list
            self.combox_sort.grid(row=2, column=2, sticky=W, padx=10, pady=10)


            #-----------------------------------------------------------------------------------------
            self.img_ok = PhotoImage(file='Ok.png')
            self.button_show = ttk.Button(
                                          self.f_top, 
                                          #text='Применить', 
                                          image=self.img_ok,
                                          command=self.get_settings
                                          )
            self.button_show.grid(row=5, column=1, padx=10, pady=10)

            #self.button_compare = ttk.Button(self.f_top, text="Сравнить с паспортом педиатрического учаска", command=self.compare)
            #self.button_compare.grid(row=6, column=4, padx=10, pady=10)

            self.show_results(df)


        def callbackFunc(self, event):
            """заполняет значения выпадающего списка уникальными значениями в выбранном столбце"""
            group_var = self.combox01.get()
            #print('уникальные значения в столбце: ',df[group_var].unique())
            self.combox02['values'] = tuple(df[group_var].unique())



        def get_settings(self, event=None):
            """принимает настройки из main, обрабатывает .csv документ как dataframe с помощью pandas"""
            #принимает значения из кнопок и списков
            
            df = pd.read_csv('temp.csv', header=0, sep=',', encoding="Windows-1251")

            if r_var.get() == 2:                                                            #print('сортировать по выбранному признаку')
                self.sort_val(df)
            elif r_var.get() == 0:                                                              #print('Суммировать по заданным признакам')
                self.summa(df)
            elif r_var.get() == 1:                                                           #print('Нарастающий итог по сумме')
                self.cumulativ_sum(df)



        def cumulativ_sum(self, df):
            """Расчет накопительного итога (накопительная сумма) - сортировка по дате, затем заполняется столбец накопит суммы"""
            final_df=df.copy()

            if date_sort.get() == 1: #с сортировкой по дате
                date = self.combo_date_for_cumsum.get()
            sum = self.combo_summ_for_cumsum.get()

            final_df= df.sort_values(by=date)
            final_df['cumsum'] = final_df[sum].cumsum()
            #сделать выбор столбца с датой и столбца с суммой!!!!!!!!!! и get()
            self.clear_tv()
            self.show_results(final_df)

        def summa(self, df):
            """кол-во пациентов по заданному параметру в заданной группе"""
            final_df=df.copy()
            param_var = self.combox01.get()
            group_var = self.combox02.get()

            final_df = df[df[param_var].isin([group_var])]


            #counterFunc = final_df.apply(lambda x: True if str(x[param_var]) == 'group_var' else False , axis=1)


            Str = 'Найдено пациентов: ' + str(len(final_df.index))
            messagebox.showinfo("Поиск завершен", Str)
            self.clear_tv()
            self.show_results(final_df)

        def sort_val(self, df):
            #final_df=df.copy()
            sort = self.combox_sort.get()
            df = df.sort_values(by=[sort], ascending=True)
            self.clear_tv()
            self.show_results(df)

        def compare(self, event=None):
            """сравнивает df из csv файла с паспортом педиатрич участка .xls"""
            df1=pd.read_csv('temp.csv', header=0, sep=',', encoding="Windows-1251")   #df из csv
            passport=filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("","*.xls"), ("all files","*.*")))
            df2=pd.read_excel(passport, header=5, sep=',', encoding="Windows-1251")   #df из xls ППУ

            if df1.columns[1] == 'PR_NOV':
                #####если сравнивается реестр:
                a = df1['FAM'].tolist() 
                b = df1['IM'].tolist()
                c = df1['OT'].tolist()
                list_FIO=[]
                #составить список ФИО из отдельных столбцов реестра
                for i in range(0, len(a)):
                    list_FIO.append(a[i] + '  ' + b[i] + '  ' + c[i])

                df3=df2.loc[df2[df2.columns[1]].isin(list_FIO)] #найти совпадения в списке ФИО и ППУ

                fam_from_ppu = []
                fam_from_ppu = df3[df3.columns[1]].tolist()

                Sp = []
                for elem in fam_from_ppu:
                    Sp.append(elem.split('  '))   #разделили список ФИО из ППУ на отдельные списки (ВЫДЕЛЯЕМ ОТДЕЛЬНО ФАМИЛИИ)

                Sp_fam = []
                for i in range(len(Sp)):
                    Sp_fam.append(Sp[i][0])       #получили список фамилий

                df4 = df1.loc[df1['FAM'].isin(Sp_fam)] #df из строк реестра, которые совпадают с ППУ
                Str='Найдено совпадений: ' + str(len(Sp_fam))
                messagebox.showinfo('Поиск завершен', Str)

                self.clear_tv()
                self.show_results(df4)
            else:
                list_FIO=[]
                #заполнить список ФИО!
                #

                df3=df2.loc[df2[df2.columns[1]].isin(list_FIO)]#совпадения в списке ФИО и ППУ
                fam_from_ppu = []
                fam_from_ppu = df3[df3.columns[1]].tolist()
                Sp_fam = []#нужно, если в исходном файле фио в три столбца 
                for i in range(len(Sp)):
                    Sp_fam.append(Sp[i][0])       #получили список фамилий

                df4 = df1.loc[df1['FAM'].isin(Sp_fam)] #df из строк реестра, которые совпадают с ППУ
                Str='Найдено совпадений: ' + str(len(Sp_fam))
                messagebox.showinfo('Поиск завершен', Str)
                self.show_results(df4)




        def show_results(self, final_df):
            
            data=()
            data = final_df.values
            rows=data  #rows=tuple(df.valuse)
            
            headings=tuple(final_df.columns) #кортеж заголовков - если нужны изменения, заменить на список

            #заполнение виджета treeview массивом данных из датафрейма
            self.treeview["columns"]=headings
            self.treeview["displaycolumns"]=headings
            for head in headings:
                self.treeview.heading(head, text=head)
                self.treeview.column(head)
            for row in rows:
                self.treeview.insert('', END, values=tuple(row))  # если не использовать кортеж, то все плохо

            self.treeview.pack_forget()
            self.treeview.pack(side=RIGHT)


        def save_xml(self, event=None):#, output_reestr):
            """сохранить .xml"""
            save_as = filedialog.asksaveasfilename(defaultextension=".xml")
            try:
                tree.write(save_as)
            except Exception:
                messagebox.showerror('Ошибка', 'Запись возможна только для XML-документов')

        def append_to_reestr(self, event=None):
            """Запись в реестр"""
            reestr = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("","*.xml"), ("all files","*.*")))
            reestr_tree = ET.parse(reestr)
            reestr_root = reestr_tree.getroot()   #получение корневого дерева тегов
            try:
                reestr_root.append(root)
                answer = messagebox.askyesno(title='Внимание!', message='Вы действительно хотите записать данные в реестр?')
                if answer == True:
                    reestr_tree.write(reestr)
                    messagebox.showinfo('Запись в реестр', 'Запись выполнена')
            except Exception:
                messagebox.showerror('Ошибка', 'Запись в реестр возможна только для XML-документов')

        def day_night(self, event=None):
            """Меняет тему оформления день/ночь"""
            self.f_top['bg']="#A1A1A1"
            self.label_date['bg']="#A1A1A1"
            self.label_summ['bg']="#A1A1A1"
            self.label_2['bg']="#A1A1A1"
            self.label_3['bg']="#A1A1A1"
            self.label_1['bg']="#A1A1A1"
            #self.lab_2['bg']="#A1A1A1"
            #self.r1['bg']="#A1A1A1"


        def clear_workspace(self, event=None):
            """Закрыть рабочий файл, очистить все"""
            answer = messagebox.askyesno(title='Завершение работы с текущим документом', message='Очистить рабочую область?')
            if answer == True:
                self.clear_tv()
                open('temp.csv', 'w').close()

        def clear_tv(self, event=None):
            """очистить treeview"""
            for i in self.treeview.get_children():
                self.treeview.delete(i)
        def tv_refresh(self, event=None):
            """очистить tv, вывести df из temp.csv"""
            self.clear_tv()
            self.main_settings()

        def close(self, event=None):
            self.main.destroy()



        def find_window(self, event=None):
            """поиск элемента в таблице"""
            self.top_find_wind = Toplevel()
            self.top_find_wind.geometry('300x200')
            self.top_find_wind.iconbitmap("icon.ico")
            self.top_find_wind.title('Введите запрос для поиска')
            
            self.lab = Label(self.top_find_wind, text='Искать:')
            self.lab.pack()

            self.entr = Entry(self.top_find_wind)
            self.entr.pack()

            self.but = Button(self.top_find_wind, text='Поиск', command=self.find)
            self.but.pack()

        def find(self, event=None):
            find_request = self.entr.get()
            




        def create_admin_win(self, event=None):
            self.ad = Toplevel()
            self.ad.geometry('300x300')
            self.ad.iconbitmap("icon.ico")
            self.ad.title('Настройки администратора')
            self.ad['bg'] = 'lightgrey'#'#1A2C34'
            self.l1 = Label(self.ad, text='Настройки аутентификации')
            self.b1 = ttk.Button(self.ad, text='Добавить пользователя',
                             command=self.add_user)

            self.b2 = ttk.Button(self.ad, text='Удалить пользователя',
                             command=self.del_user)
            self.b3 = ttk.Button(self.ad, text='Редактировать пользователя',
                             command=self.edit_user)
            self.l1.grid(row=1, column=0, sticky=W+E)
            self.b1.grid(row=3, column=0, sticky=W+E)
            self.b2.grid(row=4, column=0, sticky=W+E)
            self.b3.grid(row=5, column=0, sticky=W+E)

        def add_user(self, event=None):
            print('add user')
            self.add_us = Toplevel()
            self.add_us.geometry('300x300')
            self.add_us.iconbitmap("icon.ico")
            self.add_us.title('Настройки администратора')
            self.add_us['bg'] = 'lightgrey'#'#1A2C34'
            self.l1 = Label(self.add_us, text='Логин нового пользователя')
            self.l2 = Label(self.add_us, text='Пароль нового пользователя')
            self.elog = Entry(self.add_us, font = "16", width=20)
            self.epass = Entry(self.add_us, font = "16", width=20)
            self.b1 = ttk.Button(self.add_us, text='Добавить пользователя',
                                 command=self.add_user_to_list)
            self.l1.grid()
            self.elog.grid()
            self.l2.grid()
            self.epass.grid()
            self.b1.grid()

        def add_user_to_list(self, event=None):
            l = self.elog.get()
            p = self.epass.get()
            s = l + ' ' + p
            try:
                with open('.doc\\passwords.txt', 'a') as f:   
                    f.write('\n')
                    f.write(s)        #############добавить проверку на повторяющиеся логины    
                messagebox.showinfo("Добавление нового пользователя", "Пользователь добавлен!")
            except OSError:
                messagebox.showinfo("Добавление нового пользователя", "Ошибка! Проверьте целостность файла")
            except Exception:
                messagebox.showinfo("Добавление нового пользователя", "Ошибка! Проверьте целостность файла")

        def del_user(self, event=None):
            print('delete user')
            self.del_us = Toplevel()
            self.del_us.geometry('300x300')
            self.del_us.iconbitmap("icon.ico")
            self.del_us.title('Настройки администратора')
            self.del_us['bg'] = 'lightgrey'#'#1A2C34'
            self.l1 = Label(self.del_us, text='Выберите пользователя, которого нужно удалить:')
            self.l1.grid()
            self.lbox = Listbox(width=15)
            self.lbox.grid()
            #for i in ():
            #    lbox.insert(END, i)

        def edit_user(self, event=None):
            """изменить существующие логин-пароль"""
            #print('edit user')

        def create_sett_win(self, event=None):
            self.set = Toplevel()
            self.set.geometry('300x300')
            self.set.iconbitmap("icon.ico")
            self.set.title('Настройки отображения')
            self.set['bg'] = 'lightgrey'#'#1A2C34'
            self.l1 = Label(self.set, text="Цвет фона")
            self.l1.grid()
            self.l2 = Label(self.set, text="Размер шрифта")
            self.l2.grid()

        def show_info(self, event=None):
            """Отображение общих данных об открытом файле"""
            df1=pd.read_csv('temp.csv', header=0, sep=',', encoding="Windows-1251")
            kol_vo=str(len(df.index))
            

            self.info=Toplevel()
            self.info.geometry('300x300')
            self.info.iconbitmap("icon.ico")
            self.info.title('Данные о файле')
            #self.info['bg'] = 'lightgrey'#'#1A2C34'

            self.lab1 = ttk.Label(self.info, text="Количнство пациентов: ")
            self.lab01 = ttk.Label(self.info, text=kol_vo)
            self.lab1.grid(row=0, column=0)
            self.lab01.grid(row=0, column=1)

            self.lab_empty = ttk.Label(self.info, text="   ")
            self.lab_empty.grid(row=1, column=0)
            
            self.lab2 = ttk.Label(self.info, text="Информация по столбцу:")
            self.lab02 = ttk.Label(self.info)

            self.combo02 = ttk.Combobox(self.info)
            self.combo02['values'] = tuple(df.columns)
            self.combo02.bind("<<ComboboxSelected>>", self.stat_info)
            self.combo02.grid(row=3,column=0)
            
            self.lab2.grid(row=2, column=0)
            self.lab02.grid(row=2, column=2, rowspan=2)


            self.lab4 = ttk.Label(self.info, text="")
            self.lab5 = ttk.Label(self.info, text="")
            self.lab6 = ttk.Label(self.info, text="")
            self.lab7 = ttk.Label(self.info, text="")

        def stat_info(self, event):
            """выдает стат информацию о заданном столбце"""
            st = self.combo02.get()
            df_st = df[st]
            descr=df_st.describe()
            self.lab02['text'] = descr

        def show_help(self, event=None):#ЗАПОЛНИТЬ
            """Окно справки"""
            self.help_win=Toplevel()
            self.help_win.iconbitmap("icon.ico")
            self.help_win.title('Справка')
            self.help_win['bg'] = 'lightgrey'#'#1A2C34'

            self.help_tree=Frame(self.help_win)
            self.help_tree.pack(padx=10, pady=10, side = 'left')

            self.treeview = ttk.Treeview(self.help_tree, selectmode='browse')

            self.treeview.insert('','0','statistica', text='Обработка данных')
            self.treeview.insert('statistica', '1', 'parametr', text='Какие параметры выбрать?')

            self.treeview.insert('','0','saving', text='Сохранение данных')
            self.treeview.insert('saving', '1', 'save_table', text='Как сохранить таблицу?')
            self.treeview.insert('saving', '1', 'save_to_reestr', text='Как сохранить в реестр?')

            self.treeview.insert('','0','start', text='Начало работы')
            self.treeview.insert('start', '1', 'to_do', text='Загрузка файла')

            self.treeview.insert('', '0', 'about', text='О программе')
            self.treeview.insert('about', '1', 'how_it_works', text='Элементы окна')

            self.treeview.pack(side=LEFT)

            
            #self.textview = ttk.Treeview(self.help_tree, selectmode='browse')
            #self.textview.pack(side=LEFT)

           # self.help_img = PhotoImage(file = 'help.png')
            #self.help_label = Label(self.help_win, image=self.help_img).pack(side=LEFT)


            self.treeview.bind('<<TreeviewSelect>>', self.on_select)


        def on_select(self, event):
            item = self.treeview.selection()[0]
            #print(item)
            # Выбираем функцию из словаря, если элемента в словаре нет - выполняется действие по-умолчанию
            t = StringVar()
            self.help_entry = Entry(self.help_win, state = DISABLED, bg='white', textvariable=t)
            self.help_entry.pack(side=LEFT)

            funcs_for_items = {
                'about': lambda: t.set('ololo'),
                #'how_it_works': lambda:
                #'parametr': lambda: self.help_entry
                'to_do': lambda: print('Function 2')
                                }

            func = funcs_for_items.get(item, lambda: print('Default action'))
            func()



        def show_about(self, event=None):
            """Справка - О программе"""
            messagebox.showinfo('О программе', 'Разработано для ОГАУЗ Детская больница №1 г. Томска. Автор - Марданшина И.Н.')

        def create_menu(self):

            self.menubar = Menu(self.main, bg='lightgrey', fg='black') # формируется меню 

            self.file_menu = Menu(self.menubar, tearoff=0, bg="lightgrey", fg="black")
            self.file_menu.add_command(label="Открыть счет-реестр", command=self.open_xml_file)
            self.file_menu.add_command(label="Открыть табличную выгрузку", command=self.open_xls_file)
            self.file_menu.add_command(label="Сохранить как...", command=self.save_xml)
            self.file_menu.add_command(label="Записать в реестр", command=self.append_to_reestr)
            self.file_menu.add_separator()
            self.file_menu.add_command(label="Выход", command=self.close)

            self.edit_menu = Menu(self.menubar, tearoff=0, bg="lightgrey", fg="black")
            self.edit_menu.add_command(label="Общие данные о файле", command=self.show_info)

            self.sett_menu = Menu(self.menubar, tearoff=0, bg="lightgrey", fg="black")
            self.sett_menu.add_command(label="Настройки отображения", command=self.create_sett_win)
            if login == 'admin':
                self.sett_menu.add_command(label="Настройки пользователя (только для администратора)", 
                                            command=self.create_admin_win) #сделать недоступным для всех кроме: логин = админ!!!

            self.help_menu = Menu(self.menubar, tearoff=0, bg="lightgrey", fg="black")
            self.help_menu.add_command(label="Справка", command=self.show_help)
            self.help_menu.add_command(label="О программе", command=self.show_about)
            

            self.menubar.add_cascade(label="Файл", menu=self.file_menu)
            self.menubar.add_cascade(label="Инструменты", menu=self.edit_menu)
            self.menubar.add_cascade(label="Настройки", menu=self.sett_menu)
            self.menubar.add_cascade(label="Помощь", menu=self.help_menu)
            self.main.config(menu = self.menubar) # меню добавляется к окну

        def liveMain(self):
            self.main.mainloop()



dlg=WmDialog()
dlg.liveWnd()

#print('destroy 1-st window')

main=wmain()
main.liveMain()


