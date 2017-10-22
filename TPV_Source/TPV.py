from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as mBox
from pathlib import Path
from configobj import ConfigObj
import os.path
from datetime import datetime, date
import webbrowser
import openpyxl as op
import logging
from docx import Document


class TPV_Main():

    def __init__(self):

        # Create the Logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)

        # Create the Handler for logging data to a file
        logger_handler = logging.FileHandler('TPV_log.log')
        logger_handler.setLevel(logging.INFO)

        # Create a Formatter for formatting the log messages
        logger_formatter = logging.Formatter('%(name)s - %(levelname)s - %(message)s')

        # Add the Formatter to the Handler
        logger_handler.setFormatter(logger_formatter)

        # Add the Handler to the Logger
        self.logger.addHandler(logger_handler)


        self.TPV_Body = ttk.LabelFrame(win, text='TPV Skjema')
        self.TPV_Body.pack(expand=1)

        menuBar = Menu(win)
        win.config(menu=menuBar)

        fileMenu = Menu(menuBar, tearoff=0)
        helpMenu = Menu(menuBar, tearoff=0)

        helpMenu.add_command(label='Prosedyre', command=self.procedure)
        helpMenu.add_command(label='Hjelp', command=self.op_wiki)
        menuBar.add_cascade(label='Info', menu=helpMenu)

        fileMenu.add_command(label='Lagre Utført vedlikehold', command=self.maintanance)
        fileMenu.add_command(label='Åpne excel fil', command=self.op_saved)
        menuBar.add_cascade(label='Alternativer', menu=fileMenu)


        self.config()
        self.main()

        self.logger.info('Completed configuring logger(), and class is initsialized!')

    def config(self):

        if os.path.isfile('config.ini') == True:
            config = ConfigObj('config.ini')
            size = config['Filbehandling']['2']

            win.geometry(size)

            self.logger.info('Config file exists, using existing one!')

        else:
            config = ConfigObj(encoding='utf8', default_encoding='utf8')
            config.filename = config_name

            config['Vedlikeholdspunkt'] = {}
            config['Vedlikeholdspunkt']['1'] = 'Put your info here'
            config['Vedlikeholdspunkt']['2'] = 'Put your info here'
            config['Vedlikeholdspunkt']['3'] = 'Put your info here'
            config['Vedlikeholdspunkt']['4'] = 'Put your info here'
            config['Vedlikeholdspunkt']['5'] = 'Put your info here'

            config['Handling'] = {}
            config['Handling']['1'] = 'Put your info here'
            config['Handling']['2'] = 'Put your info here'
            config['Handling']['3'] = 'Put your info here'
            config['Handling']['4'] = 'Put your info here'
            config['Handling']['5'] = 'Put your info here'

            config['Oljetype'] = {}
            config['Oljetype']['1'] = 'Put your info here'
            config['Oljetype']['2'] = 'Put your info here'
            config['Oljetype']['3'] = 'Put your info here'
            config['Oljetype']['4'] = 'Put your info here'
            config['Oljetype']['5'] = 'Put your info here'

            config['Hyppighet'] = {}
            config['Hyppighet']['1'] = 'Put your info here'
            config['Hyppighet']['2'] = 'Put your info here'
            config['Hyppighet']['3'] = 'Put your info here'
            config['Hyppighet']['4'] = 'Put your info here'
            config['Hyppighet']['5'] = 'Put your info here'

            config['Diversje'] = {}
            config['Diversje']['1'] = 'Du kan skrive ekstra info her:'
            config['Diversje']['2'] = 'Annet:'
            config['Diversje']['3'] = ''
            config['Diversje']['4'] = '20'
            

            config['Filbehandling'] = {}
            config['Filbehandling']['1'] = ''
            config['Filbehandling']['2'] = '850x450'
            config['Filbehandling']['3'] = ''
            config['Filbehandling']['4'] = ''

            date = datetime.today()
            year = date.strftime('%Y')
            config['Diversje']['3'] = year

            config.write()

            size = config['Filbehandling']['2']
            win.geometry(size)

            self.logger.info('Config file does not exist, creating a new one!')



    def main(self):
        
        """Main body of the gui application"""

        #Accessing the config parser and getting the number of keys
        config1 = ConfigObj('config.ini')
        keys = config1.keys()

        first_key = keys[0]

        #Collecting current date and weekday in full name
        date = datetime.today()
        today = date.strftime('%A')

        #Collecting which number in the month the current day is
        today_number = date.strftime('%d')
        today_number = int(today_number)

        #Collecting which number in the year the current day is
        day_of_year = date.strftime('%j')
        day_of_year = int(day_of_year)

        #Empty list assigned for holding info on entries in the first key of the config file
        values = []

        #Establish the number of entries in the config file
        for i in config1[first_key]:
            values.append(i)

        #Assigning the length of list values to a global variable
        global length
        length = len(values)
        
        testVar = self.day_check()

        row_ved = 1
        row_han = 1
        row_olj = 1
        row_hyp = 1

        #Generating labels on the fly based on how many entries it is in the config file
        for value in config1['Vedlikeholdspunkt'].values():
            label = ttk.Label(self.TPV_Body, text=value, font=FONT1)
            label.grid(row=row_ved, column=1, sticky=W, padx=15)
            row_ved += 1


        for value in config1['Handling'].values():
            label = ttk.Label(self.TPV_Body, text=value, font=FONT1)
            label.grid(row=row_han, column=2, sticky=W, padx=15)
            row_han += 1

        for value in config1['Oljetype'].values():
            label = ttk.Label(self.TPV_Body, text=value, font=FONT1)
            label.grid(row=row_olj, column=3, sticky=W, padx=15)
            row_olj += 1


        for value in config1['Hyppighet'].values():

            lowCas = value.lower()

            if lowCas == 'daglig':
                label = ttk.Label(self.TPV_Body, text=value, font=FONT1, background='green')
                label.grid(row=row_hyp, column=4, sticky=W, padx=15)
                row_hyp += 1

            elif lowCas == 'ukentlig' and today == 'Friday':

                label = ttk.Label(self.TPV_Body, text=value, font=FONT1, background='yellow')
                label.grid(row=row_hyp, column=4, sticky=W, padx=15)
                row_hyp += 1

            elif lowCas == 'maanedlig' and today_number == testVar:

                label = ttk.Label(self.TPV_Body, text=value, font=FONT1, background='orange')
                label.grid(row=row_hyp, column=4, sticky=W, padx=15)
                row_hyp += 1

            elif lowCas == 'halvaar' and day_of_year == 183:

                label = ttk.Label(self.TPV_Body, text=value, font=FONT1, background='red')
                label.grid(row=row_hyp, column=4, sticky=W, padx=15)
                row_hyp += 1

            else:
                label = ttk.Label(self.TPV_Body, text=value, font=FONT1)
                label.grid(row=row_hyp, column=4, sticky=W, padx=15)
                row_hyp += 1



        if length == 5:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()

            self.list5 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                          self.checkVar5]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=N)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 6:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()

            self.list6 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                          self.checkVar5, self.checkVar6]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=N)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 7:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()

            self.list7 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                          self.checkVar5, self.checkVar6, self.checkVar7]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=N)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 8:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()

            self.list8 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                          self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=N)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 9:
            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()

            self.list9 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                          self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                          self.checkVar9]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 10:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()
            self.checkVar10 = IntVar()

            self.list10 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                           self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                           self.checkVar9, self.checkVar10]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)
            check10 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar10)
            check10.grid(row=10, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 11:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()
            self.checkVar10 = IntVar()
            self.checkVar11 = IntVar()

            self.list11 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                           self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                           self.checkVar9, self.checkVar10, self.checkVar11]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)
            check10 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar10)
            check10.grid(row=10, column=0, sticky=W)
            check11 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar11)
            check11.grid(row=11, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 12:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()
            self.checkVar10 = IntVar()
            self.checkVar11 = IntVar()
            self.checkVar12 = IntVar()

            self.list12 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                           self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                           self.checkVar9, self.checkVar10, self.checkVar11, self.checkVar12]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)
            check10 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar10)
            check10.grid(row=10, column=0, sticky=W)
            check11 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar11)
            check11.grid(row=11, column=0, sticky=W)
            check12 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar12)
            check12.grid(row=12, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 13:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()
            self.checkVar10 = IntVar()
            self.checkVar11 = IntVar()
            self.checkVar12 = IntVar()
            self.checkVar13 = IntVar()

            self.list13 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                           self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                           self.checkVar9, self.checkVar10, self.checkVar11, self.checkVar12,
                           self.checkVar13]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)
            check10 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar10)
            check10.grid(row=10, column=0, sticky=W)
            check11 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar11)
            check11.grid(row=11, column=0, sticky=W)
            check12 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar12)
            check12.grid(row=12, column=0, sticky=W)
            check13 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar13)
            check13.grid(row=13, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 14:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()
            self.checkVar10 = IntVar()
            self.checkVar11 = IntVar()
            self.checkVar12 = IntVar()
            self.checkVar13 = IntVar()
            self.checkVar14 = IntVar()

            self.list14 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                           self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                           self.checkVar9, self.checkVar10, self.checkVar11, self.checkVar12,
                           self.checkVar13, self.checkVar14]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)
            check10 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar10)
            check10.grid(row=10, column=0, sticky=W)
            check11 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar11)
            check11.grid(row=11, column=0, sticky=W)
            check12 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar12)
            check12.grid(row=12, column=0, sticky=W)
            check13 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar13)
            check13.grid(row=13, column=0, sticky=W)
            check14 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar14)
            check14.grid(row=14, column=0, sticky=W)

            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)

            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        elif length == 15:

            self.checkVar1 = IntVar()
            self.checkVar2 = IntVar()
            self.checkVar3 = IntVar()
            self.checkVar4 = IntVar()
            self.checkVar5 = IntVar()
            self.checkVar6 = IntVar()
            self.checkVar7 = IntVar()
            self.checkVar8 = IntVar()
            self.checkVar9 = IntVar()
            self.checkVar10 = IntVar()
            self.checkVar11 = IntVar()
            self.checkVar12 = IntVar()
            self.checkVar13 = IntVar()
            self.checkVar14 = IntVar()
            self.checkVar15 = IntVar()

            self.list15 = [self.checkVar1, self.checkVar2, self.checkVar3, self.checkVar4,
                           self.checkVar5, self.checkVar6, self.checkVar7, self.checkVar8,
                           self.checkVar9, self.checkVar10, self.checkVar11, self.checkVar12,
                           self.checkVar13, self.checkVar14, self.checkVar15]

            check1 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar1)
            check1.grid(row=1, column=0, sticky=W)
            check2 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar2)
            check2.grid(row=2, column=0, sticky=W)
            check3 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar3)
            check3.grid(row=3, column=0, sticky=W)
            check4 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar4)
            check4.grid(row=4, column=0, sticky=W)
            check5 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar5)
            check5.grid(row=5, column=0, sticky=W)
            check6 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar6)
            check6.grid(row=6, column=0, sticky=W)
            check7 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar7)
            check7.grid(row=7, column=0, sticky=W)
            check8 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar8)
            check8.grid(row=8, column=0, sticky=W)
            check9 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar9)
            check9.grid(row=9, column=0, sticky=W)
            check10 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar10)
            check10.grid(row=10, column=0, sticky=W)
            check11 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar11)
            check11.grid(row=11, column=0, sticky=W)
            check12 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar12)
            check12.grid(row=12, column=0, sticky=W)
            check13 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar13)
            check13.grid(row=13, column=0, sticky=W)
            check14 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar14)
            check14.grid(row=14, column=0, sticky=W)
            check15 = ttk.Checkbutton(self.TPV_Body, variable=self.checkVar15)
            check15.grid(row=15, column=0, sticky=W)


            lbl1 = Label(self.TPV_Body, text='Vedlikeholdspunkt:', font=FONT2)
            lbl1.grid(row=0, column=1, sticky=N, pady=5)
            lbl2 = Label(self.TPV_Body, text='Handling:', font=FONT2)
            lbl2.grid(row=0, column=2, sticky=N, pady=5)
            lbl3 = Label(self.TPV_Body, text='Oljetype:', font=FONT2)
            lbl3.grid(row=0, column=3, sticky=W, pady=5)
            lbl4 = Label(self.TPV_Body, text='Hyppighet:', font=FONT2)
            lbl4.grid(row=0, column=4, sticky=N, pady=5)

            header = config1['Diversje']['1']
            lbl16 = Label(self.TPV_Body, text=header)
            lbl16.grid(row=16, column=0, columnspan=2,  pady=3)

            self.txt = Text(self.TPV_Body, height=3, width=25)
            self.txt.grid(row=17, column=0, columnspan=2)


            button = ttk.Button(self.TPV_Body, text='Lagre', command=self.save)
            button.grid(row=21, column=0, columnspan=2, sticky=W, pady=15)

        else:
            self.logger.critical('Initzialisation of main body failed!')

        self.logger.info('Main body successfully initzialised!')


    def save(self):

        values1 = []
        values2 = []

        if length == 5:
            for i in self.list5:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 6:
            for i in self.list6:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 7:
            for i in self.list7:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 8:
            for i in self.list8:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 9:
            for i in self.list9:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 10:
            for i in self.list10:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 11:
            for i in self.list11:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 12:
            for i in self.list12:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 13:
            for i in self.list13:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 14:
            for i in self.list14:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        elif length == 15:
            for i in self.list15:
                values1.append(i)

            for i in values1:
                x = i.get()
                values2.append(x)

        else:
            self.logger.critical("Failed initialization in the save method!")
        
        self.logger.info("Initialization of save method complete!")
        
        self.values2 = values2

        #Main code for writing out the excel file used to save the data in comes here:

        #Calling the config parser and accessing filepath if present in config file
        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')
        f_name = config['Filbehandling']['1']

        date = datetime.today()

        #Dateformatting
        month = date.strftime('%B')
        month = str(month)

        #Saving dates as strings for use in excel file
        day_as_string = date.strftime('%d')
        day_as_string = int(day_as_string) + 1
        day_as_string = str(day_as_string)

        #Saving dates as int's for use in locating rows
        day_as_int = date.strftime('%d')
        day_as_int = int(day_as_int) + 1

        #Formatting date to my liking
        today = date.strftime('%d.%m.%Y, %a')


        cell = 'A' + day_as_string

        index = []

        months = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
        'August', 'September', 'October', 'November', 'December']

        for i in config['Vedlikeholdspunkt'].values():
            index.append(i)

        index.append(config['Diversje']['2'])

        #Checking if the file exist or not
        if os.path.isfile(f_name) == False or self.year_check() == 1:

            self.logger.info('Excel file does not exist, creating new one!')

            f_new = filedialog.asksaveasfilename(title='Select File',
                                                 filetypes=(("Excel files", ".xlsx"),
                                                            ("All files", "*.*")),
                                                            defaultextension="*.*")

            #Assigning the filename a place in the config file
            config['Filbehandling']['1'] = f_new

            wb = op.Workbook()

            for i in months:
                wb.create_sheet(i)

            sh = wb.get_sheet_names()
            x = wb.get_sheet_by_name('Sheet')
            wb.remove_sheet(x)
            sh = wb.get_sheet_names()

            col = 2

            for m in sh:
                ws = wb[m]
                for i in index:
                    ws.cell(row=1, column=col, value=i)
                    col += 1
                    if col > len(index) + 1:
                        col = 2
                    else:
                        continue

            ws = wb[month]

            ws[cell] = today

            col = 2

            for i in values2:
                ws.cell(row=day_as_int, column=col, value=i)
                col += 1

            try:
                op.writer.excel.save_workbook(wb, f_new)
                mBox.showinfo('', 'Resultater har blitt lagret')
                config.write()

            except FileNotFoundError:
                self.logger.critical('Saving of excel file failed!')
                pass

        elif os.path.isfile(f_name) == True and self.year_check() == 0:

            self.logger.info('Excel file exist, writing results to it!')

            f_name = config['Filbehandling']['1']

            wb = op.load_workbook(filename=f_name)
            ws = wb[month]

            ws[cell] = today

            col = 2

            for i in values2:
                ws.cell(row=day_as_int, column=col, value=i)
                col += 1

            textbox = self.txt.get('1.0', END)
            ws.cell(row=day_as_int, column=col, value=textbox)
            
            #testpart
            hyp = config['Hyppighet']
            
            test_key = []
            test_val = []
            
            for k, v in hyp.items():
                test_key.append(k)
                test_val.append(v)
                #print(k, v)
            
            rep_dag = test_val.count('Daglig')
            rep_uke = test_val.count('Ukentlig')
            rep_maan = test_val.count('Maanedlig')
            rep_hal = test_val.count('Halvaar')
            
            print(rep_dag, rep_uke, rep_maan, rep_hal)
            #print(hyp.keys)

            op.writer.excel.save_workbook(wb, f_name)

            mBox.showinfo('', 'Resultater har blitt lagret')

        else:
            self.logger.critical('Saving to excel file has failed!')

        self.win_size()


    def year_check(self):

        """A simple method for checking if we have switched year"""

        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')
        config_year = config['Diversje']['3']

        date = datetime.today()
        current_year = date.strftime('%Y')

        if config_year != current_year:
            config['Diversje']['3'] = current_year
            self.logger.info("Happy new year! Writing new year to config and returning 1.")
            return 1
        else:
            self.logger.info("Same old year, returning 0.")
            return 0


    def op_saved(self):

        """Lets you open a preexisting excel file"""

        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')

        f_exist = filedialog.askopenfilename()

        config['Filbehandling']['1'] = f_exist
        config.write()
        

    def op_wiki(self):

        """Simply opens your default browser and directs you to the wiki"""

        webbrowser.open('https://github.com/UniQueKakarot/TPV_Skjema/wiki')
        

    def procedure(self):

        """Lets you select a word file that gets linked to in the UI"""

        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')
        f_exist = config['Filbehandling']['3']

        if os.path.isfile(f_exist) == True:
            os.system(f_exist)

        elif os.path.isfile(f_exist) == False:
            f_exist = filedialog.askopenfilename()
            os.system(f_exist)
            config['Filbehandling']['3'] = f_exist
            config.write()

        else:
            self.logging.critical('Shit hit the fan in the procedure method!')
            

    def win_size(self):

        """Resizing the UI window to the minimum needed + a 100 pixels on each side"""

        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')

        width = self.TPV_Body.winfo_reqwidth()
        height = self.TPV_Body.winfo_reqheight()

        width += 100
        height += 100

        width = str(width)
        height = str(height)

        size = width + 'x' + height

        config['Filbehandling']['2'] = size
        config.write()

        print(size)
        
    def maintanance(self):
        
        """Adding in a new window for saving extra information on maintainance done"""
        
        self.new_win = Tk()
        self.new_win.title('Skjema')
        self.new_win.geometry('450x280')
        
        date = datetime.today()
        today = date.strftime('%d.%m.%Y, %a')
        
        form_body = ttk.LabelFrame(self.new_win, text='Skjema for utført Vedlikehold')
        form_body.pack(expand=1)
        
        lbl1 = Label(form_body, text='Dato: ', font=FONT1)
        lbl1.grid(row=0, column=0, sticky=N)
        self.ent1 = ttk.Entry(form_body)
        self.ent1.grid(row=0, column=1, sticky=N)
        self.ent1.insert(0, today)
        
        lbl2 = Label(form_body, text='Hvem utførte vedlikeholdet?: ', font=FONT1)
        lbl2.grid(row=1, column=0, sticky=W, pady=10)
        self.ent2 = ttk.Entry(form_body)
        self.ent2.grid(row=1, column=1, sticky=N, pady=10)
        
        lbl3 = Label(form_body, text='Hva ble gjort:', font=FONT2)
        lbl3.grid(row=2, column=0, sticky=W)
        self.maintanance_txt = Text(form_body, height=3, width=40)
        self.maintanance_txt.grid(row=3, column=0, columnspan=2, sticky=W)
        
        button1 = ttk.Button(form_body, text='Lagre', command=self.maintanance_save)
        button1.grid(row=4, column=0, sticky=N, pady=10)
            
        button2 = ttk.Button(form_body, text='Avslutt', command=self.maintanance_quit)
        button2.grid(row=4, column=1, sticky=N, pady=10)
        
        
    def maintanance_save(self):
        
        """Saving method for the main maintainance method"""
        
        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')
        f_exist = config['Filbehandling']['4']
        
        if os.path.isfile(f_exist) == True:
            
            document = Document(f_exist)
            
            date = self.ent1.get()
            name = self.ent2.get()
            textbox = self.maintanance_txt.get('1.0', END)
            
            document.add_paragraph(date)
            document.add_paragraph(name)
            document.add_paragraph(textbox)
            document.add_paragraph('')
            
            document.save(f_exist)
            
            mBox.showinfo('', 'Resultater har blitt lagret')
            
        elif os.path.isfile(f_exist) == False:
            
            f_new = filedialog.asksaveasfilename(title='Select File',
                                                 filetypes=(("Word files", ".docx"),
                                                            ("All files", "*.*")),
                                                 defaultextension="*.*")
            
            document = Document()
            
            date = self.ent1.get()
            name = self.ent2.get()
            textbox = self.maintanance_txt.get('1.0', END)
            
            document.add_paragraph(date)
            document.add_paragraph(name)
            document.add_paragraph(textbox)
            document.add_paragraph('')
            
            document.save(f_new)
            
            config['Filbehandling']['4'] = f_new
            config.write()
            
            mBox.showinfo('', 'Resultater har blitt lagret')
            
        else:
            self.logger.critical("Something failed in the maintanance_save method!")
        
        
    def maintanance_quit(self):
        
        """Simply just kills the extra save window"""
        
        self.new_win.destroy()
        
        
    def day_check(self):

        """Checks to see if the 20th day of the month falls on a weekend"""
        
        config = ConfigObj('config.ini', encoding='utf8', default_encoding='utf8')
        saved_day = config['Diversje']['4']
        
        date1 = datetime.today()
        
        year = date1.strftime('%Y')
        year = int(year)
        month = date1.strftime('%m')
        month = int(month)
        day = 20
        
        check_day = date(year, month, day).isoweekday()

        if saved_day == '20' and check_day == 6:
            day = day - 1
            config['Diversje']['4'] = day
            config.write()
            self.logger.info("20th falls on a saturday, switching to friday!")
            return day
        
        elif saved_day == '20' and check_day == 7:
            day = day + 1
            config['Diversje']['4'] = day
            config.write()
            self.logger.info("20th falls on a sunday, switching to monday!")
            return day
        
        else:
            self.logger.info("20th is a weekday, so returns 20.")
            return 20
        



        


        
        



win = Tk()
win.title("TPV Skjema")
win.geometry("850x460")

config_name = 'config.ini'
config_file = Path('TPV-Skjema/config.ini')

FONT1 = ("Calibri", 11)
FONT2 = ("Calibri", 11, "bold", "underline")

tpv = TPV_Main()

win.mainloop()
