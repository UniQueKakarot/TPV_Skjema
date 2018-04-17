import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as mBox
import os.path

from configobj import ConfigObj

from TPV import TPV_Main

class SomeWindow(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)

        self.master = master
        self.tabcontroll = ttk.Notebook(master)

        self.config_generation()

        names = []

        for i in self.config['Maskin Navn'].values():
            names.append(i)

        configs = []

        for i in self.config['Konfigurasjonsfiler'].values():
            configs.append(i)

        machines = self.config['Maskin info']['Antall Maskiner']
        machines = int(machines)

        try:
            for i in range(machines):
                TPV_Main(self.master, self.tabcontroll, names[i], configs[i])
        except IndexError:
            self.config_popup()

        self.pack()

    def config_generation(self):

        """ We are dealing with all config related stuff concerning the main app in this method """

        if os.path.isfile('app_config.ini') is True:
            self.config = ConfigObj('app_config.ini')

        else:
            config = ConfigObj(encoding='utf8', default_encoding='utf8')
            config.filename = 'app_config.ini'

            config['Maskin info'] = {}
            config['Maskin info']['Antall Maskiner'] = '1'

            config['Maskin Navn'] = {}
            config['Maskin Navn']['1'] = 'Du bestemmer 1'
            config['Maskin Navn']['2'] = 'Du bestemmer 2'
            config['Maskin Navn']['3'] = 'Du bestemmer 3'

            config['Konfigurasjonsfiler'] = {}
            config['Konfigurasjonsfiler']['config1'] = 'config1.ini'

            config.write()

    def config_popup(self):

        mBox.showerror('', 'Noe er galt med app_config.ini fila')
        




root = tk.Tk()
app = SomeWindow(master=root)

if __name__ == '__main__':
    app.mainloop()