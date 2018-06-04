import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as mBox
from pathlib import Path
import os.path

from configobj import ConfigObj

from TPV import TPV_Main

# TODO
# Fix up the logging so it actually logs any usefull error message; Mostly done, exposing error through the GUI
# Make an input field for oil filling besides entries that requires oil
# Make it possible to edit any configs from the applications itself
# Show the dates for when the next maintainance should be done
# Make the color change persist if the maintainance havent beed done on time; Taken out of the list for the time beeing.
# Make the saving of a new excel file silent when switching year

class SomeWindow(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)

        self.master = master
        self.master.title('TPV Skjema')
        try:
            self.master.iconbitmap('tpv.ico')
        except:
            self._generic_error()


        self.config_folder = Path('configs')
        self.tabcontroll = ttk.Notebook(master)

        self._config_generation()

        names = []

        for i in self.config['Maskin Navn'].values():
            names.append(i)

        configs = []

        for i in self.config['Konfigurasjonsfiler'].values():
            config_path = self.config_folder / i
            configs.append(str(config_path))

        machines = self.config['Maskin info']['Antall Maskiner']
        machines = int(machines)

        try:
            for i in range(machines):
                TPV_Main(self.master, self.tabcontroll, names[i], configs[i])
        except IndexError as e:
            self._config_popup(e)

        self.master.after(1000, self.win_size)

        size = self.config['Diversje']['1']
        self.master.geometry(size)

        self.pack()

    def _config_generation(self):

        """ We are dealing with all config related stuff concerning the main app in this method """

        app_config = self.config_folder / 'app_config.ini'

        if os.path.isfile(str(app_config)) is True:
            self.config = ConfigObj(str(app_config))

        else:
            config = ConfigObj(encoding='utf8', default_encoding='utf8')
            config.filename = app_config

            config['Maskin info'] = {}
            config['Maskin info']['Antall Maskiner'] = '1'

            config['Maskin Navn'] = {}
            config['Maskin Navn']['1'] = 'Du bestemmer 1'
            config['Maskin Navn']['2'] = 'Du bestemmer 2'
            config['Maskin Navn']['3'] = 'Du bestemmer 3'

            config['Konfigurasjonsfiler'] = {}
            config['Konfigurasjonsfiler']['config1'] = 'config1.ini'

            config['Diversje'] = {}

            config.write()
            self.config = ConfigObj(str(app_config))

    def _config_popup(self, message):

        mBox.showerror('Config Issues', '{}'.format(message))

    def _generic_error(self):

        mBox.showerror('Error', 'Something went wrong!')

    def win_size(self):

        """Resizing the UI window to the minimum needed + a 100 pixels on each side"""

        width = self.master.winfo_reqwidth()
        height = self.master.winfo_reqheight()

        width += 100
        height += 100

        width = str(width)
        height = str(height)

        size = width + 'x' + height

        self.config['Diversje']['1'] = size
        self.config.write()
        




root = tk.Tk()
app = SomeWindow(master=root)

if __name__ == '__main__':
    app.mainloop()