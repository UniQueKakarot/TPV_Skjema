import tkinter as tk
from tkinter import ttk

from Test import MyApp

class SomeWindow(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        self.tabcontroll = ttk.Notebook(master)

        MyApp(self.master, self.tabcontroll, 'tab1', 'Test1')
        MyApp(self.master, self.tabcontroll, 'tab2', 'Test2')
        MyApp(self.master, self.tabcontroll, 'tab3', 'Test3')
        MyApp(self.master, self.tabcontroll, 'tab4', 'Test4')
        MyApp(self.master, self.tabcontroll, 'tab5', 'Test5')
        #self.tab1 = ttk.Frame(self.tabcontroll)
        #self.tab2 = ttk.Frame(self.tabcontroll)

        #self.tabcontroll.add(self.tab1, text='Tab 1')
        #self.tabcontroll.add(self.tab2, text='Tab 2')

        self.pack()



root = tk.Tk()
app = SomeWindow(master=root)

if __name__ == '__main__':
    app.mainloop()