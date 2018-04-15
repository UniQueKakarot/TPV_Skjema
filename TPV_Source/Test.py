import tkinter as tk
from tkinter import ttk

class MyApp():
    def __init__(self, master, tab, tabvar, name):
        self.master = master
        self.tab = tab
        self.tabvar = tabvar

        self.tabvar = ttk.Frame(self.tab)

        self.tab.add(self.tabvar, text=name)
        self.tab.pack(expand=1, fill="both")

        self.test_widgets(10)

    def test_widgets(self, num):
        check_boxes = {}
        self.results = {}
        for i in range(num):
            self.results['checkvar{0}'.format(i)] = tk.IntVar()
            
        for item in self.results:
            check_boxes['check{0}'.format(i)] = ttk.Checkbutton(self.tabvar, variable=self.results[item]).pack()
            

        button = ttk.Button(self.tabvar, text='Lagre', command=self.save)
        button.pack()

    def save(self):
        test = []

        for i in self.results:
            test.append(self.results[i])

        #why the fuck do we have to append to a list and run .get to get the actual values??????????
        for i in test:
            print(i.get())

        #print(self.test.get())
        #print(self.results['checkvar0'])







#root = tk.Tk()
#app = MyApp(master=root)
#app.mainloop()


