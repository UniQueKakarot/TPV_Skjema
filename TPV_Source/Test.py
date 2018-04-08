import tkinter as tk
from tkinter import ttk

class MyApp(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.tabcontroll = ttk.Notebook(master)
        self.tab1 = ttk.Frame(self.tabcontroll)
        self.tab2 = ttk.Frame(self.tabcontroll)
        self.tabcontroll.add(self.tab1, text='Tab 1')
        self.tabcontroll.add(self.tab2, text='Tab 2')
        self.tabcontroll.pack()
        self.pack()
        self.test_widgets(10)

    def test_widgets(self, num):
        check_boxes = {}
        self.results = {}
        for i in range(num):
            self.results['checkvar{0}'.format(i)] = tk.IntVar()
            
        for item in self.results:
            check_boxes['check{0}'.format(i)] = ttk.Checkbutton(self.tab1, variable=self.results[item]).pack()
            

        button = ttk.Button(self.tab1, text='Lagre', command=self.save)
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







root = tk.Tk()
app = MyApp(master=root)
app.mainloop()


