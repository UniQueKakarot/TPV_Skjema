import tkinter as tk
from tkinter import ttk

class MyApp(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.counter = 1
        self.pack()
        self.test_widgets(10)

    def test_widgets(self, num):
        check_boxes = {}
        self.results = {}
        for i in range(num):
            self.results['checkvar{0}'.format(i)] = i
            
        for item in self.results:
            check_boxes['check{0}'.format(i)] = ttk.Checkbutton(self.master, variable=self.results[item]).pack()
            

        button = ttk.Button(self.master, text='Lagre', command=self.save)
        button.pack()

    def save(self):
        print(self.results)
        pass






root = tk.Tk()
app = MyApp(master=root)
app.mainloop()


