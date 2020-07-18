import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import export_file
from tkinter import messagebox


class issue_tracker():
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("App Sample")
        #self.window.wm_iconbitmap('sec_logo.ico')
        self.window.resizable(False, False)
        self.main()


        self.radio_variable = tk.StringVar()
        self.combobox_value = tk.StringVar()


    def main(self):
        self.window['padx'] = 5
        self.window['pady'] = 5

        def ExitApplication():
            choice = tk.messagebox.askquestion('Exit Application', 'Are you sure to close app? All data will be lost')

            if choice == 'yes':
                self.window.destroy()

        self.window.protocol("WM_DELETE_WINDOW", ExitApplication)

        frame1 = ttk.LabelFrame(self.window, text="Frame 1 ", relief=tk.RIDGE)
        frame1.grid(row=1, column=1, sticky=tk.E + tk.W + tk.N + tk.S)

        label1 = ttk.Label(frame1, text="Info 1 ")
        label1.grid(row=1, column=1, pady=5, padx=5, sticky=tk.W)

        label2 = ttk.Label(frame1, text="Info 2 ")
        label2.grid(row=2, column=1, pady=5, padx=5, sticky=tk.W)

        label3 = ttk.Label(frame1, text="Info 3 ")
        label3.grid(row=3, column=1, pady=5, padx=5, sticky=tk.W)

        label4 = ttk.Label(frame1, text="Drop Down 1 ")
        label4.grid(row=4, column=1, pady=5, padx=3, sticky=tk.W)

        label5 = ttk.Label(frame1, text="Drop Down 2 ")
        label5.grid(row=5, column=1, pady=5, padx=3, sticky=tk.W)

        label5 = ttk.Label(frame1, text="Text Box \n (if applicable)")
        label5.grid(row=6, column=1, pady=5, padx=5, sticky=tk.W)

        # Entry box and Text box

        entry1 = ttk.Entry(frame1, justify='left')
        entry1.grid(row=1, column=2, pady=3, sticky=tk.W)

        entry2 = ttk.Entry(frame1, justify='left')
        entry2.grid(row=2, column=2, pady=3, sticky=tk.W)

        entry3 = ttk.Entry(frame1, justify='left')
        entry3.grid(row=3, column=2, pady=3, sticky=tk.W)

        entry4 = ttk.Combobox(frame1, state='readonly',
                              values=('Mode 1 ', 'Mode 2 '))
        entry4.current(0)
        entry4.grid(row=4, column=2, pady=3, padx=3, sticky=tk.W)

        entry5 = ttk.Combobox(frame1, state='readonly', values=('M1', 'M2' ,'M3', 'M4', 'M5', 'M6'))
        entry5.current(0)
        entry5.grid(row=5, column=2, pady=3, padx=3, sticky=tk.W)

        entry6 = tk.Text(frame1, height = 5, width = 25)
        entry6.grid(row=6, column=2, pady=3, padx=3, sticky=tk.W)

        save_button = ttk.Button(frame1, text='Add to Chart >> ', command=lambda:export_file.to_chart(entry1,entry2,entry3,entry4,entry5,entry6, chat1))
        save_button.grid(row=7, column=1, pady=3, padx=3, sticky= tk.W)

        save_button = ttk.Button(frame1, text='Remove from Chart << ',
                                 command=lambda: export_file.remove_from_chart(chat1))
        save_button.grid(row=7, column=2, pady=3, padx=3, sticky=tk.W)

        Export_to_Excel = ttk.Button(frame1, text='Export chart to .xls file ',command=lambda:export_file.write_to_excel(chat1,self.window))
        Export_to_Excel.grid(row=8, column=1, pady=3, padx=3, sticky=tk.W)


        frame2 = ttk.LabelFrame(self.window, text="Frame 2 ", relief=tk.RIDGE)
        frame2.grid(row=1, column=2, sticky=tk.E + tk.W + tk.N + tk.S, padx=15)

        chat1 = ttk.Treeview(frame2, height=13, columns=('Info 1','Info 2', 'Info 3', 'Drop Down 1', 'Drop Down 2', 'Text Box 1'), selectmode="extended")

        scroll_bar = ttk.Scrollbar(frame2, orient="vertical", command=chat1.yview)
        scroll_bar.grid(row=1, column=5, sticky='ns')
        chat1.configure(yscrollcommand=scroll_bar.set)

        chat1.heading('#0', text='#', anchor=tk.CENTER)
        chat1.heading('#1', text='Info 1', anchor=tk.CENTER)
        chat1.heading('#2', text='Info 2', anchor=tk.CENTER)
        chat1.heading('#3', text='Info 3', anchor=tk.CENTER)
        chat1.heading('#4', text='Drop Down 1', anchor=tk.CENTER)
        chat1.heading('#5', text='Drop Down 2', anchor=tk.CENTER)
        chat1.heading('#6', text='Text Box 1', anchor=tk.CENTER)

        chat1.column("#0", width=2)
        chat1.column('#1', stretch=tk.YES, minwidth=50, width=100)
        chat1.column('#2', stretch=tk.YES, minwidth=50, width=100)
        chat1.column('#3', stretch=tk.YES, minwidth=50, width=120)
        chat1.column('#4', stretch=tk.YES, minwidth=50, width=100)
        chat1.column('#5', stretch=tk.YES, minwidth=50, width=100)
        chat1.column('#6', stretch=tk.YES, minwidth=50, width=120)

        chat1.grid(row=1, column=1, columnspan=4, sticky='nsew')

        frame3 = ttk.LabelFrame(self.window, text="", relief=tk.RIDGE)
        frame3.grid(row=2, column=1, columnspan=5, sticky=tk.E+tk.W)

        misc1 = ttk.Label(frame3, text="App V.1.0 - 7/12/2020")
        misc1.grid(row=1, column=1, pady=1, padx=1, sticky=tk.W)



# Create the entire GUI program
program = issue_tracker()

# Start the GUI event loop
program.window.mainloop()