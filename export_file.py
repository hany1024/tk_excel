import xlsxwriter
from tkinter import filedialog
from tkinter import simpledialog
import tkinter as tk
from tkinter import messagebox


def to_chart(v1,v2,v3,v4,v5,v6,chat1):
    ESN = v1.get()
    equip_num = v2.get()
    test_time = v3.get()
    mode = v4.get()
    err_cat = v5.get()
    detail = v6.get('1.0','end')

    chat1.insert('', 'end', text='',
                             values=(ESN, equip_num,test_time,mode,err_cat,detail))


def remove_from_chart(chat1):
    selected_item = chat1.selection()[0]
    chat1.delete(selected_item)


def write_to_excel(chat1,parent):

    data = []
    for line in chat1.get_children():
        #print(chat1.item(line)['values'])
        data.append(chat1.item(line)['values'])
        #print(chat1.item(line)['values'][0])
        #print(chat1.item(line)['values'][2])

    print(data)
    print(data[0])
    print(data[0][2])
    print(data[0][0])

    answer = simpledialog.askstring("Input", "Set file name (ex: CB3_report)",
                                    parent=parent)
    if len(answer)>0:
        file_name = answer + '.xlsx'
        workbook = xlsxwriter.Workbook('./'+file_name)
        worksheet = workbook.add_worksheet()

        row = 1
        col = 0

        bold = workbook.add_format({'bold': True})

        for i in range(0,6):
            worksheet.set_column(0, i, 15)

        worksheet.write(0,0,"ESN Last 4 #",bold)
        worksheet.write(0,1, "Equip # ",bold)
        worksheet.write(0,2, "Test Time",bold)
        worksheet.write(0,3, "Test Mode",bold)
        worksheet.write(0,4, "Error Source", bold)
        worksheet.write(0,5, "Error Detail", bold)

        for i in range(0,len(data)):
            worksheet.write(row, col, data[i][0])
            worksheet.write(row, col+1, data[i][1])
            worksheet.write(row, col+2, data[i][2])
            worksheet.write(row, col+3, data[i][3])
            worksheet.write(row, col+4, data[i][4])
            worksheet.write(row, col+5, data[i][5])
            row += 1

        row = 0
        messagebox.showinfo("Success", "Your report was exported as " + file_name)
        workbook.close()

    else:
        messagebox.showerror("Error","Please Enter file name")








