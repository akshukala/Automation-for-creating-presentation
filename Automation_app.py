# -*- coding: utf-8 -*-
"""
Created on Tue Dec 10 12:59:51 2019

@author: akshay.kale
"""

import tkinter as tk
from datetime import date
from tkinter.ttk import Style
from tkinter import messagebox
from tkinter import filedialog

from constants import MyDateEntry
from QRC_report import qrc_normal_main
from ServiceTrend import service_trend_main
from QRC_report_combine_network import qrc_main
from function_report_automation_new import main_1
from QRC_circlewise_report import circlewise_main
from QRC_circlewise_network import circlewise_network_main

'''Creating a simple UI app'''
root = tk.Tk()
s = Style()
s.theme_use('xpnative')
root.title("QRC Analysis")
root.geometry('550x450')
root.resizable(False, False) 

lbl1 = tk.Label(root, text="QRC Analysis for AT Increase.",
                fg = "#C70039", font = "Calibri 20 bold")
lbl1.grid(column=2, row=0)

'''Creating a datepicker'''
lbl2 = tk.Label(root, text="Select Start Date:",
                fg = "black", anchor="w", font = "Calibri 12")
start_date = MyDateEntry(root, year=date.today().year, month=date.today().month, day=date.today().day,
                 selectbackground='gray80', selectforeground='black', normalbackground='white',
                 normalforeground='black', background='gray90', foreground='black', bordercolor='gray90',
                 othermonthforeground='gray50', othermonthbackground='white', othermonthweforeground='gray50',
                 othermonthwebackground='white', weekendbackground='white', weekendforeground='black',
                 headersbackground='white', headersforeground='gray70', anchor="w")
lbl2.grid(column=1, row=1)
start_date.grid(column=2, row=1)

lbl3 = tk.Label(root, text="Select End Date:",
                fg = "black", anchor="w", font = "Calibri 12")
end_date = MyDateEntry(root, year=date.today().year, month=date.today().month, day=date.today().day,
                 selectbackground='gray80', selectforeground='black', normalbackground='white',
                 normalforeground='black', background='gray90', foreground='black', bordercolor='gray90',
                 othermonthforeground='gray50', othermonthbackground='white', othermonthweforeground='gray50',
                 othermonthwebackground='white', weekendbackground='white', weekendforeground='black',
                 headersbackground='white', headersforeground='gray70', anchor="w")
lbl3.grid(column=1, row=2)
end_date.grid(column=2, row=2)

'''Creating grouping of days'''
lbl4 = tk.Label(root, text="Group No. Of Days:",
                fg = "black", anchor="w", font = "Calibri 12")
spin_box = tk.Spinbox(root, from_=1, to=100, width=5)
lbl4.grid(column=1, row=3)
spin_box.grid(column=2, row=3)

'''Browsing data files for input'''
def ivr_browse_button():
    #filename = filedialog.askdirectory()
    filename = filedialog.askdirectory()
    ivr_file_path.set(filename)
    
ivr_file_path = tk.StringVar()
lbl8 = tk.Label(root, text="IVR File Directory: ", fg = "black", anchor="w", font="Calibri 12")
ivr_entry = tk.Entry(root, width=45, textvariable=ivr_file_path)
ivr_browsebutton = tk.Button(root, text="Browse", command = ivr_browse_button)
lbl8.grid(column=1, row=4)
ivr_entry.grid(column=2, row=4)
ivr_browsebutton.grid(column=3, row=4)


def st_browse_button():
    #filename = filedialog.askdirectory()
    filename = filedialog.askopenfilename()
    file_path.set(filename)
    
file_path = tk.StringVar()
lbl5 = tk.Label(root, text="Service Trend File Path: ", fg = "black", anchor="w", font="Calibri 12")
st_entry = tk.Entry(root, width=45, textvariable=file_path)
st_browsebutton = tk.Button(root, text="Browse", command = st_browse_button)
lbl5.grid(column=1, row=5)
st_entry.grid(column=2, row=5)
st_browsebutton.grid(column=3, row=5)

def output_browse_button():
    #filename = filedialog.askdirectory()
    filename = filedialog.askdirectory()
    output_file_path.set(filename)
    
output_file_path = tk.StringVar()
lbl7 = tk.Label(root, text="Output File Directory: ", fg = "black", anchor="w", font="Calibri 12")
out_entry = tk.Entry(root, width=45, textvariable=output_file_path)
out_browsebutton = tk.Button(root, text="Browse", command = output_browse_button)
lbl7.grid(column=1, row=6)
out_entry.grid(column=2, row=6)
out_browsebutton.grid(column=3, row=6)

'''Create a presentation button'''
def generatePPT():
    start = start_date.get()
    end = end_date.get()
    ivr_path = ivr_entry.get() + "/"
    service_trend_path = st_entry.get()
    output_path = out_entry.get() + "/Analysis.pptx"
    res = main_1(start, end, ivr_path, output_path)
    res1 = service_trend_main(start, end, service_trend_path, output_path)
    messagebox.showinfo("Information","Successfully created IVR Plots.")

btn = tk.Button(root, text="Overall Report", command = generatePPT)
btn.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
#btn.grid(row=4,column=1, padx=(45,5))

'''Fetching data for further analysis'''
selected = tk.IntVar()
rad2 = tk.Radiobutton(root,text='QRC Normal', value=2, variable=selected)
rad1 = tk.Radiobutton(root,text='QRC considering Network Related', value=1, variable=selected)
rad1.place(relx=0.2, rely=0.6, anchor=tk.CENTER)
rad2.place(relx=0.6, rely=0.6, anchor=tk.CENTER)

def qrc_browse_button():
    #filename = filedialog.askdirectory()
    filename = filedialog.askopenfilename()
    qrc_file_path.set(filename)

qrc_file_path = tk.StringVar()
lbl6 = tk.Label(root, text="QRC File Path: ", fg = "black", anchor="w", font="Calibri 12")
qrc_entry = tk.Entry(root, width=45, textvariable=qrc_file_path)
browsebutton1 = tk.Button(root, text="Browse", command = qrc_browse_button)
lbl6.place(relx=0.1, rely=0.7, anchor=tk.CENTER)
qrc_entry.place(relx=0.6, rely=0.7, anchor=tk.CENTER)
browsebutton1.place(relx=0.93, rely=0.7, anchor=tk.CENTER)

'''Updating the existing presentation'''
def generateQRC():
    start = start_date.get()
    end = end_date.get()
    group = spin_box.get()
    radio = selected.get()
    output_path = out_entry.get() + "/Analysis.pptx"
    if radio == 1:
        res2 = qrc_main(start, end, int(group)-1, qrc_entry.get(), output_path)
        res3 = circlewise_network_main(start, end, int(group)-1, output_path)
    else:
        res2 = qrc_normal_main(start, end, int(group)-1, qrc_entry.get(), output_path)
        res3 = circlewise_main(start, end, int(group)-1, output_path)
    messagebox.showinfo("Information","Successfully ceated QRC Plots.")
    
btn = tk.Button(root, text="QRC Report", command = generateQRC)
btn.place(relx=0.5, rely=0.8, anchor=tk.CENTER)

root.mainloop()










# =============================================================================
# import PySimpleGUI as sg
# sg.ChangeLookAndFeel('DarkAmber')
# layout = [      
#     [sg.Text('QRC Analysis for AT Increase', size=(30, 1), font=("Helvetica", 22))],      
#     [],
#     [sg.Text('Enter Start Date: '), sg.InputText('This is my text')],      
#     [],      
#     [sg.Text('_'  * 80)],      
#     [sg.Text('Choose A Folder', size=(35, 1))],      
#     [sg.Text('Your Folder', size=(15, 1), auto_size_text=False, justification='right'),      
#      sg.InputText('Default Folder'), sg.FolderBrowse()],      
#     [sg.Submit(), sg.Cancel()]      
# ]
# window = sg.Window('Everything bagel', default_element_size=(40, 10)).Layout(layout)
# button, values = window.Read()
# sg.Popup(button, values)
# =============================================================================
