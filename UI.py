from tkinter import filedialog
from tkinter import messagebox
import convert_excel_to_xml
import compare_excel_to_xml
from tkinter import ttk
from tkinter import *


filename = ''
foldername = ''

def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir = "/",title = "Browse xls/xlsx file",filetypes = (("Excel","*.xls"),("Excel Workbook","*.xlsx")))
    if filename == '':
        entry_filed.configure(text="Browse xls/xlsx file")
    else:
        entry_filed.configure(text=filename)
    #return entry_filed

def browseFolder():
    global foldername
    foldername = filedialog.askdirectory(initialdir = "/",title = "Select folder",)
    if filename == '':
        entry_folder.configure(text = "Select folder")
    else:
        entry_folder.configure(text=foldername)
    #return entry_folder

def convert():
    global filename
    global foldername
    if filename == '':
        ttk.Label(scrollable_frame).pack()
        messagebox.showwarning('Warning','Select excel file to convert')
    elif foldername == '':
        ttk.Label(scrollable_frame).pack()
        messagebox.showwarning('Warning','Select folder to Store XML')
    else:
        if convert_excel_to_xml.convert_xls_to_xml(filename,foldername) == None:
            messagebox.showinfo('Success', "Excel to XML convertion successfull")
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Select folder")
            filename = ''
            foldername = ''
        else:
            messagebox.showwarning("Failed","Error occurred while convertion")

def compare():
    global filename
    global foldername
    if filename == '':
        ttk.Label(scrollable_frame).pack()
        messagebox.showwarning('Warning','Select excel file to compare')
    elif foldername == '':
        ttk.Label(scrollable_frame).pack()
        messagebox.showwarning('Warning','Select folder where XML files are stored')
    else:
        list,value= compare_excel_to_xml.check_xml_data(filename,foldername)
        ttk.Label(scrollable_frame).pack()
        if value == 1:
            for i in list:
                ttk.Label(scrollable_frame, text=i).pack()
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Select folder")
            filename = ''
            foldername = ''
        elif value == 0:
            messagebox.showinfo('Success', "100% Match")
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Select folder")
            filename = ''
            foldername = ''
        else:
            for i in list:
                ttk.Label(scrollable_frame, text=i).pack()
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Select folder")
            filename = ''
            foldername = ''

window = Tk()

window.title('XML Converter')


first_label = Label(window, text="File name:")
first_label.grid(row = 0,column = 1)
entry_filed = Label(window, text = "Browse xls/xlsx file", width = 100, bg = 'white')
entry_filed.grid(row = 0, column = 2)
Button(window, text = "Browse", command = browseFiles).grid(row = 0 , column = 3)


first_label = Label(window, text="Folder path:")
first_label.grid(row = 1,column = 1)
entry_folder = Label(window, text = "Select folder", width = 100, bg = 'white')
entry_folder.grid(row = 1, column = 2)
Button(window, text = "Browse", command = browseFolder).grid(row = 1 , column = 3)



container = ttk.Frame(window)
canvas = Canvas(container, width = 720 )
scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)



container.grid(row=2,column=2,sticky= W,pady=4)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")


Button(window,text='Convert', command = convert).grid(row=3,column=0,sticky= W,pady=4)
Button(window,text='Compare', command = compare).grid(row=3,column=1,sticky= W,pady=4)
#Button(window,text='Search').grid(row=3,column=2,sticky= W,pady=4)
Button(window,text='Quit',command = window.quit).grid(row=3,column=2,sticky= W,pady=4)
#test_scrollbar = Scrollbar(output_label, orient = 'vertical',command = convert.yview)
# Code to add widgets will go here...
window.mainloop()
