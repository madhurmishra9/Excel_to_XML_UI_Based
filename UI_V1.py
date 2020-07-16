from tkinter import filedialog
from tkinter import messagebox
import convert_excel_to_xml
import compare_excel_to_xml
from tkinter import ttk
from tkinter import *
import transpose


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
    foldername = filedialog.askdirectory(initialdir = "/",title = "Browse folder",)
    if filename == '':
        entry_folder.configure(text = "Browse folder")
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
        list,value,xml_ros,xml_cols,excel_rows,excel_cols,sheetnames= compare_excel_to_xml.check_xml_data(filename,foldername)
        val = ''
        val1 = ''

        for i in range(0,len(sheetnames)):
            val = val + sheetnames[i] + ':' + ' Rows:' + str(excel_rows[i]) + ' Cols:' + str(excel_cols[i]) + '\n'
            val1 = val1 + sheetnames[i] + ':' + ' Rows:' + str(xml_ros[i]) + ' Cols:' + str(xml_cols[i]) + '\n'

        excel_value_label.configure(text=val)
        XML_value_label.configure(text=val1)
        ttk.Label(scrollable_frame).pack()
        if value == 1:
            for i in list:
                ttk.Label(scrollable_frame, text=i).pack()
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Browse folder")
            filename = ''
            foldername = ''
        elif value == 0:
            messagebox.showinfo('Success', "100% Match")
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Browse folder")
            filename = ''
            foldername = ''
        else:
            for i in list:
                ttk.Label(scrollable_frame, text=i).pack()
            entry_filed.configure(text="Browse xls/xlsx file")
            entry_folder.configure(text="Browse folder")
            filename = ''
            foldername = ''

def transpos():
    messagebox.showwarning('Warning','Data may become useless')
    transpose.transpose(filename)
    messagebox.showinfo(title='Message',message='A transpose file is generated with _transpose added to filename.')



window = Tk()
window.resizable(0,0)
window.title('XML Converter')

frame = Frame(window, relief=RAISED, borderwidth=1)
frame.pack(fill=BOTH, expand=True)

frame1 = Frame(frame)
frame1.pack(fill=X)

first_label = Label(frame1, text="File name:",width=14)
first_label.grid(row = 0,column = 1, padx = 5, pady = 5)

entry_filed = Label(frame1, text = "Browse xls/xlsx file", bg = 'white', width = 90)
entry_filed.grid(row = 0, column = 2)

button1 = Button(frame1, text = "Browse", command = browseFiles)
button1.grid(row = 0 , column = 3,  padx = 5, pady = 4)

frame2 = Frame(frame)
frame2.pack(fill=X)

second_label = Label(frame2, text="Folder path:",width = 14)
second_label.grid(row = 0,column = 1, padx = 5, pady = 5)

entry_folder = Label(frame2, text = "Browse folder", bg = 'white', width = 90)
entry_folder.grid(row = 0, column = 2)

button2 = Button(frame2, text = "Browse", command = browseFolder)
button2.grid(row = 0 , column = 3,  padx = 5, pady = 4)

frame3 = Frame(frame)
frame3.pack(fill=BOTH)

excel_label = Label(frame3, text="Excel Rows/Cols:")
excel_label.pack(side=LEFT, anchor=N, padx=5, pady=5)

excel_value_label = Label(frame3, height = 10, width = 40, bg = 'white')
excel_value_label.pack(side=LEFT, padx=10, pady=10)


XML_label = Label(frame3, text = "XML Rows/Cols:")
XML_label.pack(side=LEFT, anchor=N, padx=5, pady=5)

XML_value_label = Label(frame3, height = 10,width = 40, bg= 'white')
XML_value_label.pack(side=LEFT, padx=10, pady=10)

frame4 = Frame(frame)
frame4.pack(fill=BOTH)

label3 = Label(frame4, text = 'Difference:', width = 14)
label3.pack(side = LEFT)

container = ttk.Frame(frame4)
canvas = Canvas(container, width =671, bg = 'white')
scrollbar = ttk.Scrollbar(container, orient="vertical",command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor=NW)
canvas.configure(yscrollcommand=scrollbar.set)

container.pack(pady=4)
canvas.pack(side=LEFT, fill=BOTH, expand=True)
scrollbar.pack(side="right", fill="y")


convert = Button(window,text='Convert', command = convert).pack(side= LEFT, padx = 5, pady = 4)
compare = Button(window,text='Compare', command = compare).pack(side= LEFT, padx = 5, pady = 4)
Transpose = Button(window,text='Transpose', command = transpos).pack(side= LEFT, padx = 5, pady = 4)
quit = Button(window,text='Quit',command = window.quit).pack(side= RIGHT, padx = 5, pady = 4)
reset = Button(window,text='Reset').pack(side= RIGHT, padx = 5, pady = 4)



window.mainloop()
