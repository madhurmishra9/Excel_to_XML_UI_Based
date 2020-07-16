import pandas as pd
import xlrd as read_xls

def transpose():
    try:
        filename = input('File name :')
        wb = read_xls.open_workbook(filename)
        sheetnames = wb.sheet_names()
    except Exception as e:
        print(e)
        return 1

    if filename.endswith('.xls') :
        new_transpose_file_name = filename[:-len('.xls')]
        new_transpose_file_name = new_transpose_file_name + '_transpose.xls'
        writer = pd.ExcelWriter(new_transpose_file_name, engine='xlwt')
    elif filename.endswith('.xlsx'):
        new_transpose_file_name = filename[:-len('.xlsx')]
        new_transpose_file_name = new_transpose_file_name + '_transpose.xlsx'
        writer = pd.ExcelWriter(new_transpose_file_name, engine='openpyxl')

    for sheet in sheetnames:
        df = pd.read_excel(filename, sheet_name= sheet)
        df.transpose().to_excel(writer, sheet_name= sheet)
    writer.save()


#transpose()
