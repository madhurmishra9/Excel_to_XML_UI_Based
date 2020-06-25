from bs4 import BeautifulSoup as bs
import xlrd as read_excel
import datetime
import lxml

def check_xml_data(filename, saved_path):
    list = []
    value = 0
    try:
        #filename = input('Input xls/xlsx filename here:')
        #saved_path = input('Please provide path where you want to save xml files:')
        wb = read_excel.open_workbook(filename)#read xls or xlsx
        sheetnames = wb.sheet_names()
    except Exception as e:
        list.append(e)
        value = 3
        #print(e)
        #return (list,value)
    #get all sheets names
    #print(sheetnames)
    for sheet in sheetnames:# traves in sheets one by one
        #print('in %s' %(sheet))
        sheets = wb.sheet_by_name(sheet)#activate the current sheet
        ros = sheets.nrows
        cols = sheets.ncols
        try:
            xml_file = open(saved_path + '\\' + sheet + '.xml',"r")#open current sheets equivalent xml if present else print error message
        except Exception as e:
            list.append(e)
            value = 3
            #return (list,value)
            #print(e)
            #print ('file not found')
            continue  #Don't stop even xml for current sheet doesn't exists...continue to next sheet
        contents = xml_file.readlines()# reading xml file data
        contents = "".join(contents)# identify and add to next line
        soup = bs(contents,'lxml') #parse the read content back to xml for beautifulsoup

        columns = [] # defined a new list to store data of all columns
        docu = 'excel'
        docu1 = 'xml'
        for i in range(0,1000):
            column = 'column' + str(i+1)
            if soup.find_all(column) != []:
                columns.append(soup.find_all(column))# appending all columns data by column wise i.e all column1 data at index 0
            else:
                break
        try:
            if len(columns[0]) > sheets.nrows:
                ros = len(columns[0])
                #print ('xml rows greater')
                docu = 'xml'
                docu1 = 'excel'
                if len(columns) > sheets.ncols:
                    cols = len(columns)
                elif len(columns) < sheets.ncols:
                    cols = sheets.ncols
                else:
                    cols =sheets.ncols
            elif len(columns[0]) < sheets.nrows:
                ros = sheets.nrows
                #print('xls rows greater')
                if len(columns) > sheets.ncols:
                    cols = len(columns)
                elif len(columns) < sheets.ncols:
                    cols = sheets.ncols
                else:
                    cols =sheets.ncols
            elif len(column[0]) == sheets.nrows:
                ros = sheets.nrows
                if len(columns) > sheets.ncols:
                    cols = len(columns)
                elif len(columns) < sheets.ncols:
                    cols = sheets.ncols
                else:
                    cols =sheets.ncols
            else:
                ros = sheets.nrows
                cols = sheets.ncols
        except:
            pass
        for row in range(0,ros): # getting number of rows and traversing them one by one
            #print(ros)
            for col in range(0,cols): #getting number of cols and traversing them one by one
                #print (cols)
                try:
                    if (str(columns[col][row].get_text()) == str(sheets.cell(row,col).value)) :
                        pass
                    elif sheets.cell(row,col).ctype is read_excel.XL_CELL_DATE and columns[col][row].get_text() == read_excel.xldate_as_datetime(sheets.cell(row,col).value, wb.datemode).strftime('%m/%d/%Y'):
                        pass
                    else:
                        if columns[col][row].get_text() == '': #checks if xml value for a column is empty
                            Value = None
                        else:
                            Value = columns[col][row].get_text()
                        if sheets.cell(row,col).value == '':# checks if xml cell is empty
                            Cell_Value = None
                        elif sheets.cell(row,col).ctype is read_excel.XL_CELL_DATE:# checks if the xml cell have a date and formats it if the date differs than xml cell value
                            #try:
                                Cell_Value = read_excel.xldate_as_datetime(sheets.cell(row,col).value, wb.datemode).strftime('%m/%d/%Y')
                            #except:
                             #   pass
                        else:
                            Cell_Value = sheets.cell(row,col).value
                        
                        #print('data differs in %s at %d row and %d column - current value in xls is %s and in xml is %s'%(sheet,row+1,col+1,Cell_Value,Value))
                        list.append('data differs in %s at %d row and %d column - current value in xls is %s and in xml is %s'%(sheet,row+1,col+1,Cell_Value,Value))
                        value = 1
                except:
                    if docu == 'xml':
                        if columns[col][row].get_text() != '':
                            #print('data differs in %s at %d row and %d column - current value in %s is %s not found in %s'%(sheet,row+1,col+1,docu,columns[col][row].get_text(),docu1))
                            list.append('data differs in %s at %d row and %d column - current value in %s is %s not found in %s'%(sheet,row+1,col+1,docu,columns[col][row].get_text(),docu1))
                    elif docu == 'excel':
                        if sheets.cell(row,col).value!= '':
                            #print('data differs in %s at %d row and %d column - current value in %s is %s not found in %s'%(sheet,row+1,col+1,docu,sheets.cell(row,col).value,docu1))
                            list.append('data differs in %s at %d row and %d column - current value in %s is %s not found in %s'%(sheet,row+1,col+1,docu,sheets.cell(row,col).value,docu1))
                    value = 1
    return (list,value)

#check_xml_data()
