import xml.etree.cElementTree as ET
import xlrd as read_xls
import datetime

def convert_xls_to_xml(filename, saved_path):
    try:
        #filename = input("Provide filename stored at current location or complete path of file here:")
        #saved_path = input('Please provide path where xml file that you want to compare excel with are located here:')
        wb = read_xls.open_workbook(filename)
        sheetnames = wb.sheet_names()
    except Exception as e:
        print (e)
        return 1
        #input('Please rerun program ,press any key to exit:')
        #exit()
    root = ET.Element("data")
    count = 0
    i = 1
    for sheet in sheetnames:
        count = 0
        sheets = wb.sheet_by_name(sheet)
        #print(sheet)
        sub_root = ET.SubElement(root, sheet)
        for row in sheets.get_rows():
            #print (row)
            count += 1
            doc = ET.SubElement(sub_root, "row" + str(count))
            j = 1
            for cell in row:
                #print(cell)
                cells = (str(cell)).split(':')
                if cells[0] == 'xldate':
                    ET.SubElement(doc, 'column' + str(j)).text = read_xls.xldate.xldate_as_datetime(cell.value, wb.datemode).strftime('%m/%d/%Y')
                else:
                    ET.SubElement(doc,'column' + str(j)).text = str(cell.value)
                j += 1
        tree = ET.ElementTree(root)
        tree.write(saved_path + '\\' + sheet + '.xml')
        root = ET.Element("data")



#convert_xls_to_xml()
