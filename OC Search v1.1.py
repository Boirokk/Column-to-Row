# Written by Boirokk 2016-04-21 v1.1

import xlrd
import os
import csv


# Look for query in each sheet and print rows containing query
def find_in_workbook_sheets(file_location,new_file_name):

    # Open file and store in workbook variable
    workbook = xlrd.open_workbook(file_location)
   

    # Open each sheet, search for criteria entered by user and write matching rows to new file
    sheet = workbook.sheet_by_index(0)
    h_row = []
    for index in range(sheet.nrows):
        h_row.append(sheet.cell_value(index, 1))
  
    with open(new_file_name,'a',newline='') as fp:
        a =  csv.writer(fp,delimiter=',')
        data = [h_row]
        a.writerows(data)
        


# Get the .xls and .xlsx files from the root and sub dirs
def get_file_contents(new_file_name):

    
    file_location = r'S:\Yachts\ORDER CONF\2016'

    for roots, dirs, files in os.walk(file_location):
        for file in files:
            file_name = roots + '\\' + file
           
            if '.xlsx' in file_name:
                try:
                    find_in_workbook_sheets(file_name,new_file_name)
                except:
                    continue
            elif '.xls' in file_name:
                try:
                    find_in_workbook_sheets(file_name,new_file_name)
                except:
                    continue
            
# main
def main():

    new_file_name = "OC Data Search 2016.csv"
    print('Creating file... ', new_file_name)
    
    # Open new .xls file and insert headers
    try:
        with open(new_file_name,'w',newline='') as fp:
            a =  csv.writer(fp,delimiter=',')
            data = [['DATE', 'OUR REFERENCE 1', 'OUR REFERENCE 2', 'PREPARER\'S INITIALS', 'DOCUMENT TYEP',
                     'WORKORDER #', '', 'NAME OF ACCOUT', 'ADDRESS', 'CITY,STATE,ZIP', 'COUNTRY',
                     'NOTIFY ADDRESS', 'NOTIFY TEL', 'NOTIFY FAX', 'NOTIFY EMAIL', '', 'PURCHASE REFERENCE 1',
                     'PURCHASE REFERENCE 2','PURCHASE REFERENCE 3','PURCHASE REFERENCE 4','PO NUMBER',
                     'PURCHASE REFERENCE TEL','PURCHASE REFERENCE FAX','PURCHASE REFERENCE E-MAIL',
                     '','CONSIGNEE 1','CONSIGNEE 2','CONSIGNEE 3','CONSIGNEE 4','CONSIGNEE 5', 'CONSIGNEE TEL', 'CONSIGNEE FAX',
                     'CONSIGNEE E-MAIL','','MANUFACTURER/MODEL','MARKS','DATE OF DELIVERY','DESTINATION','','SHIPPING SECTION',
                     '','PAYMENT TERMS','','','CURRENCY','','USD $','EUR €','GBP £','','','DECK AREA (ft2)','DECK PRICE / ft2',
                     'DECK INSTALLATION COST / ft2','SUB-DECK PRICE / ft2','SUB-DECK INSTALLATION COST / ft2','STEPS (#)',
                     'PRICE PER STEPS','CTG CAULK REQ-D','CTG CAULK PRICE','GAL FITTING EPOXY REQ-D','GAL FITTING EPOXY PRICE',
                     'CRATING PRICE','','ITEMIZED','','FLORIDA SALES TAX PERCENTAGE','CUSTOMER TAX EXEMPT NUMBER']]
            a.writerows(data)
    except:
        print('Please close the Accumulated Data Search Document and try again.')
        error = input('Press enter to exit')
        exit()


    get_file_contents(new_file_name)
    print('Done...')
    input('Press enter to exit')
    
# Call main
main()
