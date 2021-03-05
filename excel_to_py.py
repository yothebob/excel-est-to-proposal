from openpyxl import Workbook
from openpyxl import load_workbook

file = 'Estimating Model-current 2.8.21.xlsm' #input('whats the excelfile name? (.xlsm)')

workbook = load_workbook(filename=file)

note_sheet = workbook['Notes']

estimate = []

def return_variables():
    for row in note_sheet.iter_rows(min_row=1,min_col=2,max_col=2,values_only=True):
        res = ''.join(map(str,row))
        estimate.append(res)
    return estimate[0],estimate[1],estimate[2],estimate[3],estimate[4],estimate[5]

def return_lf():
    section_lf= []
    for row in note_sheet.iter_rows(min_row=2,min_col=4,max_col=4,values_only=True):
        res = ''.join(map(str,row))
        if res != 'None':
            section_lf.append(int(res))
    return section_lf

def return_lfprice():
    section_lfprice= []
    for row in note_sheet.iter_rows(min_row=2,min_col=5,max_col=5,values_only=True):
        res = ''.join(map(str,row))
        if res != 'None':
            section_lfprice.append(int(res))
    return section_lfprice

def return_section_details(num=0):
    sections = return_lf()
    section = []
    for row in note_sheet.iter_rows(min_row=2,min_col=(7+num),max_col=(7+num),values_only=True):
            res = ''.join(map(str,row))
            if res != 'None':
                section.append(int(res))
    return section

#print(return_variables())
#return_lf()
#return_lfprice()
#return_section_details(4)

