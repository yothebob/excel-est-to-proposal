from openpyxl import Workbook, load_workbook
import sys
import os
from datetime import date

def write_to_log(filepath,data):
    '''
    Input : filepath (path to file- filename,String)
            data (dictionary of proposal log {est number : [details, ..., ...]})
    Output : NA
    '''
    sys.path.append(filepath)
    os.chdir(filepath)

    filename = "Proposal_log.xlsx"
    workbook = load_workbook(filename=filename, data_only=True)
    proposal_log = workbook.active

    estimate_number = [*data][0]
    data_values = [value for value in data.values()]

    today = date.today()
    data_values[0].insert(0,str(today.strftime("%m/%d/%Y")))
    data_values[0].insert(0,estimate_number)

    log_data = [data_values[0][item] for item in range(9)]
    proposal_log.append(log_data)
    workbook.save(filename)
