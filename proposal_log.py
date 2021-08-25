from openpyxl import Workbook, load_workbook
import sys
import os
from datetime import date

class WriteToLog:
    def __init__(self,filepath,data):
        '''
        Input : filepath (path to file- filename,String)
                filename (name of file,string)
                data (dictionary of proposal log {est number : [details, ..., ...]})
        Output : NA
        '''
        self.filepath = filepath
        self.filename = "Proposal_log.xlsx"
        self.data = data

    def write_to_log(self):
        #sys.path.append(self.filepath)
        #os.chdir(self.filepath)
        workbook = load_workbook(filename=self.filename, data_only=True)
        proposal_log = workbook.active

        estimate_number = [*self.data][0]
        data_values = [value for value in self.data.values()]

        today = date.today()
        data_values[0].insert(0,str(today.strftime("%m/%d/%Y")))
        data_values[0].insert(0,estimate_number)

        log_data = [data_values[0][item] for item in range(9)]
        proposal_log.append(log_data)
        workbook.save(self.filename)


def _testing():
    test = WriteToLog("C:/Users/Owner/Desktop/Estimating model 1.0.7.9",["data"])
    test.write_to_log()
