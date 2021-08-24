from openpyxl import Workbook, load_workbook
import sys
import os

class WriteToLog:
    def __init__(self,filepath,filename,data):
        '''
        Input : filepath (path to file- filename)
                filename (name of file)
                data (dictionary of proposal log {est number : [details, ..., ...]})
        Output : NA
        '''
        self.filepath = filepath
        self.filename = filename
        self.data = data

    def write_to_log(self):
        #sys.path.append(self.filepath)
        #os.chdir(self.filepath)
        workbook = load_workbook(filename=self.filename, data_only=True)
        proposal_log = workbook.active
        estimate_number = [*self.data][0]
        print(estimate_number)
        data_values = [value for value in self.data.values()]
        print(self.data.values())
        print(data_values[0])
        proposal_log.append(data_values[0])
        workbook.save(self.filename)


def _testing():
    test = WriteToLog("C:/Users/Owner/Desktop/Estimating model 1.0.7.9","Proposal_log.xlsx",{1234: ["this","is","a","test"]})
    test.write_to_log()
