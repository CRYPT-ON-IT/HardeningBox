import os
import pandas as pd

class FileFunctions():

    def __init__(self, file):
        self.file = file

    def checkIfFileExistsAndReadable(self):
        if os.path.exists(self.file):
            print('File specified exists !')
        else:
            print('File specified not found.')
            # throw

    def readFile(self):
        try:
            file = open(self.file, 'r')
            text = file.read()
            file.close()
        except:
            print("Couldn't read file.")
            # throw
        
        return text

    def readCsvFile(self):
        df = pd.read_csv(self.file, encoding='latin1')
        df = df.astype(str)
        return df

    def convertCsv2Excel(self):
        df = pd.read_csv(self.file)
        df = df.fillna('')
        output_excel = input("\nWhat's the name of the Excel output file ? : ")
        df.to_excel(output_excel, index=False)

    def convertExcel2Csv(self):
        df = pd.read_excel(self.file)
        df = df.fillna('')
        output_csv = input("\nWhat's the name of the CSV output file ? : ")
        df.to_csv(output_csv, index=False)