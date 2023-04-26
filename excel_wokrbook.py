from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from Errors import throw

class ExcelWorkbook:
    def __init__(self, path: str, contexts: list[dict]) -> None:
        self.path = path
        self.contexts = contexts

        self.colors_pairs = [
            {
                'color' : 'grey',
                'hex1' : 'D9D9D9',
                'hex2' : 'BFBFBF'
            },
            {
                'color' : 'green',
                'hex1' : 'DFF8D6',
                'hex2' : 'BFEFAD'
            },
            {
                'color' : 'blue',
                'hex1' : '99D9FF',
                'hex2' : '66C6FF'
            },
            {
                'color' : 'red',
                'hex1' : 'FC4868',
                'hex2' : 'FF8CA0'
            },
            {
                'color' : 'purple',
                'hex1' : 'AD4CFF',
                'hex2' : 'D19CFD'
            },
            {
                'color' : 'orange',
                'hex1' : 'FFD045',
                'hex2' : 'FDE28F'
            }
        ]

        if not self.path.endswith('.xlsx'):
            self.path += '.xlsx'

        self.workbook = Workbook()
        self.save()
        self.workbook = load_workbook(self.path)
        self.create_dashboard_sheet()
        self.create_contexts_sheet()
        self.create_all_policies_sheet()
        self.create_contexts_sheets()
        self.save()

    def save(self):
        """Saves the excel file
        """
        try:
            self.workbook.save(self.path)
        except:
            throw('An error occured while saving the Excel file, the name might be the cause.', 'high')

    def create_dashboard_sheet(self):
        """Create the dashboard sheet
        """
        sheet = self.workbook.active
        sheet.title = 'Dashboards'
        for row in sheet.iter_rows(max_row=160, max_col=50):
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="FFFFFF")
        worksheet_dashboards_ateliers = self.workbook.create_sheet(index=1, title='Dashboards - Ateliers')
        for row in worksheet_dashboards_ateliers.iter_rows(max_row=55, max_col=71):
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="FFFFFF")
    
    def create_contexts_sheet(self):
        """Create the Contexts sheet
        """
        self.workbook.create_sheet(index=2, title='Contexts')
    
    def create_all_policies_sheet(self):
        """Create the All_Policies sheet
        """
        worksheet = self.workbook.create_sheet(index=3, title='All-Policies')
        worksheet.sheet_properties.tabColor = 'FFDFDB'

    def create_contexts_sheets(self):
        sheet_index = 4
        context_number = 1
        colors_index = 0
        colors_len = len(self.colors_pairs)
        for context in self.contexts:
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-Extract')
            worksheet.sheet_properties.tabColor = self.colors_pairs[colors_index]['hex1']
            sheet_index+=1
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-log-reported')
            worksheet.sheet_properties.tabColor = self.colors_pairs[colors_index]['hex2']
            sheet_index+=1
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-scrapped-data')
            sheet_index+=1

            context_number+=1

            colors_index+=1
            # restart color if all colors where used
            if colors_index > colors_len:
                colors_index = 0
