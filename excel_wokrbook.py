import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from Errors import throw

FONT_COLOR = Font(color="FFFFFF")
WRAP_TEXT = Alignment(vertical='top', wrapText=True)

class ExcelWorkbook:
    def __init__(self, path: str, contexts: list[dict], all_policies_content: pd.DataFrame) -> None:
        self.path = path
        self.contexts = contexts
        self.all_policies_content = all_policies_content

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

        self.all_policies_content['Operator'] = self.all_policies_content['Operator'].replace('=|0','=CONCATENATE("=|0")')

        self.workbook = Workbook()
        self.save()
        self.workbook = load_workbook(self.path)
        self.create_dashboard_sheet()
        self.create_contexts_sheet()
        self.create_all_policies_sheet()
        self.create_contexts_sheets()
        self.save()

        self.append_all_policies()
        self.append_contexts_data()
        

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
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-extract')
            worksheet.sheet_properties.tabColor = self.colors_pairs[colors_index]['hex1']
            sheet_index+=1
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-log-reported')
            worksheet.sheet_properties.tabColor = self.colors_pairs[colors_index]['hex2']
            sheet_index+=1
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-scrapped-data')
            sheet_index+=1
            worksheet = self.workbook.create_sheet(index=sheet_index, title=f'{context_number}-{context["Name"]}-finding-list')
            sheet_index+=1

            context_number+=1

            colors_index+=1
            # restart color if all colors where used
            if colors_index > colors_len:
                colors_index = 0
                
    def append_all_policies(self):
        """This will add all_policies data to corresponding sheet
        """
        # Add all_policies data
        with pd.ExcelWriter(self.path,
                        mode='a',
                        engine='openpyxl',
                        if_sheet_exists='overlay'
        ) as writer:
            self.all_policies_content.to_excel(writer, sheet_name='All-Policies', index=False)

        # Format all_policies sheet
        self.workbook = load_workbook(self.path)
        worksheet = self.workbook['All-Policies']
        table = Table(displayName='All_policies',
                    ref=f'A1:{get_column_letter(self.all_policies_content.shape[1])}{len(self.all_policies_content)+1}')
        style = TableStyleInfo(name="TableStyleMedium11", showFirstColumn=False,
                        showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        table.tableStyleInfo = style
        worksheet.add_table(table)
        for cells in worksheet["1:1"]:
            cells.font = FONT_COLOR
        for columns in range(worksheet.min_column, worksheet.max_column + 1):
            worksheet.column_dimensions[get_column_letter(columns)].width = 25
        self.save()

    def append_contexts_data(self):
        """This will add contexts data to corresponding sheet
        """
        context_number = 1
        for context in self.contexts:
            # Append LOG
            with pd.ExcelWriter(self.path,
                        mode='a',
                        engine='openpyxl',
                        if_sheet_exists='overlay'
            ) as writer:
                context['Log'].to_excel(writer, sheet_name=f'{context_number}-{context["Name"]}-log-reported', index=False)

            # Format LOG
            self.workbook = load_workbook(self.path)
            worksheet = self.workbook[f'{context_number}-{context["Name"]}-log-reported']
            table = Table(displayName=f'context{context_number}_log', ref=f'A1:{get_column_letter(context["Log"].shape[1])}{len(context["Log"])+1}')
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            table.tableStyleInfo = style
            worksheet.add_table(table)
            for cells in worksheet["1:1"]:
                cells.font = FONT_COLOR
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = WRAP_TEXT
            for columns in range(worksheet.min_column, worksheet.max_column + 1):
                worksheet.column_dimensions[get_column_letter(columns)].width = 100
            self.save()

            # Append FINDING_LIST
            with pd.ExcelWriter(self.path,
                        mode='a',
                        engine='openpyxl',
                        if_sheet_exists='replace'
            ) as writer:
                context['FindingList'].to_excel(writer, sheet_name=f'{context_number}-{context["Name"]}-finding-list', index=False)
            
            # Format FINDING_LIST
            self.workbook = load_workbook(self.path)
            worksheet = self.workbook[f'{context_number}-{context["Name"]}-finding-list']
            table = Table(displayName=f'{context["Name"]}_finding_list',
                        ref=f'A1:{get_column_letter(context["FindingList"].shape[1])}{len(context["FindingList"])+1}')
            style = TableStyleInfo(name="TableStyleMedium11", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            table.tableStyleInfo = style
            worksheet.add_table(table)
            for cells in worksheet["1:1"]:
                cells.font = FONT_COLOR
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = WRAP_TEXT
            for columns in range(worksheet.min_column, worksheet.max_column + 1):
                worksheet.column_dimensions[get_column_letter(columns)].width = 50
            self.save()

            # Append EXTRACT
            with pd.ExcelWriter(self.path,
                        mode='a',
                        engine='openpyxl',
                        if_sheet_exists='overlay'
            ) as writer:
                context['Extract'].to_excel(writer, sheet_name=f'{context_number}-{context["Name"]}-extract', index=False)

            # Format EXTRACT
            self.workbook = load_workbook(self.path)
            worksheet = self.workbook[f'{context_number}-{context["Name"]}-extract']
            table = Table(displayName=f'{context["Name"]}_reported_file_finding_list',
                        ref=f'A1:{get_column_letter(context["Extract"].shape[1])}{len(context["Extract"])+1}')
            style = TableStyleInfo(name="TableStyleMedium8", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            table.tableStyleInfo = style
            worksheet.add_table(table)
            for cells in worksheet["1:1"]:
                cells.font = FONT_COLOR
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = WRAP_TEXT
            for columns in range(worksheet.min_column, worksheet.max_column + 1):
                worksheet.column_dimensions[get_column_letter(columns)].width = 50
            self.save()



            context_number+=1