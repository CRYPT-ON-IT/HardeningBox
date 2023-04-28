import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from Errors import throw

FONT_COLOR = Font(color="FFFFFF")
WRAP_TEXT = Alignment(vertical='top', wrapText=True)
FONT_COLOR_HEADER = Font(color="000000", bold=True)
BORDER_RIGHT = Border(right=Side(style="thick", color="000000"))
BORDER_TITTLE = Border(bottom=Side(style="thick", color="000000"),
                       right=Side(style="thick", color="000000"))

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
        self.construct_global_information_header()
        self.construct_global_information_rows()
        self.construct_contexts_columns()
        self.construct_contexts_table()
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
        """Create different contexts sheets
        """
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
            table = Table(displayName=f'finding_list_context{context_number}',
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
            table = Table(displayName=f'context{context_number}_reported_file_finding_list',
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

    def construct_global_information_header(self):
        """
        Table Header Global Information
        """
        worksheet = self.workbook['Contexts']
        worksheet['A1'] = 'Global Information'
        worksheet['A1'].alignment = Alignment(horizontal="center")
        worksheet['A1'].font = FONT_COLOR_HEADER
        worksheet['A1'].border = BORDER_TITTLE
        worksheet.merge_cells('A1:H1')
        worksheet['A2'] = 'ID'
        worksheet['B2'] = 'Line'
        worksheet['C2'] = 'Category'
        worksheet['D2'] = 'Name'
        worksheet['E2'] = 'CIS Level'
        worksheet['F2'] = 'Severity from original file'
        worksheet['G2'] = 'Default Value'
        worksheet['H2'] = 'Recommended Value'
        worksheet['H2'].border = BORDER_RIGHT

    def construct_global_information_rows(self):
        """
        Create all columns (ID, Line, etc.) for the global information
        """
        worksheet_all_policies = self.workbook['All-Policies']
        worksheet_contexts = self.workbook['Contexts']
        row_all_policies = worksheet_all_policies.max_row - 1
        value_columns_global = []
        for row_number in range(row_all_policies):
            value_columns_global.append([
                f"='All-Policies'!A{row_number + 2}",#ID
                "=ROW(INDEX(All_policies[Rationale],MATCH([[#This Row],[Name]],All_policies[Name],0)))-1",#Line
                "=INDEX(All_policies[Category],MATCH([[#This Row],[Name]],All_policies[Name],0))",#Category
                f"='All-Policies'!C{row_number + 2}",#Name
                "=INDEX(All_policies[Level],MATCH([[#This Row],[Name]],All_policies[Name],0))",#Level
                "=INDEX(All_policies[Severity],MATCH([[#This Row],[Name]],All_policies[Name],0))",#Severity
                "=CONCATENATE(INDEX(All_policies[DefaultValue],MATCH([[#This Row],[Name]],All_policies[Name],0)))",#DefaultValue
                "=CONCATENATE(INDEX(All_policies[RecommendedValue],MATCH([[#This Row],[Name]],All_policies[Name],0)))"#RecommendedValue
            ])
            for column_number in range(8):
                worksheet_contexts.cell(row=row_number+3,column=column_number+1).value =\
                value_columns_global[row_number][column_number]
                worksheet_contexts.cell(row=row_number+3,column=column_number+1).font =\
                    Font(size=11, color="000000", bold=True)
                worksheet_contexts.cell(row=row_number+3,column=column_number+1).fill =\
                    PatternFill(fill_type='solid', fgColor='DDDDDD')
                worksheet_contexts.cell(row=row_number+3,column=8).border = BORDER_RIGHT

    def construct_contexts_columns(self):
        """This function creates contexts column and fill them
        """
        worksheet = self.workbook['Contexts']
        context_number = 1
        start_line = 9
        end_line = start_line + 8
        colors_index = 0
        colors_len = len(self.colors_pairs)
        for context in self.contexts:
            self.construct_context_header(worksheet, context['Name'], context_number, start_line, end_line)
            self.construct_context_rows(context_number, start_line, end_line, self.colors_pairs[colors_index]['hex2'])
            context_number+=1
            colors_index+=1
            if context_number <= len(self.contexts):
                start_line+=9
                end_line+=9
            # restart color if all colors where used
            if colors_index > colors_len:
                colors_index = 0
        self.construct_workshops_header(worksheet, end_line+1)
        self.construct_workshops_rows(end_line+1)

    def construct_context_header(self, worksheet, nom_context,
                             num_context, number_column_start, number_column_end):
        """
        Table Header for each context
        """
        letter_column_start = get_column_letter(number_column_start)
        letter_column_end = get_column_letter(number_column_end)
        worksheet[letter_column_start + '1'] = 'Context-' + str(num_context) + '(' + nom_context + ')'
        worksheet[letter_column_start + '1'].alignment = Alignment(horizontal="center")
        worksheet[letter_column_start + '1'].font = FONT_COLOR_HEADER
        worksheet[letter_column_start + '1'].border = BORDER_TITTLE
        worksheet.merge_cells(letter_column_start + '1:' + letter_column_end + '1')
        list_columns = []
        for columns in range(number_column_start, number_column_end + 1):
            column_letter = get_column_letter(columns)
            list_columns.append(column_letter)
        worksheet[list_columns[0] + '2'] = 'ID Context' + str(num_context)
        worksheet[list_columns[1] + '2'] = 'Context' + str(num_context) + ' - Operator'
        worksheet[list_columns[2] + '2'] = 'Context' + str(num_context) + ' - Computed Severity'
        worksheet[list_columns[3] + '2'] = 'Context' + str(num_context) + ' - Manual Severity'
        worksheet[list_columns[4] + '2'] = 'Context' + str(num_context) + ' - Result'
        worksheet[list_columns[5] + '2'] = 'Context' + str(num_context) + ' - ComputedResult'
        worksheet[list_columns[6] + '2'] = 'Context' + str(num_context) + ' - Fixed Value'
        worksheet[list_columns[7] + '2'] = 'Context' + str(num_context) + ' - Computed Value'
        worksheet[list_columns[8] + '2'] = 'Context' + str(num_context) + ' - isRecValue'
        worksheet[list_columns[8] + '2'].border = BORDER_RIGHT

    def construct_context_rows(self, numero_context, first_column, last_column, color_context):
        """
        Create all columns (ID, Result, Severity, etc.) for each context
        """
        worksheet_all_policies = self.workbook['All-Policies']
        worksheet_contexts = self.workbook['Contexts']
        row_all_policies = worksheet_all_policies.max_row - 1
        value_columns_context = []
        for row_number in range(row_all_policies):
            value_columns_context.append([f"=INDEX(finding_list_context{numero_context}[ID],MATCH([Name],finding_list_context{numero_context}[Name],0))",
                                f"=INDEX(finding_list_context{numero_context}[Operator],MATCH([Name],finding_list_context{numero_context}[Name],0))",
                                f'=IF(ISERROR([[#This Row],[ID context{numero_context}]]),"Not Applicable",IF(ISBLANK([[#This Row],[context{numero_context} - Manual Severity]]),IFERROR(_xlfn.SWITCH([[#This Row],[context{numero_context} - Operator]],">=",IF(_xlfn.NUMBERVALUE([[#This Row],[context{numero_context} - ComputedResult]])>=_xlfn.NUMBERVALUE([[#This Row],[Recommended Value]]),"Passed",[[#This Row],[Severity from original file]]),"<=!0",IF(AND(_xlfn.NUMBERVALUE([[#This Row],[context{numero_context} - ComputedResult]])<=_xlfn.NUMBERVALUE([[#This Row],[Recommended Value]]),[[#This Row],[context{numero_context} - ComputedResult]]<>0),"Passed",[[#This Row],[Severity from original file]]),0,IF([[#This Row],[Recommended Value]]=[[#This Row],[context{numero_context} - ComputedResult]],"Passed",[[#This Row],[Severity from original file]]),"!=",IF([[#This Row],[Recommended Value]]<>[[#This Row],[context{numero_context} - ComputedResult]],"Passed",[[#This Row],[Severity from original file]]),"=",IF([[#This Row],[Recommended Value]]=[[#This Row],[context{numero_context} - ComputedResult]],"Passed",[[#This Row],[Severity from original file]]),"<=",IF(_xlfn.NUMBERVALUE([[#This Row],[context{numero_context} - ComputedResult]])<=_xlfn.NUMBERVALUE([[#This Row],[Recommended Value]]),"Passed",[[#This Row],[Severity from original file]]),"contains",IF(SEARCH([[#This Row],[Recommended Value]],[[#This Row],[context{numero_context} - Result]]),"Passed",[[#This Row],[Severity from original file]])),[[#This Row],[Severity from original file]]),[[#This Row],[context{numero_context} - Manual Severity]]))',
                                '',
                                f"=CONCATENATE(INDEX(context{numero_context}_reported_file_finding_list[Result],MATCH([Name],context{numero_context}_reported_file_finding_list[Name],0)))",
                                f'=_xlfn.SWITCH([[#This Row],[context{numero_context} - Result]],"-NODATA-",[[#This Row],[Default Value]],"BUILTIN\Administrateurs","BUILTIN\Administrators",[[#This Row],[context{numero_context} - Result]])',
                                "_",
                                f'=_xlfn.SWITCH([[#This Row],[context{numero_context} - Fixed Value]],"same",[[#This Row],[context{numero_context} - ComputedResult]],"recval",[[#This Row],[Recommended Value]],"to check","to check","_","N/A",[[#This Row],[context{numero_context} - Fixed Value]])',
                                f'=_xlfn.SWITCH([[#This Row],[context{numero_context} - Fixed Value]],"_",[[#This Row],[context{numero_context} - Computed Severity]],"recval","Passed",IFERROR(_xlfn.SWITCH([[#This Row],[context{numero_context} - Operator]],">=",IF(_xlfn.NUMBERVALUE([[#This Row],[context{numero_context} - Fixed Value]])>=_xlfn.NUMBERVALUE([[#This Row],[Recommended Value]]),"Passed",[[#This Row],[context{numero_context} - Computed Severity]]),"<=!0",IF(AND(_xlfn.NUMBERVALUE([[#This Row],[context{numero_context} - Fixed Value]])<=_xlfn.NUMBERVALUE([[#This Row],[Recommended Value]]),[[#This Row],[context{numero_context} - Fixed Value]]<>0),"Passed",[[#This Row],[context{numero_context} - Computed Severity]]),0,IF([[#This Row],[Recommended Value]]=[[#This Row],[context{numero_context} - Fixed Value]],"Passed",[[#This Row],[context{numero_context} - Computed Severity]]),"!=",IF([[#This Row],[Recommended Value]]<>[[#This Row],[context{numero_context} - Fixed Value]],"Passed",[[#This Row],[context{numero_context} - Computed Severity]]),"<=",IF(_xlfn.NUMBERVALUE([[#This Row],[context{numero_context} - Fixed Value]])<=_xlfn.NUMBERVALUE([[#This Row],[Recommended Value]]),"Passed",[[#This Row],[context{numero_context} - Computed Severity]]),"contains",IF(SEARCH([[#This Row],[Recommended Value]],[[#This Row],[context{numero_context} - Fixed Value]]),"Passed",[[#This Row],[context{numero_context} - Computed Severity]])),[[#This Row],[context{numero_context} - Computed Severity]]))'])
            for column_number in range(9):
                worksheet_contexts.cell(row=row_number+3,column=column_number+first_column).value =\
                    value_columns_context[row_number][column_number]
                worksheet_contexts.cell(row=row_number+3,column=column_number+first_column).fill =\
                    PatternFill(fill_type='solid', fgColor=color_context)
                worksheet_contexts.cell(row=row_number+3,column=last_column).border = BORDER_RIGHT

    def construct_workshops_header(self, worksheet, number_column_start):
        """
        Table header for different workshops, commentary etc.
        """
        letter_column_start = get_column_letter(number_column_start)
        worksheet[letter_column_start + '1'] = 'Summary'
        worksheet[letter_column_start + '1'].alignment = Alignment(horizontal="center")
        worksheet[letter_column_start + '1'].font = FONT_COLOR_HEADER
        worksheet[letter_column_start + '1'].border = BORDER_TITTLE
        list_columns = []
        for columns in range(number_column_start, number_column_start + 11):
            column_letter = get_column_letter(columns)
            list_columns.append(column_letter)
        worksheet[list_columns[0] + '2'] = 'Ateliers'
        worksheet[list_columns[1] + '2'] = 'All Passed'
        worksheet[list_columns[2] + '2'] = 'At least one "to check"'

        if len(self.contexts) == 1:
            number_of_columns = 7
        else:
            number_of_columns = 6 + len(self.contexts)
        letter_column_end = get_column_letter(number_column_start + number_of_columns)
        worksheet.merge_cells(letter_column_start + '1:' + letter_column_end + '1')

        context_index = 1
        col = 3
        all_column_name = 'Context1'
        for _ in self.contexts:
            worksheet[list_columns[col] + '2'] = f'Context{context_index}'
            if len(self.contexts) > 1:
                all_column_name+=f' & Context{context_index}'
            context_index+=1
            col+=1
        if len(self.contexts) > 1:
            worksheet[list_columns[col] + '2'] = all_column_name
            col+=1
        worksheet[list_columns[col] + '2'] = 'Actions (Coté Cryptonit)'
        worksheet[list_columns[col+1] + '2'] = 'Actions (Coté Client)'
        worksheet[list_columns[col+2] + '2'] = 'Conclusion'
        worksheet[list_columns[col+2] + '2'].border = BORDER_RIGHT

    def construct_workshops_rows(self, number_column_start):
        """
        Create all columns for the summary
        """
        worksheet_all_policies = self.workbook['All-Policies']
        worksheet_contexts = self.workbook['Contexts']
        row_all_policies = worksheet_all_policies.max_row - 1
        value_columns_global = []
        
        for row_number in range(row_all_policies):
            current_row_formulas = []
            current_row_formulas.append('_')

            formula_1 = '=AND('
            end_formula_1 = ')'

            formula_2 = '=OR('
            end_formula_2 = ')'

            context_index = 1
            for _ in self.contexts:
                formula_1+=f'[[#This Row],[Context{context_index} - Computed Severity]]="Passed"'
                formula_2+=f'[[#This Row],[Context{context_index} - Fixed Value]]="to check"'
                if len(self.contexts) > context_index:
                    formula_1+=','
                    formula_2+=','
                context_index+=1

            formula_1+=end_formula_1
            formula_2+=end_formula_2

            current_row_formulas.append(formula_1)
            current_row_formulas.append(formula_2)

            context_index = 1
            formula_3 = '=AND('
            end_formula_3 = ')'
            for _ in self.contexts:
                current_row_formulas.append(f'=IF(IFERROR([[#This Row],[ID Context{context_index}]],FALSE)=FALSE,FALSE,TRUE)')
                formula_3+=f'[[#This Row],[Context{context_index}]]'
                if len(self.contexts) > context_index:
                    formula_3+=','
                context_index+=1

            formula_3+=end_formula_3
            
            if len(self.contexts) > 1:
                current_row_formulas.append(formula_3)
            
            current_row_formulas+=['','','','']

            value_columns_global.append(current_row_formulas)
            
            default_rng = 6
            if len(self.contexts) > 1:
                default_rng+=len(self.contexts)
                

            for column_number in range(default_rng):
                worksheet_contexts.cell(row=row_number+3,
                                        column=column_number+number_column_start).value =\
                value_columns_global[row_number][column_number]
                worksheet_contexts.cell(row=row_number+3,
                                        column=column_number+number_column_start).font =\
                    Font(size=11, color="000000", bold=True)
                worksheet_contexts.cell(row=row_number+3,
                                        column=column_number+number_column_start).fill =\
                    PatternFill(fill_type='solid', fgColor='F2F2F2')
                worksheet_contexts.cell(row=row_number+3,column=8).border = BORDER_RIGHT

    def construct_contexts_table(self):
        """
        Create the table of contexts sheet
        """
        worksheet_contexts = self.workbook['Contexts']
        table = Table(displayName='table_contexts',
                    ref=f'A2:{get_column_letter(worksheet_contexts.max_column)}{worksheet_contexts.max_row}')
        style = TableStyleInfo(name="TableStyleLight15", showFirstColumn=False, showLastColumn=False,
                            showRowStripes=True, showColumnStripes=False)
        table.tableStyleInfo = style
        worksheet_contexts.add_table(table)
        for header_cells in worksheet_contexts["2:2"]:
            header_cells.font = FONT_COLOR_HEADER
            header_cells.fill = PatternFill(fill_type='solid', fgColor='9F9F9F')
        for columns in range(worksheet_contexts.min_column, worksheet_contexts.max_column + 1):
            worksheet_contexts.column_dimensions[get_column_letter(columns)].width = 15
        for row in worksheet_contexts.iter_rows(min_row=3, max_row=worksheet_contexts.max_row):
            for cell in row:
                cell.alignment = Alignment(vertical='center', wrapText=True)
