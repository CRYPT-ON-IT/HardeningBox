#! /usr/bin/env python3

import os
import sys
import pandas as pd
from Errors import throw
from file_functions import FileFunctions
from update_main_csv import UpdateMainCsv, policy_subdivision
from cis_pdf_scrapper import CISPdfScrapper
from excel_wokrbook import ExcelWorkbook


def check_arguments():
    """
        This function will check all arguments given by the user and assign values to variables.
        It permits to a user to not interact with the program (if all arguments are given).
    """
    #chosen_tool = False
    help_args = ['-h', '--help']
    if any(x in help_args for x in sys.argv):
        print("""
        ---------------------------- HELP MENU -----------------------------
        
            Tools :
                -a, --audit-result : Add audit result to another csv
                    You should add -of or --original-file to specify the original file
                    You should add -af or --adding-file to specify the adding file
                    Usage : 
                        ./main.py --audit-result --original-file <file.csv> --adding-file <file.csv>
                        ./main.py -a -of <file.csv> -af file.csv

                -l, --msft-link : Add Microsoft policy column to a csv
                    You should add -of or --original-file to specify the original file
                    Usage : 
                        ./main.py --msft-link --original-file <file.csv>
                        ./main.py -l -of <file.csv>

                -s, --scrap : Scrap policies from a CIS Benchmark (pdf)
                    You should add -pdf or --pdf-to-txt to specify the PDF2TXT file
                    You should add -o or --output to specify the output filepath
                    Usage : 
                        ./main.py --scrap --pdf-to-txt <file.txt> --output <file.csv>
                        ./main.py -s -pdf <file.pdf> -o <file.csv>

                -as, --add-scrapped : Add scrapped data to a csv file
                    You should add -of or --original-file to specify the original file
                    You should add -af or --adding-file to specify the adding file
                    Usage : 
                        ./main.py --add-scrapped --original-file <file.csv> --adding-file <file.csv>
                        ./main.py -as -of <file.pdf> -af <file.csv>

                -x, --xlsx : Convert CSV file and Excel File
                    You should add --csv2xlsx to transform a csv in an Excel file
                    or You should add --xlsx2csv to transform an Excel in a csv file
                    Usage : 
                        ./main.py -x --csv2xlsx --csv-file <file.csv> --output <file.xlsx>
                        ./main.py -x --xlsx2csv --xlsx-file <file.xlsx> --output <file.csv>

                -p, --pptx : Transform a csv file into PowerPoint slides
                    You should add -csv or --csv-file to specify the csv file
                    You should add -o or --output to specify the saved file location
                    Usage : 
                        ./main.py --pptx --csv-file <file.csv> --output <file.pptx>
                        ./main.py --pptx -csv <file.csv> -o <file.pptx>

                -m, --merge-2-csv : Merge 2 csv files and remove duplicates by "Names"
                    You must add -f1 or --first-file to specify the first csv file
                    You must add -f2 or --second-file to specify the second csv file
                    You should add -o or --output to specify the saved file location
                    Usage : 
                        ./main.py -m --first-file <file1.csv> --second-file <file2.csv>
                        ./main.py --merge-2-csv --first-file <file1.csv> --second-file <file2.csv> --output <output.csv>

                -r, --rm-defaults-values : Replace all default values with "-NODATA-"
                    You must add -f or --input-file to specify the csv file finding list
                    You must add -o or --ouput-file to specify the name of the output csv file
                    Usage :
                        ./main.py -r -f <file.csv> -o <ouput.csv>

                -cx, --csv2report : transform csv files into an Excel report file
                    You must add -c or --client-name to specify your client name
                    You must add -cn or --contexts-names to specify contexts names, separated by a comma
                    You must add -cc or --contexts-configurations to specify context contexts configurations, separated by a comma
                    You must add -cl or --contexts-logs to specify contexts logs, separated by a comma
                    You must add -cf or --contexts-finding-lists to specify finding list of given context, separated by a comma
                    You must add -ap or --all-policies to sepecify the all policies file
                    You must add -o or --output to specify the output excel path 
                    Usage :
                        ./main.py --csv2report --client-name 'CRYPT.ON IT' --contexts-name Administrator,Standard --contexts-configurations /path/to/conf1,/path/to/conf2
                        --contexts-logs /path/to/log1,/path/to/log2 --contexts-finding-lists /path/to/fl1,/path/to/fl2 --all-policies /path/to/all_policies --output here.xlsx
                    Hint :
                        --contexts-name, --contexts-configurations, --contexts-logs and --contexts-finding-lists might have the same number of attribute

                -xc, --report2csv : Transfrom a report file into multiple csv to apply with HardeningKitty
                    You must add -xf or --xlsx-file to specify the Excel report file path
                    You must add -f or --finding-lists to specify finding list linked to every context
                    You must add -ls or --lot-size to specify the max number of policies to have in a file
                    You can add -rf or --registry-filtered to specify that the output should be filtered with Registry method
                    You can add -nrf or --not-registry-filtered to specify that the output shoould not be filtered by method
                    If you have multiple contexts, you have to specify each finding list for the contexts, separated by a comma
                        ./main.py -report2csv --report-file report.xlsx --finding-lists finding_list_1.csv,finding_list_2.csv

            Others :
                -h, --help : show this help menu
                Usage :
                    ./main.py --help
                    ./main.py -h
        --------------------------------------------------------------------
        """)
        throw('Help menu invoked !', 'low')

    audit_result_args = ['-a', '--audit-result']
    if any(x in audit_result_args for x in sys.argv):
        chosen_tool = '1'
        return chosen_tool

    msft_link_args = ['-l', '--msft-link']
    if any(x in msft_link_args for x in sys.argv):
        chosen_tool = '2'
        return chosen_tool

    scrap_args = ['-s', '--scrap']
    if any(x in scrap_args for x in sys.argv):
        chosen_tool = '3'
        return chosen_tool

    add_scrapped_args = ['-as', '--add-scrapped']
    if any(x in add_scrapped_args for x in sys.argv):
        chosen_tool = '4'
        return chosen_tool

    xlsx_args = ['-x', '--xlsx']
    if any(x in xlsx_args for x in sys.argv):
        chosen_tool = '5'
        return chosen_tool

    pptx_args = ['-p', '--pptx']
    if any(x in pptx_args for x in sys.argv):
        chosen_tool = '6'
        return chosen_tool

    mrg_args = ['-m', '--merge-2-csv']
    if any(x in mrg_args for x in sys.argv):
        chosen_tool = '7'
        return chosen_tool

    rm_args = ['-r', '--rm-defaults-values']
    if any(x in rm_args for x in sys.argv):
        chosen_tool = '8'
        return chosen_tool
    
    xc_args = ['-cx', '--csv2report']
    if any(x in xc_args for x in sys.argv):
        chosen_tool = '9'
        return chosen_tool
    
    xc_args = ['-xc', '--report2csv']
    if any(x in xc_args for x in sys.argv):
        chosen_tool = '10'
        return chosen_tool

    chosen_tool = False
    return chosen_tool

print("""
    #################################################################################################################### _ 0 X #
    #                                                                                                                          #
    #   /$$   /$$                           /$$                     /$$                     /$$$$$$$                           #
    #  | $$  | $$                          | $$                    |__/                    | $$__  $$                          #
    #  | $$  | $$  /$$$$$$   /$$$$$$   /$$$$$$$  /$$$$$$  /$$$$$$$  /$$ /$$$$$$$   /$$$$$$ | $$  \ $$  /$$$$$$  /$$   /$$      #
    #  | $$$$$$$$ |____  $$ /$$__  $$ /$$__  $$ /$$__  $$| $$__  $$| $$| $$__  $$ /$$__  $$| $$$$$$$  /$$__  $$|  $$ /$$/      #
    #  | $$__  $$  /$$$$$$$| $$  \__/| $$  | $$| $$$$$$$$| $$  \ $$| $$| $$  \ $$| $$  \ $$| $$__  $$| $$  \ $$ \  $$$$/       #
    #  | $$  | $$ /$$__  $$| $$      | $$  | $$| $$_____/| $$  | $$| $$| $$  | $$| $$  | $$| $$  \ $$| $$  | $$  >$$  $$       #
    #  | $$  | $$|  $$$$$$$| $$      |  $$$$$$$|  $$$$$$$| $$  | $$| $$| $$  | $$|  $$$$$$$| $$$$$$$/|  $$$$$$/ /$$/\  $$      #
    #  |__/  |__/ \_______/|__/       \_______/ \_______/|__/  |__/|__/|__/  |__/ \____  $$|_______/  \______/ |__/  \__/      #
    #                                                                             /$$  \ $$                                    #
    #                                                                            |  $$$$$$/                                    #
    #                                                                             \______/                                     #
    #                                                                                                                          #
    ################################################## By Guillaume de Rybel ###################################################              

    Welcome to the Hardening Box !

    This tool box allows you to use and transform Hardening Data. You will be able to transform CSV extract into PowerPoint slides or Excel tables in easy ways !

    This is based on CIS policies, so it might differ with other organizations.
    
    """)

CHOSEN_TOOL = check_arguments()

if not CHOSEN_TOOL:
    CHOSEN_TOOL = input("""
        1. Add audit result to a CSV file
        2. Add Microsoft Links to CSV (Beta)
        3. Scrap policies from CIS pdf file (https://downloads.cisecurity.org/#/)
        4. Add scrapped data to CSV file
        5. Excel <-> CSV convertion
        6. Transform CSV into PowerPoint slides
        7. Merge 2 csv files and remove duplicates by "Names"
        8. Replace all default values with "-NODATA-"
        9. CSV to Excel Report File
        10. Excel Report File to CSV

    Choose your tool (1->10): """)

# Add audit result to a CSV file
if CHOSEN_TOOL == '1':

    ORIGINAL_FILEPATH = ''
    original_filepath_args = ['-of', '--original-file']
    for original_filepath_arg in original_filepath_args:
        for arg in sys.argv:
            if original_filepath_arg == arg:
                ORIGINAL_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if ORIGINAL_FILEPATH == '':
        ORIGINAL_FILEPATH = input(
            'Which base hardening file should I look for (e.g. : filename.csv) : '
            )
    original_file = FileFunctions(ORIGINAL_FILEPATH)
    original_file.file_exists()
    original_dataframe = original_file.read_csv_file()

    ADDING_FILEPATH = ''
    adding_filepath_args = ['-af', '--adding-file']
    for adding_filepath_arg in adding_filepath_args:
        for arg in sys.argv:
            if adding_filepath_arg == arg:
                ADDING_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if ADDING_FILEPATH == '':
        ADDING_FILEPATH = input("""
        Which audit result file should I look for (e.g. : filename.csv) : 
        """)
    adding_file = FileFunctions(ADDING_FILEPATH)
    adding_file.file_exists()
    adding_dataframe = adding_file.read_csv_file()

    OUTPUT_FILEPATH = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                OUTPUT_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if OUTPUT_FILEPATH == '':
        OUTPUT_FILEPATH = input("""
        How should we name the output file ? :  
        """)

    csv = UpdateMainCsv(
        original_dataframe,
        ORIGINAL_FILEPATH,
        adding_dataframe,
        ADDING_FILEPATH,
        OUTPUT_FILEPATH
    )
    csv.add_audit_result()

    throw('Audit column added successfully.', 'low')

# Add Microsoft Links to CSV (Beta)
elif CHOSEN_TOOL == '2':
    HARDENING_FILEPATH = ''
    hardening_filepath_args = ['-of', '--original-file']
    for hardening_filepath_arg in hardening_filepath_args:
        for arg in sys.argv:
            if hardening_filepath_arg == arg:
                HARDENING_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if HARDENING_FILEPATH == '':
        HARDENING_FILEPATH = input(
            '\nWhich hardening file should I look for (e.g. : filename.csv) : '
        )
    hardening_file = FileFunctions(HARDENING_FILEPATH)
    hardening_file.file_exists()
    hardening_dataframe = hardening_file.read_csv_file()

    csv = UpdateMainCsv(hardening_dataframe, HARDENING_FILEPATH)
    csv.add_microsoft_links()

    throw('Microsoft Link and Possible Values columns added successfully.', 'low')

# Scrap policies from CIS pdf file (https://downloads.cisecurity.org/#/)
elif CHOSEN_TOOL == '3':
    if len(sys.argv) == 0:
        input("""\033[93m
    In order to prepare this tool, you need to transfer pdf text data into a txt file.
    To do that, you need to open your pdf with a pdf reader, and select the whole text (CTRL+A), it might take few seconds, and copy it (CTRL+C).
    When the data is copied, you need to paste it in a file and save it as a txt file.
    
    yes(y) ? : \033[0m""")

    PDF2TXT_FILEPATH = ''
    pdf2txt_filepath_args = ['-pdf', '--pdf-to-txt']
    for pdf2txt_filepath_arg in pdf2txt_filepath_args:
        for arg in sys.argv:
            if pdf2txt_filepath_arg == arg:
                PDF2TXT_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if PDF2TXT_FILEPATH == '':
        PDF2TXT_FILEPATH = input('\nWhich text file should I look for (e.g. : filename.txt) : ')
    pdf2txt_file = FileFunctions(PDF2TXT_FILEPATH)
    pdf2txt_file.file_exists()
    pdf2txt_content = pdf2txt_file.read_file()

    OUTPUT_FILEPATH = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                OUTPUT_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if OUTPUT_FILEPATH == '':
        OUTPUT_FILEPATH = input('Where should we output the result (e.g. : output.csv) : ')

    PDF2TXT = CISPdfScrapper(pdf2txt_content, OUTPUT_FILEPATH)
    PDF2TXT.ScrapPdfData()

    throw('CIS pdf data has been scrapped successfully.', 'low')

# Add scrapped data to CSV file
elif CHOSEN_TOOL == '4':
    ORIGINAL_FILEPATH = ''
    original_filepath_args = ['-of', '--original-file']
    for original_filepath_arg in original_filepath_args:
        for arg in sys.argv:
            if original_filepath_arg == arg:
                ORIGINAL_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if ORIGINAL_FILEPATH == '':
        ORIGINAL_FILEPATH = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    original_file = FileFunctions(ORIGINAL_FILEPATH)
    original_file.file_exists()
    original_dataframe = original_file.read_csv_file()

    ADDING_FILEPATH = ''
    adding_filepath_args = ['-af', '--adding-file']
    for adding_filepath_arg in adding_filepath_args:
        for arg in sys.argv:
            if adding_filepath_arg == arg:
                ADDING_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if ADDING_FILEPATH == '':
        ADDING_FILEPATH = input(
            'Which pdf scrapped data file should I look for (e.g. : filename.csv) : '
        )
    adding_file = FileFunctions(ADDING_FILEPATH)
    adding_file.file_exists()
    adding_dataframe = adding_file.read_csv_file()

    OUTPUT_FILEPATH = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                OUTPUT_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if OUTPUT_FILEPATH == '':
        OUTPUT_FILEPATH = input("""
        How should we name the output file ? :  
        """)

    csv = UpdateMainCsv(
        original_dataframe,
        ORIGINAL_FILEPATH,
        adding_dataframe,
        ADDING_FILEPATH,
        OUTPUT_FILEPATH
    )
    csv.add_scrapped_data_to_csv()

    throw('Scrapped data added successfully.', 'low')

# Excel <-> CSV convertion
elif CHOSEN_TOOL == '5':
    CHOICE = ''
    if '--csv2xlsx' in sys.argv:
        CHOICE = '1'
    elif '--xlsx2csv' in sys.argv:
        CHOICE = '2'

    if CHOICE == '':
        CHOICE = input('''
Would you like to :

1. Convert a Csv file to an Excel file 
2. Convert an Excel file to a csv file

(1 or 2) : 
''')

    if CHOICE == '1':
        CSV_FILEPATH = ''
        csv_filepath_args = ['-csv', '--csv-file']
        for csv_filepath_arg in csv_filepath_args:
            for arg in sys.argv:
                if csv_filepath_arg == arg:
                    CSV_FILEPATH = sys.argv[sys.argv.index(arg)+1]
        if CSV_FILEPATH == '':
            CSV_FILEPATH = input('\nCsv file location : ')
        csv_file = FileFunctions(CSV_FILEPATH)
        csv_file.file_exists()
        csv_file.convert_csv_2_excel()

    elif CHOICE == '2':
        EXCEL_FILEPATH = ''
        excel_filepath_args = ['-xlsx', '--xlsx-file']
        for excel_filepath_arg in excel_filepath_args:
            for arg in sys.argv:
                if excel_filepath_arg == arg:
                    EXCEL_FILEPATH = sys.argv[sys.argv.index(arg)+1]
        if EXCEL_FILEPATH == '':
            EXCEL_FILEPATH = input('\nExcel file location : ')
        excel_file = FileFunctions(EXCEL_FILEPATH)
        excel_file.file_exists()
        excel_file.convert_excel_2_csv()

    else:
        throw('Wrong choice, exiting.', 'high')
    throw("File has been converted successfully.", "low")

# Transform CSV into PowerPoint slides
elif CHOSEN_TOOL == '6':
    HARDENING_FILEPATH = ''
    hardening_filepath_args = ['-csv', '--csv-file']
    for hardening_filepath_arg in hardening_filepath_args:
        for arg in sys.argv:
            if hardening_filepath_arg == arg:
                HARDENING_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if HARDENING_FILEPATH == '':
        HARDENING_FILEPATH = input("""
        Which base hardening file should I look for (e.g. : filename.csv) :
        """)
    hardening_file = FileFunctions(HARDENING_FILEPATH)
    hardening_file.file_exists()
    hardening_dataframe = hardening_file.read_csv_file()

    POWERPOINT_FILEPATH = ''
    powerpoint_filepath_args = ['-o', '--output']
    for powerpoint_filepath_arg in powerpoint_filepath_args:
        for arg in sys.argv:
            if powerpoint_filepath_arg == arg:
                POWERPOINT_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if POWERPOINT_FILEPATH == '':
        POWERPOINT_FILEPATH = input("""
        Where should I output the PowerPoint (e.g. : filename.pptx) : 
        """)

    CONTEXT = None
    contexts = []
    context_columns = []
    print("""\033[93m
Actual Columns : 

    • Name
    • Level (Empty if column does not exists)
    • Severity
    • PossibleValues (Empty if column does not exists)
    • DefaultValue
    • RecomendedValue
    • Description (Empty if column does not exists)
    • MicrosoftLink (Empty if column does not exists)\033[0m""")
    while CONTEXT != '':
        CONTEXT = input("\nIf there is any other column you would like to add, enter the name : ")
        if CONTEXT == '':
            break
        elif CONTEXT in hardening_dataframe.columns:
            contexts.append(CONTEXT)
            context_name = input('Please enter the name to show in the slides : ')
            context_columns.append(context_name)
        else:
            throw('Column not found in CSV, exiting.', 'high')

    hardening_file.create_powerpoint(
        hardening_dataframe, contexts, context_columns, POWERPOINT_FILEPATH)
    throw('PowerPoint has been successfully created.', 'low')

# Add scrapped data to CSV file
elif CHOSEN_TOOL == '7':
    FIRST_FILEPATH = ''
    first_filepath_args = ['-f1', '--first-file']
    for first_filepath_arg in first_filepath_args:
        for arg in sys.argv:
            if first_filepath_arg == arg:
                FIRST_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if FIRST_FILEPATH == '':
        FIRST_FILEPATH = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    first_file = FileFunctions(FIRST_FILEPATH)
    first_file.file_exists()
    first_dataframe = first_file.read_csv_file()

    SECOND_FILEPATH = ''
    second_filepath_args = ['-f2', '--second-file']
    for second_filepath_arg in second_filepath_args:
        for arg in sys.argv:
            if second_filepath_arg == arg:
                SECOND_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if SECOND_FILEPATH == '':
        SECOND_FILEPATH = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    second_file = FileFunctions(SECOND_FILEPATH)
    second_file.file_exists()
    second_dataframe = second_file.read_csv_file()

    OUTPUT_FILEPATH = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                OUTPUT_FILEPATH = sys.argv[sys.argv.index(arg)+1]
    if OUTPUT_FILEPATH == '':
        OUTPUT_FILEPATH = input('Where should we output the result (e.g. : output.csv) : ')

    csv = UpdateMainCsv(
        first_dataframe,
        FIRST_FILEPATH,
        second_dataframe,
        SECOND_FILEPATH,
        OUTPUT_FILEPATH
    )
    csv.merge_two_csv()

    throw('Scrapped data added successfully.', 'low')

# Replace all default values with "-NODATA-"
elif CHOSEN_TOOL == '8':
    # input file
    FILE_FINDING_LIST_PATH = ''
    file_finding_list_path_args = ['-f', '--input-file']
    for file_finding_list_path_arg in file_finding_list_path_args:
        for arg in sys.argv:
            if file_finding_list_path_arg == arg:
                FILE_FINDING_LIST_PATH = sys.argv[sys.argv.index(arg)+1]
    if FILE_FINDING_LIST_PATH == '':
        FILE_FINDING_LIST_PATH = input("""
Which file_finding_list file should I look for (e.g. : filename.csv) : """)

    # output file
    OUTPUT_CSV = ''
    output_csv_args = ['-o', '--output-file']
    for output_csv_arg in output_csv_args:
        for arg in sys.argv:
            if output_csv_arg == arg:
                OUTPUT_CSV = sys.argv[sys.argv.index(arg)+1]
    if OUTPUT_CSV == '':
        OUTPUT_CSV = input("\nWhat's the name of the CSV output file ? : ")

    file_finding_list_file = FileFunctions(FILE_FINDING_LIST_PATH)
    file_finding_list_file.file_exists()
    NEW_FILE_FINDING_LIST = file_finding_list_file.replace_defaults_values(OUTPUT_CSV)

    throw('Microsoft Link and Possible Values columns added successfully.', 'low')

# CSV to Excel Report File
elif CHOSEN_TOOL == '9':
    CLIENT_NAME = ''
    client_name_args = ['-c', '--client-name']
    for client_name_arg in client_name_args:
        for arg in sys.argv:
            if client_name_arg == arg:
                CLIENT_NAME = sys.argv[sys.argv.index(arg)+1]
    if CLIENT_NAME == '':
        CLIENT_NAME = input('Enter the name of your client : ')

    CONTEXTS_NAMES = ''
    contexts_names_args = ['-cn', '--contexts-names']
    for contexts_names_arg in contexts_names_args:
        for arg in sys.argv:
            if contexts_names_arg == arg:
                CONTEXTS_NAMES = sys.argv[sys.argv.index(arg)+1]
    CONTEXTS_NAMES = CONTEXTS_NAMES.split(',')

    CONTEXTS_CONFIGURATIONS = ''
    contexts_configurations_args = ['-cn', '--contexts-configurations']
    for contexts_configurations_arg in contexts_configurations_args:
        for arg in sys.argv:
            if contexts_configurations_arg == arg:
                CONTEXTS_CONFIGURATIONS = sys.argv[sys.argv.index(arg)+1]
    CONTEXTS_CONFIGURATIONS = CONTEXTS_CONFIGURATIONS.split(',')

    CONTEXTS_LOGS = ''
    contexts_logs_args = ['-cl', '--contexts-logs']
    for contexts_logs_arg in contexts_logs_args:
        for arg in sys.argv:
            if contexts_logs_arg == arg:
                CONTEXTS_LOGS = sys.argv[sys.argv.index(arg)+1]
    CONTEXTS_LOGS = CONTEXTS_LOGS.split(',')

    CONTEXTS_FINDING_LISTS = ''
    contexts_finding_lists_args = ['-cf', '--contexts-finding-lists']
    for contexts_finding_lists_arg in contexts_finding_lists_args:
        for arg in sys.argv:
            if contexts_finding_lists_arg == arg:
                CONTEXTS_FINDING_LISTS = sys.argv[sys.argv.index(arg)+1]
    CONTEXTS_FINDING_LISTS = CONTEXTS_FINDING_LISTS.split(',')

    CONTEXTS = []

    if CONTEXTS_NAMES != [''] and len(CONTEXTS_NAMES) == len(CONTEXTS_CONFIGURATIONS) == len(CONTEXTS_LOGS) == len(CONTEXTS_FINDING_LISTS):
        CONTINUE = False
        for index, value in enumerate(CONTEXTS_NAMES):
            extract_file = FileFunctions(CONTEXTS_CONFIGURATIONS[index])
            extract_file.file_exists()
            extract_content = extract_file.read_csv_file()

            log_file = FileFunctions(CONTEXTS_LOGS[index])
            log_file.file_exists()
            log_content = log_file.read_log_file()

            finding_list_file = FileFunctions(CONTEXTS_FINDING_LISTS[index])
            finding_list_file.file_exists()
            finding_list_content = finding_list_file.read_csv_file()

            CONTEXTS.append({
                'Name' : value,
                'Extract' : extract_content,
                'Log' : log_content,
                'FindingList' : finding_list_content
            })

    else:
        CONTINUE = True

    context_i = 1
    while CONTINUE:
        # Ask for context name
        name_of_context = input(f'\nName of the context {context_i} : ')
        # Ask for extract path
        extract_path = input('Path to the configuration extract : ')
        extract_file = FileFunctions(extract_path)
        extract_file.file_exists()
        extract_content = extract_file.read_csv_file()
        # Ask for log path
        log_path = input('Path to the hardening log file : ')
        log_file = FileFunctions(log_path)
        log_file.file_exists()
        log_content = log_file.read_log_file()
        # Ask for finding list
        finding_list_path = input('Path to the finding list corresponding to the context : ')
        finding_list_file = FileFunctions(finding_list_path)
        finding_list_file.file_exists()
        finding_list_content = finding_list_file.read_csv_file()
        CONTEXTS.append({
            'Name' : name_of_context,
            'Extract' : extract_content,
            'Log' : log_content,
            'FindingList' : finding_list_content
        })
        ask = input('Would you like to add another context ? (y/n) : ')
        if ask not in ['y', 'Y']:
            CONTINUE = False
        context_i+=1

    all_policies_path = ''
    all_policies_path_args = ['-ap', '--all-policies']
    for all_policies_path_arg in all_policies_path_args:
        for arg in sys.argv:
            if all_policies_path_arg == arg:
                all_policies_path = sys.argv[sys.argv.index(arg)+1]
    if all_policies_path == '':
        all_policies_path = input('Path to all policies file, a merge of every single (can be created with tool 7) : ')
    all_policies_file = FileFunctions(all_policies_path)
    all_policies_file.file_exists()
    all_policies_content = all_policies_file.read_csv_file()

    xlsx_name = ''
    xlsx_name_args = ['-o', '--output']
    for xlsx_name_arg in xlsx_name_args:
        for arg in sys.argv:
            if xlsx_name_arg == arg:
                xlsx_name = sys.argv[sys.argv.index(arg)+1]
    if xlsx_name == '':
        xlsx_name = input('What is the name of the output Excel file : ')
    
    xlsx_file = ExcelWorkbook(xlsx_name, CONTEXTS, all_policies_content)

    throw('Successfully generated report file', 'low')

# Excel Report File to CSV
elif CHOSEN_TOOL == '10':
    # report file
    REPORT_PATH = ''
    report_file_path_args = ['-xf', '--xlsx-file']
    for report_file_path_arg in report_file_path_args:
        for arg in sys.argv:
            if report_file_path_arg == arg:
                REPORT_PATH = sys.argv[sys.argv.index(arg)+1]
    if REPORT_PATH == '':
        REPORT_PATH = input('\nPlease enter the excel report path : ')
    report_file = FileFunctions(REPORT_PATH)
    report_file.file_exists()
    report_contexts = report_file.read_xlsx_contexts_sheet()

    # registry filter
    registry_filtered = None
    registry_filtered_args = ['-rf', '--registry-filtered']
    for registry_filtered_arg in registry_filtered_args:
        if registry_filtered_arg in sys.argv:
            registry_filtered = True
    not_registry_filtered_args = ['-nrf', '--not-registry-filtered']
    for not_registry_filtered_arg in not_registry_filtered_args:
        if not_registry_filtered_arg in sys.argv:
            registry_filtered = False
    if registry_filtered is None:    
        registry_filtered = input('Should the file be separated by method (Registry | Else) ? This could be useful when applying through GPO. (y/n) : ')
        if registry_filtered.lower() == 'y' or registry_filtered.lower() == 'o':
            registry_filtered = True
        else:
            registry_filtered = False

    NUMBER_OF_CONTEXTS = report_file.get_number_of_context()

    # finding lists
    CONTEXTS_LIST = []
    context_finding_lists_args = ['-f', '--finding-lists']
    for context_finding_lists_arg in context_finding_lists_args:
        for arg in sys.argv:
            if context_finding_lists_arg == arg:
                CONTEXT_FINDING_LISTS = sys.argv[sys.argv.index(arg)+1].split(',')
                if len(CONTEXT_FINDING_LISTS) != NUMBER_OF_CONTEXTS:
                    throw(f'Error : {NUMBER_OF_CONTEXTS} contexts were found in excel file but {len(CONTEXT_FINDING_LISTS)} finding lists were given.', 'high')
                else:
                    CONTEXT = 1
                    for FINDING_LIST in CONTEXT_FINDING_LISTS:
                        context_file = FileFunctions(FINDING_LIST)
                        context_file.file_exists()
                        context_df = context_file.read_csv_file()

                        CONTEXTS_LIST.append({
                            'ContextName' : f'Context{CONTEXT}',
                            'ContextDataframe' : context_df
                        })
                        CONTEXT += 1
    
    if CONTEXTS_LIST == []:
        for CONTEXT in range(NUMBER_OF_CONTEXTS):
            CONTEXT_FINDING_LIST = input(f'\nPlease enter the path of the finding list for context {CONTEXT + 1} : ')
            context_file = FileFunctions(CONTEXT_FINDING_LIST)
            context_file.file_exists()
            context_df = context_file.read_csv_file()

            CONTEXTS_LIST.append({
                'ContextName' : f'Context{CONTEXT + 1}',
                'ContextDataframe' : context_df
            })

    # lot size
    LOT_SIZE = None
    lot_size_args = ['-ls', '--lot-size']
    for lot_size_arg in lot_size_args:
        for arg in sys.argv:
            if lot_size_arg == arg:
                LOT_SIZE = sys.argv[sys.argv.index(arg)+1]
    if LOT_SIZE is None:
        LOT_SIZE = input('\nPlease enter the lot size (default is 10) : ')
    
    if LOT_SIZE == '':
        LOT_SIZE = 10
    
    try:
        LOT_SIZE = int(LOT_SIZE)
    except ValueError:
        throw('The lot size given is not an integer.', 'high')

    for CONTEXT in range(NUMBER_OF_CONTEXTS):
        CONTEXT_FINDING_LIST = input(f'\nPlease enter the path of the finding list for context {CONTEXT + 1} : ')
        context_file = FileFunctions(CONTEXT_FINDING_LIST)
        context_file.file_exists()
        context_df = context_file.read_csv_file()

        CONTEXTS_LIST.append({
            'ContextName' : f'Context{CONTEXT}',
            'ContextDataframe' : context_df
        })
        CONTEXT += 1
    
    if CONTEXTS_LIST == []:
        for CONTEXT in range(NUMBER_OF_CONTEXTS):
            CONTEXT_FINDING_LIST = input(f'\nPlease enter the path of the finding list for context {CONTEXT + 1} : ')
            context_file = FileFunctions(CONTEXT_FINDING_LIST)
            context_file.file_exists()
            context_df = context_file.read_csv_file()

            CONTEXTS_LIST.append({
                'ContextName' : f'Context{CONTEXT + 1}',
                'ContextDataframe' : context_df
            })

    parent_path = "./hardening_policies/"
    if not os.path.exists(parent_path):
        os.mkdir(parent_path)

    ### Create Global Hardening Files

    for CONTEXT in CONTEXTS_LIST:
        column_name_result = CONTEXT['ContextName'] + ' - ComputedResult'
        column_name_value = CONTEXT['ContextName'] + ' - Computed Value'
        column_name_fixed_value = CONTEXT['ContextName'] + ' - Fixed Value'

        choosed_policies = report_contexts.loc[(report_contexts['Ateliers'].str.startswith("Atelier")) & (report_contexts[column_name_value] != 'to check') & (report_contexts[column_name_value] != "N/A") & (report_contexts[column_name_fixed_value] != "_")]

        del CONTEXT['ContextDataframe']['RecommendedValue']

        new_file_finding_list = CONTEXT['ContextDataframe'].merge(choosed_policies[['Name',column_name_value]], on=['Name'])
        new_file_finding_list = new_file_finding_list.rename(columns={column_name_value: "RecommendedValue"})

        if registry_filtered:
            new_file_finding_list_registry = new_file_finding_list.loc[(new_file_finding_list["Method"] == "Registry")]
            new_file_finding_list_registry.to_csv(path_or_buf=parent_path + 'Registry_Based_Policies_' + CONTEXT['ContextName'] + '.csv',index=False)
            new_file_finding_list_no_registry = new_file_finding_list.loc[(new_file_finding_list["Method"] != "Registry")]
            new_file_finding_list_no_registry.to_csv(path_or_buf=parent_path + 'No_Registry_Based_Policies_' + CONTEXT['ContextName'] + '.csv',index=False)
        else:
            new_file_finding_list.to_csv(path_or_buf=parent_path + CONTEXT['ContextName'] + '.csv',index=False)

        ### Create Hardening Files By Workshop

        cpt = 0
        workshops = report_contexts["Ateliers"].unique()
        for workshop in workshops:
            byworkshop_choosed_policies = choosed_policies.loc[(report_contexts['Ateliers'] == workshop)]
            ### For each category
            categories = byworkshop_choosed_policies["Category"].unique()
            for category in categories:
                new_filtered_excel_file = byworkshop_choosed_policies.loc[(report_contexts['Category'] == category)]
                new_file_finding_list = pd.merge(CONTEXT['ContextDataframe'], new_filtered_excel_file[['Name', column_name_value]], on=['Name'])
                new_file_finding_list = new_file_finding_list.rename(columns={column_name_value: "RecommendedValue"})
                category = category.replace(":", "-")
                
                bycontext_path = f"{parent_path}{CONTEXT['ContextName']}/"
                if not os.path.exists(bycontext_path):
                    os.mkdir(bycontext_path)
                byworkshop_path = f"{bycontext_path}{workshop}/"
                if not os.path.exists(byworkshop_path):
                    os.mkdir(byworkshop_path)
                bycategory_path = f"{byworkshop_path}{category}/"
                if not os.path.exists(bycategory_path):
                    os.mkdir(bycategory_path)

                if registry_filtered:
                    base_name = bycategory_path + 'Registry_Based_Policies_' + CONTEXT['ContextName'] + "_" + workshop + "_" + category
                    new_file_finding_list_registry = new_file_finding_list.loc[(new_file_finding_list["Method"] == "Registry")]
                    policy_subdivision(new_file_finding_list_registry, base_name, LOT_SIZE)
                    base_name = bycategory_path + 'No_Registry_Based_Policies_' + CONTEXT['ContextName'] + "_" + workshop + "_" + category
                    new_file_finding_list_no_registry = new_file_finding_list.loc[(new_file_finding_list["Method"]!= "Registry")]
                    policy_subdivision(new_file_finding_list_no_registry, base_name, LOT_SIZE)
                else:
                    base_name = bycategory_path + CONTEXT['ContextName'] + "_" + workshop + "_" + category
                    policy_subdivision(new_file_finding_list, base_name, LOT_SIZE)

    throw(f'Output was saved in \'{parent_path}\' folder.', 'low')
else:
    throw('Tool selected not in list, exiting.', 'high')
