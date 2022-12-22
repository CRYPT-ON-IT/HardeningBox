#! /usr/bin/env python3

import sys
from Errors import throw
from file_functions import FileFunctions
from update_main_csv import UpdateMainCsv
from cis_pdf_scrapper import CISPdfScrapper


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
                    You should add -pdf or --pdf-to-txt to specify the pdf2txt file
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
                        ./main.py --csv2xlsx --csv-file <file.csv> --output <file.xlsx>
                        ./main.py --xlsx2csv --xlsx-file <file.xlsx> --output <file.csv>

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

                -t, --trace : Convert Excel trace file to CSV applicable per context
                    You must add -tf or --trace-file to specify the Excel trace file
                    Usage :
                        ./main.py -t -tf <trace_file.xlsx>
                        ./main.py --trace --trace-file <trace_file.xlsx>

                -r, --rm-defaults-values : Replace all default values with "-NODATA-"
                    You must add -f or --input-file to specify the csv file finding list
                    You must add -o or --ouput-file to specify the name of the output csv file
                    Usage :
                        ./main.py -r -f <file.csv> -o <ouput.csv>
                        ./main.py --rm-defaults-values --input-file <file.csv> --ouput-file <ouput.csv>

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
    
    trc_args = ['-tf', '--trace-file']
    if any(x in trc_args for x in sys.argv):
        chosen_tool = '8'
        return chosen_tool

    trc_args = ['-r', '--rm-defaults-values']
    if any(x in trc_args for x in sys.argv):
        chosen_tool = '9'
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
        8. Create CSV contexts applicable files from Excel trace file 
        9. Replace all default values with "-NODATA-"

    Choose your tool (1->9): """)

# Add audit result to a CSV file
if CHOSEN_TOOL == '1':

    original_filepath = ''
    original_filepath_args = ['-of', '--original-file']
    for original_filepath_arg in original_filepath_args:
        for arg in sys.argv:
            if original_filepath_arg == arg:
                original_filepath = sys.argv[sys.argv.index(arg)+1]
    if original_filepath == '':
        original_filepath = input('Which base hardening file should I look for (e.g. : filename.csv) : ')
    original_file = FileFunctions(original_filepath)
    original_file.file_exists()
    original_dataframe = original_file.read_csv_file()

    adding_filepath = ''
    adding_filepath_args = ['-af', '--adding-file']
    for adding_filepath_arg in adding_filepath_args:
        for arg in sys.argv:
            if adding_filepath_arg == arg:
                adding_filepath = sys.argv[sys.argv.index(arg)+1]
    if adding_filepath == '':
        adding_filepath = input("""
        Which audit result file should I look for (e.g. : filename.csv) : 
        """)
    adding_file = FileFunctions(adding_filepath)
    adding_file.file_exists()
    adding_dataframe = adding_file.read_csv_file()

    output_filepath = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                output_filepath = sys.argv[sys.argv.index(arg)+1]
    if output_filepath == '':
        output_filepath = input("""
        How should we name the output file ? :  
        """)

    csv = UpdateMainCsv(original_dataframe, original_filepath, adding_dataframe, adding_filepath, output_filepath)
    csv.add_audit_result()

    throw('Audit column added successfully.', 'low')

# Add Microsoft Links to CSV (Beta)
elif CHOSEN_TOOL == '2':
    hardening_filepath = ''
    hardening_filepath_args = ['-of', '--original-file']
    for hardening_filepath_arg in hardening_filepath_args:
        for arg in sys.argv:
            if hardening_filepath_arg == arg:
                hardening_filepath = sys.argv[sys.argv.index(arg)+1]
    if hardening_filepath == '':
        hardening_filepath = input('\nWhich hardening file should I look for (e.g. : filename.csv) : ')
    hardening_file = FileFunctions(hardening_filepath)
    hardening_file.file_exists()
    hardening_dataframe = hardening_file.read_csv_file()

    csv = UpdateMainCsv(hardening_dataframe, hardening_filepath)
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

    pdf2txt_filepath = ''
    pdf2txt_filepath_args = ['-pdf', '--pdf-to-txt']
    for pdf2txt_filepath_arg in pdf2txt_filepath_args:
        for arg in sys.argv:
            if pdf2txt_filepath_arg == arg:
                pdf2txt_filepath = sys.argv[sys.argv.index(arg)+1]
    if pdf2txt_filepath == '':
        pdf2txt_filepath = input('\nWhich text file should I look for (e.g. : filename.txt) : ')
    pdf2txt_file = FileFunctions(pdf2txt_filepath)
    pdf2txt_file.file_exists()
    pdf2txt_content = pdf2txt_file.read_file()

    output_filepath = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                output_filepath = sys.argv[sys.argv.index(arg)+1]
    if output_filepath == '':
        output_filepath = input('Where should we output the result (e.g. : output.csv) : ')

    pdf2txt = CISPdfScrapper(pdf2txt_content, output_filepath)
    pdf2txt.ScrapPdfData()

    throw('CIS pdf data has been scrapped successfully.', 'low')

# Add scrapped data to CSV file
elif CHOSEN_TOOL == '4':
    original_filepath = ''
    original_filepath_args = ['-of', '--original-file']
    for original_filepath_arg in original_filepath_args:
        for arg in sys.argv:
            if original_filepath_arg == arg:
                original_filepath = sys.argv[sys.argv.index(arg)+1]
    if original_filepath == '':
        original_filepath = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    original_file = FileFunctions(original_filepath)
    original_file.file_exists()
    original_dataframe = original_file.read_csv_file()

    adding_filepath = ''
    adding_filepath_args = ['-af', '--adding-file']
    for adding_filepath_arg in adding_filepath_args:
        for arg in sys.argv:
            if adding_filepath_arg == arg:
                adding_filepath = sys.argv[sys.argv.index(arg)+1]
    if adding_filepath == '':
        adding_filepath = input('Which pdf scrapped data file should I look for (e.g. : filename.csv) : ')
    adding_file = FileFunctions(adding_filepath)
    adding_file.file_exists()
    adding_dataframe = adding_file.read_csv_file()

    output_filepath = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                output_filepath = sys.argv[sys.argv.index(arg)+1]
    if output_filepath == '':
        output_filepath = input("""
        How should we name the output file ? :  
        """)

    csv = UpdateMainCsv(original_dataframe, original_filepath, adding_dataframe, adding_filepath, output_filepath)
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
        csv_filepath = ''
        csv_filepath_args = ['-csv', '--csv-file']
        for csv_filepath_arg in csv_filepath_args:
            for arg in sys.argv:
                if csv_filepath_arg == arg:
                    csv_filepath = sys.argv[sys.argv.index(arg)+1]
        if csv_filepath == '':
            csv_filepath = input('\nCsv file location : ')
        csv_file = FileFunctions(csv_filepath)
        csv_file.file_exists()
        csv_file.convert_csv_2_excel()

    elif CHOICE == '2':
        excel_filepath = ''
        excel_filepath_args = ['-xlsx', '--xlsx-file']
        for excel_filepath_arg in excel_filepath_args:
            for arg in sys.argv:
                if excel_filepath_arg == arg:
                    excel_filepath = sys.argv[sys.argv.index(arg)+1]
        if excel_filepath == '':
            excel_filepath = input('\nExcel file location : ')
        excel_file = FileFunctions(excel_filepath)
        excel_file.file_exists()
        excel_file.convert_excel_2_csv()

    else:
        throw('Wrong choice, exiting.', 'high')    
    throw("File has been converted successfully.", "low")

# Transform CSV into PowerPoint slides
elif CHOSEN_TOOL == '6':
    hardening_filepath = ''
    hardening_filepath_args = ['-csv', '--csv-file']
    for hardening_filepath_arg in hardening_filepath_args:
        for arg in sys.argv:
            if hardening_filepath_arg == arg:
                hardening_filepath = sys.argv[sys.argv.index(arg)+1]
    if hardening_filepath == '':
        hardening_filepath = input("""
        Which base hardening file should I look for (e.g. : filename.csv) :
        """)
    hardening_file = FileFunctions(hardening_filepath)
    hardening_file.file_exists()
    hardening_dataframe = hardening_file.read_csv_file()

    powerpoint_filepath = ''
    powerpoint_filepath_args = ['-o', '--output']
    for powerpoint_filepath_arg in powerpoint_filepath_args:
        for arg in sys.argv:
            if powerpoint_filepath_arg == arg:
                powerpoint_filepath = sys.argv[sys.argv.index(arg)+1]
    if powerpoint_filepath == '':
        powerpoint_filepath = input("""
        Where should I output the PowerPoint (e.g. : filename.pptx) : 
        """)

    context = None
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
    while context != '':
        context = input("\nIf there is any other column you would like to add, enter the name : ")
        if context == '':
            break
        elif context in hardening_dataframe.columns:
            contexts.append(context)
            context_name = input('Please enter the name to show in the slides : ')
            context_columns.append(context_name)
        else:
            throw('Column not found in CSV, exiting.', 'high')

    hardening_file.create_powerpoint(
        hardening_dataframe, contexts, context_columns, powerpoint_filepath)
    throw('PowerPoint has been successfully created.', 'low')

# Add scrapped data to CSV file
elif CHOSEN_TOOL == '7':
    first_filepath = ''
    first_filepath_args = ['-f1', '--first-file']
    for first_filepath_arg in first_filepath_args:
        for arg in sys.argv:
            if first_filepath_arg == arg:
                first_filepath = sys.argv[sys.argv.index(arg)+1]
    if first_filepath == '':
        first_filepath = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    first_file = FileFunctions(first_filepath)
    first_file.file_exists()
    first_dataframe = first_file.read_csv_file()

    second_filepath = ''
    second_filepath_args = ['-f1', '--second-file']
    for second_filepath_arg in second_filepath_args:
        for arg in sys.argv:
            if second_filepath_arg == arg:
                second_filepath = sys.argv[sys.argv.index(arg)+1]
    if second_filepath == '':
        second_filepath = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    second_file = FileFunctions(second_filepath)
    second_file.file_exists()
    second_dataframe = second_file.read_csv_file()

    output_filepath = ''
    output_filepath_args = ['-o', '--output']
    for output_filepath_arg in output_filepath_args:
        for arg in sys.argv:
            if output_filepath_arg == arg:
                output_filepath = sys.argv[sys.argv.index(arg)+1]
    if output_filepath == '':
        output_filepath = input('Where should we output the result (e.g. : output.csv) : ')

    csv = UpdateMainCsv(first_dataframe, first_filepath, second_dataframe, second_filepath, output_filepath)
    csv.merge_two_csv()

    throw('Scrapped data added successfully.', 'low')

# Create CSV from trace file
elif CHOSEN_TOOL == '8':
    tracefile_filepath = ''
    tracefile_filepath_args = ['-tf', '--trace-file']
    for tracefile_filepath_arg in tracefile_filepath_args:
        for arg in sys.argv:
            if tracefile_filepath_arg == arg:
                tracefile_filepath = sys.argv[sys.argv.index(arg)+1]
    if tracefile_filepath == '':
        tracefile_filepath = input('Which trace file should I look for (e.g. : filename.xlsx) : ')
    tracefile_file = FileFunctions(tracefile_filepath)
    tracefile_file.file_exists()
    
    # load Excel sheets
    df_all_policies, df_contexts = tracefile_file.read_xlsx_tracefile()

    contexts_columns = df_contexts.columns.values.tolist()

    # count contexts
    nb_contexts = 0
    for column in contexts_columns:
        if column.startswith("Context-"):
            nb_contexts+=1
    if nb_contexts == 0:
        throw("No contexts were found.", "high")

    # set the first row has header
    df_contexts.columns = df_contexts.iloc[0]
    df_contexts = df_contexts[1:]

    # Check if policy has a workshop attributed
    ws_policies = df_contexts[df_contexts["Ateliers"]!="_"]

    # add contexts with fixed value to a list
    contexts = []
    for context in range(nb_contexts):
        set_policies = ws_policies[ws_policies["Context"+str(context+1)+" - Fixed Value"]!="to check"]
        contexts.append(set_policies)

    result = tracefile_file.create_applicable_csv(contexts, df_all_policies)

    if result:
        throw('Applicable CSV created successfully.', 'low')
    else:
        throw("Couldn't create CSV files.", "high")

# Replace all default values with "-NODATA-"
elif CHOSEN_TOOL == '9':
    # input file
    file_finding_list_path = ''
    file_finding_list_path_args = ['-f', '--input-file']
    for file_finding_list_path_arg in file_finding_list_path_args:
        for arg in sys.argv:
            if file_finding_list_path_arg == arg:
                file_finding_list_path = sys.argv[sys.argv.index(arg)+1]
    if file_finding_list_path == '':
        file_finding_list_path = input('\nWhich file_finding_list file should I look for (e.g. : filename.csv) : ')

    # output file
    output_csv = ''
    output_csv_args = ['-o', '--output-file']
    for output_csv_arg in output_csv_args:
        for arg in sys.argv:
            if output_csv_arg == arg:
                output_csv = sys.argv[sys.argv.index(arg)+1]
    if output_csv == '':
        output_csv = input("\nWhat's the name of the CSV output file ? : ")

    file_finding_list_file = FileFunctions(file_finding_list_path)
    file_finding_list_file.file_exists()
    new_file_finding_list = file_finding_list_file.replace_defaults_values(output_csv)

    throw('Microsoft Link and Possible Values columns added successfully.', 'low')

else:
    throw('Tool selected not in list, exiting.', 'high')