from FileFunctions import *
from UpdateMainCsv import *
from CISPdfScrapper import *
from Errors import throw

tool = input("""
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

    1. Add audit result to a CSV file
    2. Add Microsoft Links to CSV (Beta)
    3. Scrap policies from CIS pdf file (https://downloads.cisecurity.org/#/)
    4. Add scrapped data to CSV file
    5. Excel <-> CSV convertion
    6. Transform CSV into PowerPoint slides

Choose your tool (1->6): """)

if tool == '1':
    # Add audit result to a CSV file
    original_filepath = input('Which base hardening file should I look for (e.g. : filename.csv) : ')
    original_file = FileFunctions(original_filepath)
    original_file.checkIfFileExistsAndReadable()
    original_dataframe = original_file.readCsvFile()

    adding_filepath = input('Which audit result file should I look for (e.g. : filename.csv) : ')
    adding_file = FileFunctions(adding_filepath)
    adding_file.checkIfFileExistsAndReadable()
    adding_dataframe = adding_file.readCsvFile()

    csv = UpdateMainCsv(original_dataframe, original_filepath, adding_dataframe, adding_filepath)
    csv.AddAuditResult()

    throw('Audit column added successfully.', 'low')

elif tool == '2':
    # Add Microsoft Links to CSV (Beta)
    hardening_filepath = input('\nWhich hardening file should I look for (e.g. : filename.csv) : ')
    hardening_file = FileFunctions(hardening_filepath)
    hardening_file.checkIfFileExistsAndReadable()
    hardening_dataframe = hardening_file.readCsvFile()

    csv = UpdateMainCsv(hardening_dataframe, hardening_filepath)
    csv.AddMicrosoftLinks()

    throw('Microsoft Link and Possible Values columns added successfully.', 'low')

elif tool == '3':
    # Scrap policies from CIS pdf file (https://downloads.cisecurity.org/#/)
    input("""\033[93m
    In order to prepare this tool, you need to transfer pdf text data into a txt file.
    To do that, you need to open your pdf with a pdf reader, and select the whole text (CTRL+A), it might take few seconds, and copy it (CTRL+C).
    When the data is copied, you need to paste it in a file and save it as a txt file.
    
    You also need to remove every page until first policy (Recommandation part only),
    then you can remove every data after the policies aswell (Appendix).
    
    yes(y) ? : \033[0m""")

    pdf2txt_filepath = input('\nWhich text file should I look for (e.g. : filename.txt) : ')
    pdf2txt_file = FileFunctions(pdf2txt_filepath)
    pdf2txt_file.checkIfFileExistsAndReadable()
    pdf2txt_content = pdf2txt_file.readFile()

    output_filepath = input('Where should we output the result (e.g. : output.csv) : ')

    pdf2txt = CISPdfScrapper(pdf2txt_content, output_filepath)
    pdf2txt.ScrapPdfData()

    throw('CIS pdf data has been scrapped successfully.', 'low')

elif tool == '4':
    # Add scrapped data to CSV file
    original_filepath = input('Which hardening file should I look for (e.g. : filename.csv) : ')
    original_file = FileFunctions(original_filepath)
    original_file.checkIfFileExistsAndReadable()
    original_dataframe = original_file.readCsvFile()

    adding_filepath = input('Which pdf scrapped data file should I look for (e.g. : filename.csv) : ')
    adding_file = FileFunctions(adding_filepath)
    adding_file.checkIfFileExistsAndReadable()
    adding_dataframe = adding_file.readCsvFile()

    csv = UpdateMainCsv(original_dataframe, original_filepath, adding_dataframe, adding_filepath)
    csv.AddScrappedDataToCsv()

    throw('Scrapped data added successfully.', 'low')

elif tool == '5':
    # Excel <-> CV convertion
    choice = input('''
Would you like to :

1. Convert a Csv file to an Excel file 
2. Convert an Excel file to a csv file

(1 or 2) : 
''')

    if choice == '1':
        csv_filepath = input('\nCsv file location : ')
        csv_file = FileFunctions(csv_filepath)
        csv_file.checkIfFileExistsAndReadable()
        csv_file.convertCsv2Excel()

    elif choice == '2':
        excel_filepath = input('\nExcel file location : ')
        excel_file = FileFunctions(excel_filepath)
        excel_file.checkIfFileExistsAndReadable()
        excel_file.convertExcel2Csv()

    else:
        throw('Wrong choice, exiting.', 'high')
    
    throw("File has been converted successfully.", "low")

elif tool == '6':
    # Transform CSV into PowerPoint slides
    hardening_filepath = input('Which base hardening file should I look for (e.g. : filename.csv) : ')
    hardening_file = FileFunctions(hardening_filepath)
    hardening_file.checkIfFileExistsAndReadable()
    hardening_dataframe = hardening_file.readCsvFile()

    powerpoint_filepath = input('Where should I output the PowerPoint (e.g. : filename.pptx) : ')

    context = None
    contexts = []
    context_columns = []
    print("""\033[93m
Actual Columns : 

    • PossibleValues (Empty if column does not exists)
    • DefaultValue
    • RecomendedValue
    • Comment (Empty)
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

    hardening_file.CreatePPTX(hardening_dataframe, contexts, context_columns, powerpoint_filepath)
    throw('PowerPoint has been successfully created.', 'low')

else:
    throw('Tool selected not in list, exiting.', 'high')
