from FileFunctions import *
from UpdateMainCsv import *
from CISPdfScrapper import *

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

    output_column_name = input('What will be the name of the output column (e.g. : context1) : ')
    output_column_index = input('What will be the index of the output column (e.g. : 15) : ')

    csv = UpdateMainCsv(original_dataframe, original_filepath, adding_dataframe, adding_filepath)
    csv.AddAuditResult()

    print('\nAudit column added successfully.')

elif tool == '2':
    # Add Microsoft Links to CSV (Beta)
    hardening_filepath = input('\nWhich hardening file should I look for (e.g. : filename.csv) : ')
    hardening_file = FileFunctions(hardening_filepath)
    hardening_file.checkIfFileExistsAndReadable()
    hardening_dataframe = hardening_file.readCsvFile()

    csv = UpdateMainCsv(hardening_dataframe, hardening_filepath)
    csv.AddMicrosoftLinks()

    print('\nMicrosoft Link and Possible Values columns added successfully.')

elif tool == '3':
    # Scrap policies from CIS pdf file (https://downloads.cisecurity.org/#/)
    input("""
    In order to prepare this tool, you need to transfer pdf text data into a txt file.
    To do that, you need to open your pdf with a pdf reader, and select the whole text (CTRL+A), it might take few seconds, and copy it (CTRL+C).
    When the data is copied, you need to paste it in a file and save it as a txt file.
    
    You also need to remove every page until first policy (Recommandation part only),
    then you can remove every data after the policies aswell (Appendix).
    
    yes(y) ? : """)

    pdf2txt_filepath = input('\nWhich hardening file should I look for (e.g. : filename.csv) : ')
    pdf2txt_file = FileFunctions(pdf2txt_filepath)
    pdf2txt_file.checkIfFileExistsAndReadable()
    pdf2txt_content = pdf2txt_file.readFile()

    pdf2txt = CISPdfScrapper(pdf2txt_content)
    pdf2txt.ScrapPdfData()

    print('CIS pdf data has been scrapped successfullys.')

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

    print('\nScrapped data added successfully.')

elif tool == '5':
    # Excel <-> CV convertion
    pass

elif tool == '6':
    # Transform CSV into PowerPoint slides
    pass

else:
    print('\nTool selected not in list.')


print("""
Thanks for using HardeningBox, see you later !

##############################################################################################################################
""")