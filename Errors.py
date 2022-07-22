def throw(message, level):
    switcher = { 
        "low": "\033[92m", # green
        "medium": "\033[93m", # yellow
        "high": "\033[91m", # red
    } 
    
    color = switcher.get(level)

    print('\n'+color + message + '\033[0m')
    print("""
Thanks for using HardeningBox, see you later !

##############################################################################################################################
""")
    exit()

