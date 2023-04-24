# HardeningBox
 This is a tool box for CIS Windows Hardening Benchmarks.<br/>
 It allows you to use any existing data for further use.

 You can, for example, transform a CSV file into PowerPoint Slides, or an Excel File.

 You can even transform a pdf benchmark into a CSV file, wich is very useful.

 ## Requirements

You can install python requirements by launching this command :
```
pip3 install -r requirements.txt
```

## Tools
Actually, this tool box presents 6 tools :

### 1. Add Audit result to a CSV file

Simply, it adds csv columns from a csv file to another csv file. <br/>

Let's say, for example, that you have done an Hardening Audit with HardeningKitty, and you want to present an Excel file with all details (RegistryKey, DefaultValue, Result, RecommendedValue, ...), so you might merge your result CSV with you finding list.

This tool will add Result column to your finding list, and you will be able to present it easily.

### 2. Add Microsoft Links to CSV (Beta)

This tool will look for Microsoft Links going with any policy in your finding list and add a column to your file in order to use it afterward.

### 3. Scrap policies from CIS pdf file

<i>Any CIS benchmarks can be found here : (https://downloads.cisecurity.org/#/).</i>

This tool will read a file containing text data of a CIS benchmark and fetch it to obtain a CSV file with those columns : 
Default Value, Recommended Value, Description, Impact and Rationale.

<b>In order to use this tool, you might transfer pdf text data into a txt file.
To do that, you need to open your pdf with a pdf reader, and select the whole text (CTRL+A), it might take few seconds, and copy it (CTRL+C).
When the content is copied, you need to paste it in a file and save it as a txt file.

You also need to remove every page until first policy (Recommendation part only),
then you can remove every data after the policies aswell (Appendix).</b>

### 4. Add Scrapped data to CSV

This tool will add scrapped data from the previous tool to your main CSV file.

### 5. Excel <-> Csv convertion

This tool will simply transform an Excel file into a Csv file or a Csv file into a Excel file. This might be useful if you need to edit values easily, with Excel.

### 6. Transform CSV into PowerPoint file

This one is a powerfull tool that will transform your finding list into PowerPoint slides.

Each slide will contain policy values, and it will help you to present an Hardening project to everyone else.

### 7. Merge 2 csv files and remove duplicates by "Names"

This tool allows you to merge two csv files based on a column value

### 8. Replace all default values with "-NODATA-"

Replace default value in csv by -NODATA-, it allows a better understanding of the actual status.

### 9. Excel report file to CSV

This tool will transform an excel report file into multiple csv file applicable, by context, by workshop & by category