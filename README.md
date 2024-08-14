# Medicaid NPI Extraction Tool

## License & Copyright

Copyright © Blue Ridge Medical Center, 2024. All Rights Reserved.  
License: GNU GPL Version 3

## Rationale

This tool is used to extract local data from the Medicaid enrolled provider master file weekly State (Virginia) extraction. I decided to write this tool when I found our credentialling specialist attempting to validate our local provider data by digging through the State extract spreadsheet, which at that time exceeded 150,000 records. The file was so large that Excel was very sluggish and her manual searching was taking her many hours every week. This code is based on the format of the Virginia extract [found here: ](https://vamedicaid.dmas.virginia.gov/provider/mco). I would expect other states to follow a similar format and this code should be easily adaptable.

After briefly reviwing the data, I determined that I could use the provider's NPI (National Provider Identifier) as a key and extract our local data from this sheet, providing her with a sheet consisting of only our data, and only the data fields she needed. At the time of initial development, this distilled the > 150K records of the state level extract to 60 records of local data. The extraction took under 30 seconds (most of which as load time on the raw data), and her subsequent validation took her a matter of several minutes vs. the former many hours.

## Reqiurements

- Python 3.11
- PySimpleGUI 5.0
- openpyxl 3.1.2

A Distribution Key for PySimpleGUI (Commercial License) is embedded in the code.

The requirements for the program are simple. There needs to be a source file that is a spreadsheet (.xlsx format) contains a list of local NPIs. In our case, a local spreadsheet existed that listed the provider name in column A and their NPI in column D, starting on row 8. You can create a spreadsheet with this specification, or modify the code to match the structure of your existing spreadsheet. The program will expect "Providers.xlsx" in the current directory. You can specify a file and location with '-f' or '--file' argument. 

Our state (Virginia) provides the extract file as a table in an .xlsd file. The extraction program requires an .xlsx file. Conversion is simply opening the file in Excel and saving it as an .xlsx, performing any sorts you want prior to saving. (Sorting is still relatively quick -- filtering gets mind-numbingly slow on this many records!)

## Operation

1. Launch the program. This can be done from the command line, or via a shortcut. (The included installation batch file will expect a shorcut named "NPI Extraction.lnk", which you will need to create as appropriate for your environment. The installation batch file also expects the python source file python-3.11.5-amd64.exe to be in this directory. Update this file per your requirments.)
       Examples:
           python npi.py
           python npi.py -f d:\myNPIlist.xlsx
           python npi.py --file d:\myNPIlist.xlsx
2.  On launch, the two options presented are "Open" and "Quit". Quit should be self-explanatory.
3.  Clicking "Open" will load the local NPI file (a count of NPIs loaded will be displayed) and a file open dialog will appear.
4.  Select the appropriate .xlsx file (NOT the original .xlsb) and click "OK".
5.  The program will load the source file and then parse the records.
6.  Once parsing starts, a count of records processed will be displayed.
7.  When complete, the total records processed, the number of records extracted, and the time elapsed will be displayed.
8.  A pop-up will appear showing that the extracted data has been saved in your Downloads folder under the name "MedicaidProviderExtract.xlsx".
9.  Clicking "OK" will close the program.
