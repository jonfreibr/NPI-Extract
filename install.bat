@echo off
rem Example install batch file for Windows environment.
rem Run from the application directory
rem Required files:
rem     logo.png                    logo file displayed by application at runtime
rem     NPI Extraction.lnk          create this shortcut file to reflect your installation and runtime requirements!!
rem     npi.py                      The main application file
rem     python-3.11.5-amd64.exe     The Python source file
rem     requirements.txt            Used by Python to install required libraries
rem     unattend.xml                Required to run a "hands-off" install of Python.
rem
rem echo Installing Python 3.11.5
python-3.11.5-amd64.exe /passive
rem echo upgrading pip
%LocalAppData%\Programs\Python\Python311\python.exe -m pip install --upgrade pip
rem echo Adding package requients
%LocalAppData%\Programs\Python\Python311\Scripts\pip.exe install -r requirements.txt
rem echo Copying files
copy /y "NPI Extraction.lnk" %USERPROFILE%\Desktop
rem echo Done with installation
rem pause
rem exit