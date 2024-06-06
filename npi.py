#!/usr/bin/env python3
"""
Author : Jon Freivald <jfreivald@brmedical.com>
        : Copyright © Blue Ridge Medical Center, 2024. All Rights Reserved.
        : License: GNU GPL Version 3
Date   : 2024-06-05
Purpose: NPI extraction tool
        : Uses Excel spreadsheets as data sources.
        : Extracts NPIs from master Medicaid spreadsheet (>150,000 records) based on local list of provider NPIs.
        : Version change log at EoF.
"""

# Distribution license for PiSimpleGUI v5
PySimpleGUI_License = 'ecyJJvMkaRWQNqlebknkN2lJVrHslwwvZ7SAIY6cIykURJpSct3zRNylamWaJu1yddGqlHvBbziMIuswIokQxdppY52UVNuqcd2NVcJbRUC7Iz6mMUTGcixvMnDwg03kN0jvcf2pMTSPwKiiTWGjlEjDZWWE5XzMZYUtRolEcLGMxTvie3WQ1llnbdnXRMWUZoXFJVzDacWM9wukIfjPoDiINvSr4swwIXiHw5iOTfmuFXtPZeUNZApScZnbNE0hIujto6iHSPmx9juzI2iIw7inTKmuF4tIZMU3xXhDcK3HQKiVODiZJRGPc1mtVcpld4mvFws8ZFChIZs6ICk2NTvlbpXQBBhobtnvk8isO7iBJYCNb4H7VBlhIlFdJapdZHGYdolsIBER1hlbZYGnlOjJYOWQwZgnQg2yVzu1d4GbVWyQIpiww3iFQD3nVczddWGV9RtlZZX6JpJtRACsI66RIAj6kP3TNXTeEgi7LRCdJKEhY3X9RWlvSrXVNszUdgW4VckUIvjGoAi2M2jlAgymNvCm0Ew5Miy80lx0OFSWIBsnIokeR6hOdPG4V0FPeaHNBTpJcSmmVuzkI5jZoIiLMxjgALymNgSN0NwtM0yR0uxCOESDISsmI8kMVOtrYVWQlAsGQkWFRak0cdmRVlz0c8ytIs6hIbmGpIm3cmm5V6pTdZmqFMsYZCE9BFifc9ml1ilGZHGZl6jeYRWGwYuMYF2t9ltnI6iNwOiSSUVDBcB2ZeGpRoyKZXXMNZzEINjCojifM2jUAC1BLwj3ILyMMDCo4vyKMGzeUaufMRzKQGiyf0Qt=E=z7800ccb7385edf61c64307a5cea538951e0f20b3d78dac31ff5922a3108d9321821be312038eead936338932ccb9b1753a30efdf49da75ee9d1e68d1b18b1e1bf0f04eb933839587d8962a38b0799dba4fd9e4a35a179b9cf8b67450e3213879dcd95b40e1520faa24344afa9254b8a1a77630c539f714659edaf53359eaacbb4be729d5deb60cb2c8e826aa6121c7187aecc119c460b0c3c949fc23d8e607ce9735225ef2e27bf2df565fc7d6a2aac6068980bd262dadcdb8cfd53e9656829dd1c83d8b275d216e4599bf342ab133434446daefa04e2042095ae88862164fbe0018a48de85914f2481beacc60e206bc306e512e38d129f7f6f2b3a9b4f6924a8da132afa74900d8ad4d603b1bd30f44e5293f1f70c972631db070ec41a2440a14caaf1092e4c1edec3c2b3e90b984a6daf4b5acaf1272364aa8a1dd8b9b8a1f851c6eed0ca4a0499ec394c285bcc9d308d03d82e2e96098e862cf0524b902361385b604bf3fa39c58eab6c7a6874670c53cebcfde6e08897243d288188b302338c82ad89ab803066082206d0240193e569c24ea742a3a860cd8b4eb5f4f04c7563bc77ec8724b6466fc06c819dd0f9684f7d777318cafc1c5fe0010112cf9dcfb7bd017a01e36b37c973e70ebc83de2f0431f11a8ab9271d9aaff7c8b473c54fc1be2a786c86be97ff2e22261638a63e7ac2d861d96fbbab2c33de2bd3ff191'

import openpyxl
from openpyxl import Workbook
import argparse
import PySimpleGUI as sg
import os
import pickle

BRMC = {'BACKGROUND': '#73afb6',
                 'TEXT': '#00446a',
                 'INPUT': '#ffcf01',
                 'TEXT_INPUT': '#00446a',
                 'SCROLL': '#ce7067',
                 'BUTTON': ('#ffcf01', '#00446a'),
                 'PROGRESS': ('#ffcf01', '#00446a'),
                 'BORDER': 1, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0,
                 }
sg.theme_add_new('BRMC', BRMC)

progver = 'v 0.2'
mainTheme = 'BRMC'
errorTheme = 'HotDogStand'
config_file = (f'{os.path.expanduser("~")}/npi_config.dat')
output_file = (f'{os.path.expanduser("~")}/Downloads/MedicaidProviderExtract.xlsx')

# --------------------------------------------------
def get_args():
	"""Process any command line arguments"""

	parser = argparse.ArgumentParser(
		description='View provider Medicaid validation status.',
		formatter_class=argparse.ArgumentDefaultsHelpFormatter)

	parser.add_argument('-f',
		'--file',
		help='Local provider list containing NPIs to look up.',
		metavar='filename',
		type=str,
		default='//brmc-fs2012.int.brmedical.com/ADMIN/Renee Dolder/Credentials/BRMC/Provider List & Charts/Provider lists & charts/Provider Charts - in house/Providers.xlsx')

	args = parser.parse_args()

	return args

# --------------------------------------------------
def get_user_settings():

    user_config = {}

    try:
        with open(config_file, 'rb') as fp:
            user_config = pickle.load(fp)
        fp.close()
    except:
        user_config['Theme'] = mainTheme

    return user_config

# --------------------------------------------------
def write_user_settings(user_config):

    try:
        with open(config_file, 'wb') as fp:
            pickle.dump(user_config, fp)
        fp.close()
    except:
        sg.theme(errorTheme)
        layout = [ [sg.Text(f'File or data error: {config_file}. Updates NOT saved!')],
					[sg.Button('Quit')] ]
        window = sg.Window('FILE ERROR!', layout, finalize=True)

        while True:
            event, values = window.read()
            if event in (sg.WIN_CLOSED, 'Quit'): # if user closes window or clicks quit
                window.close()
                return

# --------------------------------------------------
def create_extract(output_file, args, window):

    npiList = [] # Provider List w/NPI
    npiFile = openpyxl.load_workbook(args.file)

    currentSheet = npiFile['Providers']
    currentProvider = ''
    for row in range(8, currentSheet.max_row + 1): # start @ 2 because of the header row!
        cellA = (f'A{row}')
        cellD = (f'D{row}')
        if currentSheet[cellA].value != None: # We were pulling blank rows for some reason -- filter them out!
            if currentSheet[cellA].value != currentProvider:
                currentProvider = currentSheet[cellA].value
            currentNPI = currentSheet[cellD].value
            npiList.append([currentNPI])

    npiFile.close()

    window['-STATUS_MSG-'].update('Local NPIs loaded. Loading Medicaid source file...')
    window.refresh()

    try:
        src = openpyxl.load_workbook(sg.popup_get_file('Medicaid source file'))
    except:
        return False
    
    window['-STATUS_MSG-'].update('Processing Medicaid source file, please be patient..!')
    window.refresh()

    currentSheet = src.active

    wb = Workbook() # create our output file
    ws = wb.active

    ws['A1'] = 'NPI'
    ws['B1'] = 'Last Name'
    ws['C1'] = 'First Name'
    ws['D1'] = 'Middle Name'
    ws['E1'] = 'Enrollment Type'
    ws['F1'] = 'Revalidation Date'
    ws['G1'] = 'Provider Type Description'
    ws['H1'] = 'Zip Code'

    for row in range(2, currentSheet.max_row + 1): # our file has headers
        c_npi = (f'B{row}')
        c_lname = (f'E{row}')
        c_fname = (f'C{row}')
        c_mi = (f'D{row}')
        c_en_type = (f'K{row}')
        c_rv_date = (f'M{row}')
        c_type_desc = (f'H{row}')
        c_zip_code = (f'L{row}')

        if currentSheet[c_npi].value != None: # eliminate blank lines
            for i in npiList:
                if currentSheet[c_npi].value in i:
                    npi = currentSheet[c_npi].value
                    lname = currentSheet[c_lname].value
                    fname = currentSheet[c_fname].value
                    mi = currentSheet[c_mi].value
                    en_type = currentSheet[c_en_type].value
                    rv_date = currentSheet[c_rv_date].value
                    type_desc = currentSheet[c_type_desc].value
                    zip_code = currentSheet[c_zip_code].value
                    ws.append([npi, lname, fname, mi, en_type, rv_date, type_desc, zip_code])

    window['-STATUS_MSG-'].update('Done! Writing output file...')
    window.refresh()
    wb.save(output_file) # write the output file
    return True

# --------------------------------------------------
def extract_NPI_data():

    args = get_args()
    user_config = get_user_settings()
    if 'winLoc' in user_config:
        winLoc = user_config['winLoc']
    else:
        winLoc = (50, 50)
    if 'winSize' in user_config:
        winSize = user_config['winSize']
    else:
        winSize = (450, 190)
    
    
    sg.theme(user_config['Theme'])
    layout = [  [sg.Image('logo.png', size=(400, 96))],
                [sg.Text('Extract local NPI data from Medicaid source file.')],
                [sg.Text('', key='-STATUS_MSG-')],
                [sg.Button('Open'), sg.Text('<-- Medicaid file'), sg.Push(), sg.Button('Quit')],
                 [sg.Push(), sg.Text('Copyright © Blue Ridge Medical Center, 2023, 2024')] ]
    window = sg.Window(f'Provider NPI Query Tool {progver}', layout, location=winLoc, size=winSize, element_justification='center', grab_anywhere=True, resizable=True, finalize=True)
    window.BringToFront()

    while True:
        
        event, values = window.read(timeout=5000)
        winLoc = window.CurrentLocation()
        winSize = window.Size

        if event in (sg.WIN_CLOSED, 'Quit'): # if user closes window or clicks cancel
            if event == 'Quit':
                user_config['winLoc'] = winLoc
                user_config['winSize'] = winSize
                write_user_settings(user_config)
            break

        if event == 'Open':
            if create_extract(output_file, args, window):
                sg.popup(f'Extraction complete. Your data is in {output_file}')
            break
                
    window.close()

# --------------------------------------------------
if __name__ == '__main__':
	extract_NPI_data()
        
"""Change log:

    v 0.1   : 240605    : Initial version
    v 0.2   : 240606    : Added status update messages
"""