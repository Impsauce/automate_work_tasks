#This program copies filenames(values) from a column in an excel file, looks for those filenames in a user-inputted directory,
#if those filenames exist in the directory, the program copies the files into another user-inputted directory

import pandas as pd
import os
import shutil

folder = input('Enter the absolute filepath of'
               ' the directory you wish to copy from: ')
validated = False
while not validated:
    try:
        excel_file = input("Enter the full path of your excel file: ")
        df = pd.read_excel(excel_file)
        validated = True
    except(IOError) as e:
        print(e)      
print("File read successfully")
column_name = input("Enter the column header containing the file names exactly as it is on your sheet: ")
column_list = df[column_name].tolist()
destination = input("Enter destination folder's absolute filepath: ")
print('Copying files...')
for folders, subfolders, filenames in os.walk(folder):
        for filename in filenames:
            if filename in column_list:
                shutil.copy(os.path.join(folders, filename), destination)
print ("Finished copying files")