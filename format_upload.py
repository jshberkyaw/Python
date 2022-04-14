from datetime import datetime
import os
import pandas as pd #import the pandas module 
#if not already installed run 'pip install pandas' at command prompt

# This script will take an input in a .xlsx file and 
# format it for upload into t24 with a .csv extension

try:
    #insert the location of the file sent by ebanking
    file_path =r'C:\Users\jobeng-berkyaw\Downloads\Reversals 14 Apr 2022 batch.xlsx'

    new_trans = pd.read_excel(file_path)
    print(f'reading {file_path} ....')
    #insert constant columns
    new_trans.loc[:,'trans_label'] = 'AC89' 
    new_trans.loc[:, 'currency'] = 'GHS' 
    new_trans.loc[:, 'short_account'] ='GH0010001'
    new_trans.loc[:,'empty_column']=''

    #rearrange the columns to match output file
    new_trans_formatted = new_trans[['trans_label', 'Debit Account','currency'
    ,'Amount','empty_column','Narration','Narration',  
    'currency','Credit Account', 'short_account']]
    date = datetime.now().strftime('%Y%m%d')
    #export output file with no headers
    new_trans_formatted.to_csv(f'E-{date}REV1.csv', index=False, header=None)
    print('exporting final file......')
    print('file export completed.')
    print('press any key to clear completed file...')
    os.system('pause')
    #comment the next line out if you'd like the input file to be deleted
    #os.remove(file_path) 
    print('cleaning up.....')
except FileNotFoundError as e:
    print("Error!! File not found")
    os.system('pause')
except PermissionError as e:
   print(e)
except KeyError as e:
    print(e)
except Exception as e:
    print(e)
