import pandas as pd
import json
import re
import traceback
import sys

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)   
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
  
os.chdir(application_path)

input_json_filepath = input("Json file path : ").strip().strip('"')

if not os.path.exists(input_json_filepath):
    print("File doesn't exists : ", input_json_filepath)
    input('\nPress enter to exit...')
    sys.exit(1)

output_filepath = "output.xlsx"
if os.path.exists(output_filepath):
    c = 1
    while True:
        output_filepath = f"output-({c}).xlsx"
        if not os.path.exists(output_filepath):
            break
        c+=1
try:
    with open(input_json_filepath,'r',encoding='utf8') as f:
        data = json.load(f)

    # preparing xlsx header get all dict keys if inconsistent.
    header = {}
    for entry in data:
        for key in entry.keys():
            if not header.get(key):
                header[key] = 1

    header = list(header.keys())

    # creating excel from json records
    writer = pd.ExcelWriter(output_filepath, engine='xlsxwriter')

    info_df = pd.DataFrame.from_records(data,columns=header) 
    info_df.to_excel(writer,sheet_name="DataSheet",index=False)  

    writer.save()
except:
    print(traceback.format_exc())
    print('\n\n')
    
if os.path.exists(output_filepath):
    if os.path.getsize(output_filepath):
        print('Successfully converted json to excel.')
    else:
        os.remove(output_filepath)
        print('Something went wrong!')
else:
    print('Something went wrong!')
    
input('\nPress enter to exit...')
sys.exit(1)