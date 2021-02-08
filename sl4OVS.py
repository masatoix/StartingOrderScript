# %%
import pandas as pd
import glob
import re
import os


# %%
import_file_path = 'original/'
export_file_path = 'split/'
finalize_file_path = 'finalized/'


# %%
for i in glob.glob(import_file_path+'*.xlsx'):
    input_book = pd.ExcelFile(i)
    for sheet_name in input_book.sheet_names:
        input_split = input_book.parse(sheet_name)
        input_split.to_excel(export_file_path+sheet_name+'.xlsx')

for fname in glob.glob(export_file_path+'*.xlsx'):
    file_name = os.path.split(fname)[1]

    df_change = pd.read_excel(fname)
    df_change['e'] = df_change['a'].str.split(pat='\s', expand=True)[0]
    df_change['f'] = df_change['a'].str.split(pat='\s', expand=True)[1]
    
    # cols = df_change.columns.tolist()
    df_finalize = df_change.loc[:, ['e', 'f', 'b', 'c', 'd']]
    
    df_finalize.to_excel(finalize_file_path+'f_'+file_name, index=False, header=False)
