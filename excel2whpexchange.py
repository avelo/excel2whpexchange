import sys
import os
import datetime as dt
import pandas as pd
# 
try:
    SPREADSHEET_FILE = sys.argv[1]
except:
    sys.exit('This script will take an spreadsheet (.xlsx or .ods) as argument and convert to WHP-Exchange BOTTLE format\n\n' +
          'FIRST (from the left) sheet has to contain the table of data WITH names in 1st row, units in 2nd one and data below\n' +
          'SECOND (from the left) sheet has to contain the metadata if any or be an empty sheet\n\n' +
          'please use the following format: excel2whpexchange.py [file_name] [INSTITUTION]\n'+
          '  example: python excel2whpexchange.py cruisefile.xlsx CSICIIMAVL' )
    #SPREADSHEET_FILE = 'input_file.xlsx'
    #SPREADSHEET_FILE = 'input_file.ods'
#
try:
    INSTITUTION = sys.argv[2]
except:
    INSTITUTION = 'INSTITUTION'

PRETTY_PRINT = True
#
## Define Precisions
precision = {}
for k in {'STNNBR', 'CASTNO', 'BTLNBR', 'SAMPNO'}:
    precision[k] = 0
for k in {'CTDOXY', 'OXYGEN', 'SILCAT', 'NITRAT', 'NITRIT', 'PHSPHT', 'TCARBN', 'ALKALI', 'CFC113', 'SF6', 'CCL4', 'HELIUM', 'HELIUM_ERR', 'DELHE3', 'DELHE3_ERR'}:
    precision[k] = 2
for k in {'CFC11', 'CFC12', 'TRITUM', 'TRITUM_ERR'}:
    precision[k] = 3
for k in {'LATITUDE', 'LONGITUDE', 'CTDPRS', 'CTDTMP', 'CTDSAL', 'SALNTY', 'PHTS25P0', 'NEON', 'NEON_ERR'}:
    precision[k] = 4

## Read FIRST sheet with data
if not os.path.isfile(SPREADSHEET_FILE):
    sys.exit(f'File {SPREADSHEET_FILE} not found')
ext = SPREADSHEET_FILE.split('.')[-1]
if not ext in ['xlsx', 'ods']:
    sys.exit(f'Only spreadsheets in Excel (.xlsx) or OpenDocument (.ods) format allowed')
engine = 'odf' if ext == 'ods' else 'xlrd'
print(f'Processing {SPREADSHEET_FILE}')
df = pd.read_excel(SPREADSHEET_FILE, 0, dtype=object, encoding='utf-8', engine=engine)
units = df.iloc[0, :]
df = df.iloc[1:, :]
# Set precisions
for k, v in precision.items():
    try:
        df[k] = df[k].astype(float).apply(lambda x: f'{x:{9 if PRETTY_PRINT else ""}.{v}f}')
    except:
        pass
        #print(f'.. unable to format or unexistent: {k}')
try:
    df['TIME'] = df['TIME'].astype(float).apply(lambda x: f'{x:04.0f}')
except:
    pass
# Remove -999 cols
cols_to_remove = []
cols_to_remove2 = []
for col in df:
    if (df[col].astype(float, errors='ignore') == -999).all():
        cols_to_remove.append(col)
        if f'{col}_FLAG_W' in df:
            cols_to_remove2.append(f'{col}_FLAG_W')
print(f'..following columns removed due to -999: {",".join(cols_to_remove)}')
print(f'..following columns removed due to nonsense without previous ones: {",".join(cols_to_remove2)}')
cols_to_remove.extend(cols_to_remove2)
df = df.drop(columns=cols_to_remove)
units = units.drop(cols_to_remove)
## Read SECOND sheet with metadata
try:
    metadata = pd.read_excel(SPREADSHEET_FILE, 1, dtype=str, encoding='utf-8', engine=engine, header=None, na_filter=False).to_csv(header=False, index=False)
except:
    metadata = ''
# Write file
outfile = SPREADSHEET_FILE.split('.')[0] + '_hy1.csv'
print(f'Writting WHP-Exchange BOTTLE file in: {outfile}')
with open(outfile, 'wt', encoding='utf-8', newline='') as fout:
    yyyymmdd = dt.datetime.now().strftime('%Y%m%d')
    fout.write(f'BOTTLE,{yyyymmdd}{INSTITUTION}\n')
    for line in metadata.splitlines():
        line = line.strip('"')
        fout.write(f'# {line}\n')
    fout.write(','.join(df.columns)+'\n')
    fout.write(','.join(list(units.fillna('')))+'\n')
    fout.write(df.to_csv(header=False, index=False, line_terminator='\n'))
    fout.write('END_DATA')
print('Done')
