Small tool to convert an spreadsheet Excel (.xlsx) or LibreOffice (.ods) file to WOCE [WHP-Exchange format](https://exchange-format.readthedocs.io/en/latest/) (BOTTLE)

Can be used by calling python file or notebook locally or by launching the notebook through COLAB

Excel file needs to have the following format:
 - FIRST (from the left) sheet has to contain the table of data WITH header names in 1st row, units in 2nd one and data below
 - SECOND (from the left) sheet has to contain the metadata if any or be an empty sheet

if python file is called directly, please use the following format: ```excel2whpexchange.py [file_name] [institution]```

- file_name will be the spreadsheet
- institution will be the identificator for creator institution or user as [recommended by WHP-Exchange[(https://exchange-format.readthedocs.io/en/latest/common.html#file-identification-stamp)

example for me:
```$ python excel2whpexchange.py cruisefile.xlsx CSICIIMAVL```

