# excel2whpexchange
[![DOI](https://zenodo.org/badge/254050254.svg)](https://zenodo.org/badge/latestdoi/254050254)
<a href="https://colab.research.google.com/github/avelo/stuff/blob/master/excel2whpexchange.ipynb" target="_parent"><img src="https://colab.research.google.com/assets/colab-badge.svg" alt="Open In Colab"/></a>

Small tool to convert an spreadsheet Excel (.xlsx) or LibreOffice (.ods) file to WOCE [WHP-Exchange format](https://exchange-format.readthedocs.io/en/latest/) (BOTTLE)

This tool was developed for helping users to generate a properly formatted WHP-Exchange format file, automatically doing the following:
 - Generating BOTTLE tag
 - Including and commenting metadata
 - Generating header and units line
 - Fixing known paramenters to recommended precision decimals
 - Setting END_BOTTLE tag

Can be used by calling python file or notebook locally or by launching the notebook through COLAB

Excel file needs to have the following format:
 - FIRST (from the left) sheet has to contain the table of data WITH header names in 1st row, units in 2nd one and data below
 - SECOND (from the left) sheet has to contain the metadata as text (if any) or be an empty sheet

if python file is called directly, please use the following format: ```excel2whpexchange.py [file_name] [institution]```

- file_name will be the spreadsheet
- institution will be the identificator for creator institution or user as [recommended by WHP-Exchange](https://exchange-format.readthedocs.io/en/latest/common.html#file-identification-stamp)

example for me:
```$ python excel2whpexchange.py cruisefile.xlsx CSICIIMAVL```

