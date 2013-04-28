# Mercurial google spreadsheet reporter
=====================

This small python script takes the history of several mercurial repositories, and forms a simple time report, 
based on commits date and time.

Before using it please be sure that Google preserve history properly. It is even suggested to make a
copy of your Google Spreadsheet. There is no guarantee that this script won't damage your already
exist data, as it does pretty aggressive resizing of the spreadsheet. It is designed to preserve
the old history, but there is no guarantee. Also be sure that Spreadsheet has supported template.
Otherwise you probably loose your data. Use this tool on your own risk.!!!

This script uses [gspread]https://github.com/burnash/gspread, so you should install it before running the script.
Just type:
```sh
pip install gspread
```

Or:

```sh
easy_install gspread
```
Thanks to [burnash](https://github.com/burnash) for providing [gspread]https://github.com/burnash/gspread.

# Using the script:

Before running the script, the spreadsheet should be created. OAuth2 currently is not supported by [gspread]https://github.com/burnash/gspread, so the scrpt doesn't support it either.

First row of the worksheet should have headers - currently hgSpreadsheetReporter doesn't support creating it.

The required headers are:
 - Name
 - Date
 - Total

The header names can be changed in config.ini

You should fill up config.ini as well. The example can be taken from config_example.ini

After filling up config.ini just run the script.
