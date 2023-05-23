# PyODCSVConverter

PyODCSVConverter (Python OpenDocument Comma Separated Values Converter) is a Python 3 script that automatically exports every sheet within a spreadsheet document to a CSV file.

This script was heavily based on [PyODConverter](https://github.com/mirkonasato/pyodconverter/).

(Please see the warning about flashing lights after the Usage section)


## Usage

You must first have a LibreOffice server running in the background. Use the following command:

```
$ soffice "--accept=socket,port=2002;urp;"
```

This script expects two arguments. The first is the 'input file', the spreadsheet file to convert/export sheets from, and the 'output directory', which is where each CSV file will be stored.

```
$ python3 CsvConverter.py my_file.odt my_file_sheets
```

Upon running this command, every sheet from `my_file.odt` will be exported as a CSV file in `my_file_sheets`. Each CSV file will be named after the sheet it was converted from.

If the sheet names are inconvenient or incompatible with your filesystem, you can use `--ignore-sheet-names` which will instead name them `sheet-1.csv`, `sheet-2.csv`, and so on.

The input spreadsheet does not have to be a `.odt` file. It can be any file that LibreOffice will open as a spreadsheet.


## Flashing lights warning

This script will attempt to open a spreadsheet file in LibreOffice, cycle through each sheet as quickly as it can, then close the document.

This process is automated so it may occur very quickly, resulting in rapidly flashing lights on your screen.

I've included an optional argument `--slow` which will intentionally slow down the speed at which it moves through the spreadsheet, which will hopefully reduce the risk of uncomfortable flashing lights.

Usage with slow mode enabled:

```
$ python3 CsvConverter.py --slow my_file.odt my_file_sheets
```

Regardless, you use this script at your own risk.

