#!/usr/bin/env python3

#
# PyODCsvConverter (Python OpenDocument CSV Converter)
#
# This script converts a spreadsheet document to one or more CSV files by
# connecting to a LibreOffice instance via Python-UNO bridge.
#
# Copyright (C) 2023 Benjamin Hottell, licensed under the GNU LGPL v2.1 or any
# later version
#
# This script is heavily based on the original PyODConverter by Mirko Nasato
# (C) 2008-2012 licensed under the GNU LGPL v2.1 or any later version.
#
# Link to PyODConverter: https://github.com/mirkonasato/pyodconverter/tree/master
#
# Link to GNU LGPL v2.1: http://www.gnu.org/licenses/lgpl-2.1.html
#
# The script was modified beginning 2023 May 18 referencing this specific
# version of DocumentConverter.py:
# https://github.com/mirkonasato/pyodconverter/blob/9eb6097b5716928fef2bf3ccb0ef37ed03837704/DocumentConverter.py
#
# The changes made to DocumentConverter.py from PyODConverter are not endorsed
# by the original authors.
#

import sys
import os
import time
import argparse

import uno

from com.sun.star.beans import PropertyValue
from com.sun.star.task import ErrorCodeIOException
from com.sun.star.connection import NoConnectException



# --- configuration & constants ---

DEFAULT_OPENOFFICE_HOST = 'localhost'
DEFAULT_OPENOFFICE_PORT = 2002

# see http://wiki.services.openoffice.org/wiki/Framework/Article/Filter

# most formats are auto-detected; only those requiring options are defined here
# (taken from PyODConverter)
IMPORT_FILTER_MAP = {
    "txt": {
        "FilterName": "Text (encoded)",
        "FilterOptions": "utf8"
    },
    "csv": {
        "FilterName": "Text - txt - csv (StarCalc)",
        "FilterOptions": "44,34,0"
    }
}

CSV_EXPORT_FILTER = {
    "FilterName": "Text - txt - csv (StarCalc)",
    "FilterOptions": "44,34,0"
}



# --- reusable utility functions ---

# if path is 'file.xyz', returns 'xyz' (without the dot)
# if path does not have an extension, returns None (implicitly)
def get_file_ext(path:str):
    ext = os.path.splitext(path)[1]
    if ext is not None:
        return ext[1:].lower()

# print to stderr
def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)

def path_to_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))

def dict_to_uno_properties(d):
    props = []
    for key in d:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = d[key]
        props.append(prop)
    return tuple(props)



# --- script-specific functionality ---

class CsvConversionException(Exception):

    def __init__(self, message):
        self.message = message

    def __str__(self):
        return self.message

class LibreOfficeConnectionException(CsvConversionException):

    def __init__(self, host, port, message=None):
        if message is None:
            message = f"Failed to connect to LibreOffice at host '{host}' and port {port}"
        self.host = host
        self.port = port
        super().__init__(message)

    def __str__(self):
        return self.message

class CsvConverter:
    
    def __init__(self, host=DEFAULT_OPENOFFICE_HOST, port=DEFAULT_OPENOFFICE_PORT):
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)

        try:
            context = resolver.resolve(f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext")
        except NoConnectException:
            raise LibreOfficeConnectionException(host=host, port=port)

        self.desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

    def convert(
            self,
            input_file,
            output_dir,
            slow=False,
            keep_open=False,
            ignore_sheet_names=False):

        output_ext = 'csv'

        load_properties = {}

        input_ext = get_file_ext(input_file)

        if input_ext in IMPORT_FILTER_MAP:
            load_properties.update(IMPORT_FILTER_MAP[input_ext])
        
        document = self.desktop.loadComponentFromURL(
            path_to_url(input_file),
            "_blank", 0, dict_to_uno_properties(load_properties))

        try:
            document.refresh()
        except AttributeError:
            pass

        try:

            save_properties = CSV_EXPORT_FILTER

            doc_sheets = document.Sheets

            for sheet_no in range(doc_sheets.Count):

                if slow:
                    time.sleep(1)

                sheet = doc_sheets.getByIndex(sheet_no)

                document.CurrentController.setActiveSheet(sheet)

                if ignore_sheet_names:
                    output_name_base = f"sheet-{sheet_no+1}"
                else:
                    output_name_base = sheet.Name

                output_url = path_to_url(os.path.join(output_dir, output_name_base + '.' + output_ext))

                document.storeToURL(output_url, dict_to_uno_properties(save_properties))

        finally:
            if not keep_open:
                document.close(True)

def main():

    argp = argparse.ArgumentParser(
        description="Automatically exports sheets from LibreOffice spreadsheets into CSV files.")

    argp.add_argument("input_file", metavar='INPUT_FILE',
        help="the spreadsheet file to open in LibreOffice")

    argp.add_argument("output_dir", metavar='OUTPUT_DIR',
        help="the directory to save the resultant CSV files in")

    argp.add_argument("-P", "--port", type=int, default=DEFAULT_OPENOFFICE_PORT,
        help=f"the port the LibreOffice server is running on (default {DEFAULT_OPENOFFICE_PORT})")

    argp.add_argument("-H", "--host", type=str, default=DEFAULT_OPENOFFICE_HOST,
        help=f"the host the LibreOffice server is running on (default {DEFAULT_OPENOFFICE_HOST})")

    argp.add_argument("--slow", default=False, action='store_true',
        help="slows down the rate at which sheets are converted (may help with flashing lights)")

    argp.add_argument("--keep-open", default=False, action='store_true',
        help="do not close the document after the last sheet is converted")

    argp.add_argument("--ignore-sheet-names", default=False, action='store_true',
        help="name the resultant CSV files 'sheet-X.csv', starting at X=1")

    args = argp.parse_args()

    input_file = args.input_file
    output_dir = args.output_dir

    if not os.path.isfile(input_file):
        eprint(f"No such file: {input_file}")
        sys.exit(1)

    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    if not os.path.isdir(output_dir):
        eprint(f"Not a directory: {output_dir}")
        sys.exit(1)

    try:
        converter = CsvConverter(host=args.host, port=args.port)    
        converter.convert(
            input_file,
            output_dir,
            slow=args.slow,
            keep_open=args.keep_open,
            ignore_sheet_names=args.ignore_sheet_names)

    except LibreOfficeConnectionException as exception:
        eprint("Error: " + str(exception))
        eprint("Make sure to start a LibreOffice server before running this script.")
        eprint("Example:")
        eprint(f"$ soffice \"--accept=socket,port={args.port};urp;\"")
        sys.exit(1)

    except CsvConversionException as exception:
        eprint("Error: " + str(exception))
        sys.exit(1)

    except ErrorCodeIOException as exception:
        eprint("Error: ErrorCodeIOException " + str(exception.ErrCode))
        sys.exit(1)


if __name__ == "__main__":
    main()

