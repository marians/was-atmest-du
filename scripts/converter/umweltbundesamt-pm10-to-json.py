# encoding: utf-8

import xlrd
import os
import sys
import json


# This script takes all excel files (yearly reports) from
# the source folder and outputs the data in structured
# JSON to the specified location.

SOURCE_FOLDER = 'data/source/umweltbundesamt/pm10'
DESTINATION_FILE = 'data/refined/umweltbundesamt/pm10.json'


def get_certain_files(path, valid_extensions):
    files = []
    listing = os.listdir(path)
    for infile in listing:
        extension = os.path.splitext(infile)[1].split('.')[1]
        if extension in valid_extensions:
            files.append(infile)
    return files


def get_year_data(filename):
    # assumption: file name is like PM10_<year>.xls[x]
    year = int(os.path.splitext(filename)[0].split('_')[1])
    print year
    book = xlrd.open_workbook(SOURCE_FOLDER + os.sep + filename)
    sheet = book.sheet_by_index(0)
    for row in range(sheet.nrows):
        #print sheet.row(row)
        pass

if __name__ == '__main__':
    files = get_certain_files(SOURCE_FOLDER, ['xls', 'xlsx'])
    for f in files:
        yeardata = get_year_data(f)
