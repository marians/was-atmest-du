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

# some assumptions on the content of the found data
DATA_ASSUMPTIONS_AND_MAPPING = {
    'year_2002': {
        'values_start_in_row': 39,  # starting to count with 0
        'column_headers_in_row': 37,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-\nmittelwert\nin \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der \nTageswerte \n> 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            },
            6: {
                'headline': u"Zahl der\nTageswerte\n > 65 \xb5g/m\xb3",
                'field': 'days_exceeding_65',
                'type': 'int'
            }
        }
    },
    'year_2003': {
        'values_start_in_row': 39,  # starting to count with 0
        'column_headers_in_row': 37,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-\nmittelwert \nin \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der \nTageswerte \n> 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            },
            6: {
                'headline': u"Zahl der \nTageswerte \n> 60 \xb5g/m\xb3",
                'field': 'days_exceeding_60',
                'type': 'int'
            }
        }
    },
    'year_2004': {
        'values_start_in_row': 39,  # starting to count with 0
        'column_headers_in_row': 37,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': None
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-\nmittelwert\nin \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der\n Tageswerte \n> 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            },
            6: {
                'headline': u"Zahl der \nTageswerte \n> 55 \xb5g/m\xb3",
                'field': 'days_exceeding_55',
                'type': 'int'
            }
        }
    },
    'year_2005': {
        'values_start_in_row': 39,  # starting to count with 0
        'column_headers_in_row': 37,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-mittelwert in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der Tageswerte > 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    },
    'year_2006': {
        'values_start_in_row': 39,  # starting to count with 0
        'column_headers_in_row': 37,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-mittelwert in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der Tageswerte > 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    },
    'year_2007': {
        'values_start_in_row': 38,  # starting to count with 0
        'column_headers_in_row': 36,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-mittelwert in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der Tageswerte > 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    },
    'year_2008': {
        'values_start_in_row': 39,  # starting to count with 0
        'column_headers_in_row': 37,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-mittelwert in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der Tageswerte > 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    },
    'year_2009': {
        'values_start_in_row': 42,  # starting to count with 0
        'column_headers_in_row': 40,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-mittelwert in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der Tageswerte > 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    },
    'year_2010': {
        'values_start_in_row': 31,  # starting to count with 0
        'column_headers_in_row': 29,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-mittelwert in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der \nTageswerte \n> 50 \xb5g/m\xb3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    },
    'year_2011': {
        'values_start_in_row': 34,  # starting to count with 0
        'column_headers_in_row': 32,
        'columns': {
            0: {
                'headline': u"Stationscode",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name / Messnetz',
                'field': None
            },
            2: {
                'headline': u'Stationsumgebung',
                'field': None
            },
            3: {
                'headline': u'Art der Station',
                'field': None
            },
            4: {
                'headline': u"Jahres-\nmittelwert \nin \xb5g/m3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"Zahl der Tageswerte \n> 50 \xb5g/m3",
                'field': 'days_exceeding_50',
                'type': 'int'
            }
        }
    }
}


def get_certain_files(path, valid_extensions):
    files = []
    listing = os.listdir(path)
    for infile in listing:
        extension = os.path.splitext(infile)[1].split('.')[1]
        if extension in valid_extensions:
            files.append(infile)
    return files


def get_year_data(year, filename):
    data = []
    assumptions_key = 'year_' + str(year)
    if assumptions_key in DATA_ASSUMPTIONS_AND_MAPPING:
        mapping = DATA_ASSUMPTIONS_AND_MAPPING[assumptions_key]
        book = xlrd.open_workbook(SOURCE_FOLDER + os.sep + filename)
        sheet = book.sheet_by_index(0)
        # check headers
        headers = sheet.row(mapping['column_headers_in_row'])
        for n in range(0, len(headers)):
            header_text = sheet.cell_value(mapping['column_headers_in_row'], n)
            if header_text != mapping['columns'][n]['headline']:
                print "ERROR: Year=" + str(year) + ", found this header for column " + str(n) + ":", [header_text]
                print "       Expected:", [mapping['columns'][n]['headline']]
                sys.exit(1)
        # read data cells
        for row in range(mapping['values_start_in_row'], sheet.nrows):
            dataset = {}
            for n in mapping['columns']:
                if mapping['columns'][n]['field'] is not None:
                    #print "row", row, ", col", n
                    val = sheet.cell_value(row, n)
                    if val == '':
                        val = None
                    if val != None:
                        if 'type' in mapping['columns'][n]:
                            if mapping['columns'][n]['type'] == 'int':
                                val = int(val)
                    dataset[mapping['columns'][n]['field']] = val
                    data.append(dataset)
        return data
    else:
        print "ERROR: No mapping configuration for year " + str(year) + " available. Please add to DATA_ASSUMPTIONS_AND_MAPPING."
        sys.exit(1)

if __name__ == '__main__':
    files = get_certain_files(SOURCE_FOLDER, ['xls', 'xlsx'])
    data = {}
    for f in files:
        # assumption: file name is like PM10_<year>.xls[x]
        year = int(os.path.splitext(f)[0].split('_')[1])
        data[year] = get_year_data(year, f)
    #print json.dumps(data, indent=4)
    f = open(DESTINATION_FILE, 'w')
    f.write(json.dumps(data, indent=4, sort_keys=True))
    f.close()
