# encoding: utf-8

import xlrd
import os
import sys
import json


# This script takes all excel files (yearly reports) from
# the source folder and outputs the data in structured
# JSON to the specified location.

SOURCE_FOLDER = 'data/source/umweltbundesamt/no2'
DESTINATION_FILE = 'data/refined/umweltbundesamt/no2.json'

# some assumptions on the content of the found data
DATA_ASSUMPTIONS_AND_MAPPING = {
    'year_2001': {
        'values_start_in_row': 43,  # starting to count with 0
        'column_headers_in_row': 41,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            3: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 290 \xb5g/m\xb3",
                'field': 'hours_exceeding_290',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2002': {
        'values_start_in_row': 43,  # starting to count with 0
        'column_headers_in_row': 41,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            3: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 280 \xb5g/m\xb3",
                'field': 'hours_exceeding_280',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2003': {
        'values_start_in_row': 42,  # starting to count with 0
        'column_headers_in_row': 40,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            3: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 270 \xb5g/m\xb3",
                'field': 'hours_exceeding_270',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2004': {
        'values_start_in_row': 41,  # starting to count with 0
        'column_headers_in_row': 39,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            3: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 260 \xb5g/m\xb3",
                'field': 'hours_exceeding_260',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2005': {
        'values_start_in_row': 43,  # starting to count with 0
        'column_headers_in_row': 41,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            3: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 250 \xb5g/m\xb3",
                'field': 'hours_exceeding_250',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2006': {
        'values_start_in_row': 34,  # starting to count with 0
        'column_headers_in_row': 32,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            3: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 240 \xb5g/m\xb3",
                'field': 'hours_exceeding_240',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2007': {
        'values_start_in_row': 34,  # starting to count with 0
        'column_headers_in_row': 32,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            3: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 230 \xb5g/m\xb3",
                'field': 'hours_exceeding_230',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2008': {
        'values_start_in_row': 34,  # starting to count with 0
        'column_headers_in_row': 32,
        'columns': {
            0: {
                'headline': u"Station",
                'field': 'station_id'
            },
            1: {
                'headline': u'Name/Messnetz',
                'field': None
            },
            2: {
                'headline': u'Emissinsquelltyp',
                'field': None
            },
            3: {
                'headline': u'Umgebungstyp',
                'field': None
            },
            4: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'yearly_average'
            },
            5: {
                'headline': u"in \xb5g/m\xb3",
                'field': 'max_hourly'
            },
            6: {
                'headline': u"> 220 \xb5g/m\xb3",
                'field': 'hours_exceeding_220',
                'type': 'int'
            },
            7: {
                'headline': u"> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_200',
                'type': 'int'
            }
        }
    },
    'year_2009': {
        'values_start_in_row': 42,  # starting to count with 0
        'column_headers_in_row': 40,
        'columns': {
            0: {
                'headline': u"Stations-code",
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
                'headline': u"Zahl* der Stundenwerte > 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_220',
                'type': 'int'
            },
            6: {
                'headline': u"Zahl* der Stundenwerte > 210 \xb5g/m\xb3",
                'field': 'hours_exceeding_210',
                'type': 'int'
            }
        }
    },
    'year_2010': {
        'values_start_in_row': 41,  # starting to count with 0
        'column_headers_in_row': 39,
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
                'headline': u"Zahl* der Stundenwerte > 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_220',
                'type': 'int'
            }
        }
    },
    'year_2011': {
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
                'headline': u"Zahl* der Stundenwerte \n> 200 \xb5g/m\xb3",
                'field': 'hours_exceeding_220',
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
                    if val == '' or val == '---':
                        val = None
                    if val != None:
                        if 'type' in mapping['columns'][n]:
                            if mapping['columns'][n]['type'] == 'int':
                                val = int(val)
                    dataset[mapping['columns'][n]['field']] = val
            if dataset['station_id'] is not None:
                dataset['station_id'] = dataset['station_id'].replace('*', '')
                dataset['station_id'] = dataset['station_id'].replace("'", "")
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
    # aggregate per station
    stationdata = {}
    for year in data:
        for entry in data[year]:
            #print entry
            if entry['station_id'] not in stationdata:
                stationdata[entry['station_id']] = {}
            for key in entry:
                if key == 'station_id':
                    continue
                if key not in stationdata[entry['station_id']]:
                    stationdata[entry['station_id']][key] = {}
                stationdata[entry['station_id']][key][year] = entry[key]
    f = open(DESTINATION_FILE, 'w')
    f.write(json.dumps(stationdata, indent=4, sort_keys=True))
    f.close()

