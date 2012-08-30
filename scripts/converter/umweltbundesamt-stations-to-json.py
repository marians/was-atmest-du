# encoding: utf-8

"""
This script converts measurement station data from the
Umweltbundesamt (CSV data) to one JSON file
"""

import csv
import json

STATION_CSV_FILE = 'data/source/umweltbundesamt/Bericht_EU_Meta_Stationen.csv'
STATION_PARAMS_CSV_FILE = 'data/source/umweltbundesamt/Bericht_EU_Meta_Stationsparameter.csv'
DESTINATION_FILE = 'data/refined/umweltbundesamt/stations.json'


def read_station_list(path):
    csvreader = csv.reader(open(path, 'rb'), delimiter=';')
    rowcount = -1
    headers = []
    data = {}
    for row in csvreader:
        rowcount += 1
        if rowcount == 0:
            continue
        if rowcount == 1:
            headers = row
        if rowcount > 1:
            dataset = {}
            for n in range(0, len(headers)):
                if n < len(row):
                    val = row[n]
                    if val == '':
                        val = None
                    if headers[n] in ['station_latitude_dms',
                        'station_longitude_dms', '']:
                        continue
                    if val is not None:
                        if headers[n] == 'station_altitude':
                            val = int(val)
                        elif headers[n] in ['station_latitude_d',
                            'station_longitude_d']:
                            val = float(val)
                        else:
                            val = val.decode('cp1252')
                    dataset[headers[n]] = val
            code = dataset['station_code']
            del dataset['station_code']
            data[code] = dataset
    return data

def read_stations_params(path):
    csvreader = csv.reader(open(path, 'rb'), delimiter=';')
    rowcount = -1
    headers = []
    data = {}
    for row in csvreader:
        rowcount += 1
        if rowcount == 0:
            continue
        if rowcount == 1:
            headers = row
        if rowcount > 1:
            dataset = {}
            for n in range(0, len(headers)):
                if n < len(row):
                    val = row[n]
                    if val in ['', 'n.a.', 'unknown']:
                        val = None
                    if val is not None:
                        if headers[n] in ['component_code']:
                            val = int(val)
                        else:
                            val = val.decode('cp1252')
                    dataset[headers[n]] = val
            # append to data keyed by station code
            code = dataset['station_code']
            if code not in data:
                data[code] = []
            del dataset['station_code']
            data[code].append(dataset)
    return data


def create_output_file(stations, params, path):
    for station_id in stations:
        stations[station_id]['parameters'] = None
        if station_id in params:
            #print params[station_id]
            stations[station_id]['parameters'] = params[station_id]
    f = open(DESTINATION_FILE, 'w')
    f.write(json.dumps(stations, indent=4, sort_keys=True))
    f.close()



if __name__ == '__main__':
    stations_raw = read_station_list(STATION_CSV_FILE)
    stations_params = read_stations_params(STATION_PARAMS_CSV_FILE)
    create_output_file(stations_raw, stations_params, DESTINATION_FILE)
