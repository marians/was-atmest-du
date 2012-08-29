was-atmest-du
=============

Air quality data display for Germany

This is an open repository for data on air quality in Germany plus, maybe in the future, more refined versions of that data and scripts to automate the refinement process.

##data

This is where data lives

##data/refined

Data that has been dealt with in some way goes here.

##data/refined/umweltbundesamt/pm10.json

This is the raw content of the Excel files data/source/umweltbundesamt/pm10 (PM10 annual reports) in structured form.

##data/source

Here is where original data as copied from providers is placed

##data/source/umweltbundesamt

Data from the Umweltbundesamt, Germany's federal environment agency

* Bericht_EU_Meta_Stationen.csv: List of stations, including historic data. Source: http://www.env-it.de/stationen/public/downloadRequest.do
* Bericht_EU_Meta_Stationsparameter.csv: List of stations with additional information, including historic data. Source: http://www.env-it.de/stationen/public/downloadRequest.do

##data/source/umweltbundesamt/pm10

Data on particle emissions finer than 10 µm. Files originate from http://www.env-it.de/umweltbundesamt/luftdaten/documents.fwd

Each Excel file contains yearly average data per station.

##data/source/umweltbundesamt/no2

Data on nitrogen dioxide emissions (Stickstoffdioxid). Files originate from http://www.env-it.de/umweltbundesamt/luftdaten/documents.fwd

Each Excel file contains yearly average values plus the count of measures above a certain threshold per station.


##scripts

This should contain Python tools to work with the data.

Here is a flexible way to handle the requirements.

1. Install virtualenv
2. Go to the top folder of your local clone of this repository and do:

     virtualenv venv
     . venv/bin/activate
     pip install xlrd

Now whenever you want to run one of the scripts, start your virtual environment first using

    . venv/bin/activate

##scripts/converter

Scripts converting data from one format to another

##scripts/converter/umweltbundesamt-pm10-to-json.py

Extracts data from the Excel files in data/source/umweltbundesamt/pm10 to data/refined/umweltbundesamt/pm10.json.