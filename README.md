
# Checkin-Checkout (CICO) Report Exporter

## Overview
This Python script retrieves timestamps from an Excel file containing swipe records and exports them to a new Excel file. It allows you to filter the data by person's name, year, and month.

## Prerequisites
- Python 3.x
- xlrd library for reading Excel files
- openpyxl library for writing Excel files

## Installation
1. Install Python 3.x from the official Python website.
2. Create a Python virtual environment and activate it:
        
```bash
python -m venv .venv
source .v
```

3. Install the required packages:
    
```bash
pip install -r requirements.txt
```


## Usage
- Run the script with the following command:
```bash
python cicoReportExporter.py [--year YEAR] [--month MONTH] [--output OUTPUT_FILE] [--name NAME]
```

## Arguments
- `--year, -y`: Specify the year for filtering the data (optional).
- `--month, -m`: Specify the month for filtering the data (optional).
- `--output, -o`: Specify the output file name (default is "cicoReport").
- `--name, -n`: Specify the name of the person for filtering the data (optional).

## Examples
### Example 1
Export data for all persons in April 2024 to a file named "cico_report":

```bash
python cicoReportExporter.py --year 2024 --month April
```
### Example 2
Export data for a specific person named "John" in May 2024 to a file named "john_report":

```bash
python cicoReportExporter.py --year 2024 --month May --name John --output john_report
```

### License
This project is licensed under the MIT License. See the LICENSE file for details.