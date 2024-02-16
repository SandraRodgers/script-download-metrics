# CSV Download Automation

This Python script automates downloading a CSV report from a web dashboard and appending it to a master Excel file.

## Packages Used

- Selenium
- Pandas
- Requests

## Overview

The script does the following:

- Logs into the web dashboard using credentials loaded from .env
  Navigates to the CSV export page
- Clicks the export button to download the CSV file
- Checks for an existing master.xlsx file. If it doesn't exist, creates a new Excel file. If it exists, appends the CSV data into a new sheet
- Names the sheet based on the date in the CSV data Usage

## Usage

1. Clone the repo
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Add your credentials to .env

4. Run the script:

```bash
python main.py
```

5. Find downloaded CSVs in /csv_exports and master.xlsx in root folder

## Customization

- Update the login and CSV URL in main.py
- Adjust the download directory, Excel filename, etc.

## Purpose

Automating the CSV download process allows updated data to be captured on a schedule into a central Excel dashboard for reporting and analysis.

This saves manual time of downloading the files and combining them each period.
