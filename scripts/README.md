# Scripts

This directory contains Python scripts developed for the renal cancer survival analysis workflow.

## Current Scripts

### encori_scraper.py

This script automates the retrieval of survival analysis information from the ENCORI database.

### Features

- Reads candidate lncRNAs from Excel worksheets
- Queries ENCORI automatically using Selenium
- Retrieves Kaplan–Meier survival information
- Extracts hazard ratios and log-rank p-values
- Saves screenshots for each analysed gene
- Writes survival statistics back to Excel

## Requirements

- Python 3.x
- Selenium
- Pandas
- OpenPyXL
- Chrome WebDriver
