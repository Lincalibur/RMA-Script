# Missing RMA Tracker

This project provides a Python script to manage and update a list of missing RMA (Return Merchandise Authorization) records by searching through folders containing tracking and IMEI (International Mobile Equipment Identity) information. The script updates the missing list by removing found entries and saving the updated lists to Excel files.

## Summary

The script performs the following main tasks:
1. Loads a list of missing RMAs from an Excel file.
2. Searches for tracking and IMEI numbers in specified folders.
3. Updates the missing list by removing found entries.
4. Saves the updated missing list and found entries to new Excel files.
5. Merges found entries with additional data and saves the results.

## Prerequisites

Ensure you have the following installed:
- Python 3.x
- pandas
- openpyxl
- xlsxwriter

You can install the required Python packages using pip:
```sh
pip install pandas openpyxl xlsxwriter
