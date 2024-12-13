#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 12 20:01:11 2024

@author: Jasmine Chen
"""

import pandas as pd
import time

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1xE_o-a6Pz7Lx5jAKmjq04UZ4LBjEGoqWK-nJa3Wcvtg/edit?gid=0#gid=0"
LOCAL_EXCEL_FILE = "spotify_data.xlsx"
SYNC_INTERVAL = 3600

def fetch_google_sheet(sheet_url):
    '''
    Fetches data from a Google Sheet.

    Parameters
    ----------
    sheet_url : str
        Public URL of the Google Sheet

    Returns
    -------
    DataFrame containing the data

    '''
    google_sheets_csv_url = sheet_url.replace("/edit#gid=", "/export?format=csv&gid=")
    
    print("Fetching data from Google Sheet...")
    return pd.read_csv(google_sheets_csv_url)

def append_to_excel(new_data, excel_file):
    '''
    Appends new Spotify data to the local Excel file.

    Parameters
    ----------
    new_data : df.DataFrame
        DataFrame containing new data
    excel_file : str
        Path to the local Excel file

    '''
    try:
        existing_data = pd.read_excel(excel_file)
    except FileNotFoundError:
        existing_data = pd.DataFrame()
    
    combined_data = pd.concat([existing_data, new_data]).drop_duplicates()
    
    combined_data.to_excel(excel_file, index=False)
    print(f"Data appended to {excel_file}. Total rows: {len(combined_data)}")
    
def main():
    try:
        new_data = fetch_google_sheet(GOOGLE_SHEET_URL)
        
        append_to_excel(new_data, LOCAL_EXCEL_FILE)
    
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    while True:
        main()
        print(f"Sync completed. Waiting {SYNC_INTERVAL} seconds for the next sync.")
        time.sleep(SYNC_INTERVAL)



















