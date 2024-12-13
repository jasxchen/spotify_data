#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Dec  6 22:46:42 2024

@author: Jasmine Chen
"""

import pandas as pd
from datetime import datetime
from collections import Counter

INPUT_EXCEL_FILE = "spotify_data.xlsx"
OUTPUT_EXCEL_FILE = "spotify_analysis.xlsx"

current_datetime = datetime.now()
current_month = current_datetime.strftime("%B")
current_month = current_datetime.year

def load_data(file_path):
    '''
    Load Spotify data from an Excel file.

    Parameters
    ----------
    file_path : str
        Path to the Excel File

    Returns
    -------
    DataFrame containing the Spotify data
    
    '''
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.lower()
    required_columns = {'date', 'artist', 'song'}
    if not required_columns.issubset(df.columns):
        raise ValueError(f"The input file must contain the following columns: {required_columns}")
    return df

def filter_data_by_year(df, year):
    '''
    Filter the data for a specific year.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame containing Spotify data
    year : int
        The year to filter by

    Returns
    -------
    Filtered DataFrame

    '''
    return df[df['date'].str.contains(str(year), na=False)]

def analyze_monthly_counts(df):
    '''
    Calculate the number of songs played each month.

    Parameters
    ----------
    df : df.DataFrame
        DataFrame containing Spotify data

    Returns
    -------
    Series with monthly counts

    '''
    return df['date'].str.extract(r'(\w+)')[0].value_counts().sort_index()

def get_top_items(counts, top_n=10):
    '''
    Get the top N items from a Counter.

    Parameters
    ----------
    counts : TYPE
        Counter object containing item counts.
    top_n : TYPE, optional
        Number of top items to retrieve. The default is 10.

    Returns
    -------
    Dictionary of top N items

    '''
    return dict(sorted(counts.items(), key=lambda x: x[1], reverse=True)[:top_n])

def main():
    try: 
        df = load_data(INPUT_EXCEL_FILE)
        
        current_year = datetime.now().year
        filtered_data = filter_data_by_year(df, current_year)
        
        monthly_counts = analyze_monthly_counts(filtered_data)
        artist_counts = Counter(filtered_data['artist'])
        song_counts = Counter(filtered_data['song'])
        
        top_artists = get_top_items(artist_counts, top_n=10)
        top_songs = get_top_items(song_counts, top_n=10)
        
        unique_artists = [artist for artist, count in artist_counts.items() if count == 1]
        unique_songs = [song for song, count in song_counts.items() if count == 1]
        
        summary = {
            "Total Songs Played": len(filtered_data),
            "Total Different Artists": len(artist_counts),
            "Total Different Songs": len(song_counts),
            "Artists with One Play": len(unique_artists),
            "Songs with One Play": len(unique_songs)
        }
        
        # save_to_excel(monthly_counts, top_artists, top_songs, unique_artists, unique_songs, summary, OUTPUT_EXCEL_FILE)
        # print(f"Analysis saved to {OUTPUT_EXCEL_FILE}")
    
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
        




















