#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 20:18:25 2024

@author: lazuardifp
"""

import pandas as pd

# Load the data 
jan_2024 = pd.read_csv('202401-divvy-tripdata.csv')
feb_2024 = pd.read_csv('202402-divvy-tripdata.csv')
mar_2024 = pd.read_csv('202403-divvy-tripdata.csv')
apr_2024 = pd.read_csv('202404-divvy-tripdata.csv')
may_2024 = pd.read_csv('202405-divvy-tripdata.csv')

# Combine the data frame
df = pd.concat([jan_2024, feb_2024, mar_2024, apr_2024, may_2024])

# Check the amount of null values in each column
df.isnull().sum()

# Delete the null values
df.dropna(inplace = True)
df.isnull().sum()

# Remove columns that are not necessary for the analysis
df = df.drop(['start_station_name', 'start_station_id', 'end_station_name', 
              'end_station_id', 'start_lat', 'start_lng', 'end_lat', 'end_lng'], axis = 1)

df.duplicated().sum()
df.drop_duplicates()

df["started_at"] = pd.to_datetime(df["started_at"])
df["ended_at"] = pd.to_datetime(df["ended_at"])

df.info()

# Create new columns to extract hour, date, day, and month
df['hour'] = df['started_at'].dt.hour
df['date'] = df['started_at'].dt.day
df['day'] = df['started_at'].dt.day_name()
df['month'] = df['started_at'].dt.month_name()

day_ordered = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 
               'Saturday', 'Sunday']

df['day'] = pd.Categorical(df['day'], categories=day_ordered, ordered=True)

month_ordered = ['January', 'February', 'March', 'April', 'May', 'June',
                 'July', 'August', 'September', 'October', 'November', 'December']

df['month'] = pd.Categorical(df['month'], categories=month_ordered, ordered=True)

