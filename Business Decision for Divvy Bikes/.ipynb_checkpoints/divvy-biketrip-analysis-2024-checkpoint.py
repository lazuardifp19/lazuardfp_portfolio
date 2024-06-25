#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 20:18:25 2024

@author: lazuardifp
"""

import pandas as pd

jan_2024 = pd.read_csv('202401-divvy-tripdata.csv')
feb_2024 = pd.read_csv('202402-divvy-tripdata.csv')
mar_2024 = pd.read_csv('202403-divvy-tripdata.csv')
apr_2024 = pd.read_csv('202404-divvy-tripdata.csv')
may_2024 = pd.read_csv('202405-divvy-tripdata.csv')

df = pd.concat([jan_2024, feb_2024, mar_2024, apr_2024, may_2024])
