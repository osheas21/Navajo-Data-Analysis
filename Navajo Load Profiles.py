# Set current working directory to the folder containing this file
import os
os.chdir(os.path.dirname(os.path.realpath(__file__)))

import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from sklearn.cluster import KMeans

from Navajo_Load_Profiles_Functions import *


folders = create_folders()

data_filepaths = [f"{folders['Energy_Data_New']}\\{filename}" for filename \
    in os.listdir(folders['Energy_Data_New']) \
    if os.path.isfile(f"{folders['Energy_Data_New']}\\{filename}")]
data_filepaths = [filepath for filepath in data_filepaths if '2021' in filepath or '2022' in filepath]

battery_analysis_filepaths = [f"{folders['Battery_Analysis']}\\{filename}" for filename \
    in os.listdir(folders['Battery_Analysis']) \
    if os.path.isfile(f"{folders['Battery_Analysis']}\\{filename}")]

# battery_data[location][year]
battery_data = extract_battery_data(battery_analysis_filepaths)

# Create a DataFrame with averaged load profiles for each location and year.
# Rows are indexed with the 10-minute period number (0-144) and columns are
# multiindexed by location then year
avg_load_profiles_df = agg_avg_load_profiles(data_filepaths, MAX_NUM_MISSING_VALS, battery_data)
avg_load_profiles_df.to_csv(f"{folders['Aggregated Load Profiles']}\\Average Load Profiles.csv")

# Create a DataFrame with averaged load profiles for each location and year for
# both weekday and weekend data. Rows are indexed with the 10-minute period
# number (0-144) and columns are multiindexed by location then year
weekday_df, weekend_df = calc_load_profile_weekday_weekend(data_filepaths, MAX_NUM_MISSING_VALS, battery_data)
weekday_df.to_csv(f"{folders['Aggregated Load Profiles']}\\Avg Weekday Load Profiles.csv")
weekend_df.to_csv(f"{folders['Aggregated Load Profiles']}\\Avg Weekend Load Profiles.csv")

# Create a DataFrame with averaged load profiles for each location, year, and
# season. Rows are indexed with the 10-minute period number (0-144) and columns
# are multiindexed by location then year then season.
avg_seasonal_load_profiles_df = agg_avg_seasonal_load_profiles(data_filepaths, MAX_NUM_MISSING_VALS, battery_data)
avg_seasonal_load_profiles_df.to_csv(f"{folders['Aggregated Load Profiles']}\\Average Seasonal Load Profiles.csv")

# Create boxplots
print("Creating Boxplots...", end='')
for filepath in data_filepaths:
    location, metadata = extract_metadata(filepath)
    data_df, other_data = extract_power_data(filepath)
    year = metadata['yr_start']
    print()

    # Remove data with more than 18 missing values.
    data_df = preprocess_df(data_df, other_data, battery_data[location][year]['Outage Flag'], MAX_NUM_MISSING_VALS)
    print(data_df)


    year = metadata['yr_start']
    boxplot_load_profile(data_df, location, year, save_folder=folders['Load Profile Boxplots'])
    boxplot_load_profile_by_season(data_df, location, year, save_folder=folders['Seasonal Boxplots'])
    boxplot_load_profile_weekday_weekend(data_df, location, year, save_folder=folders['Weekday-Weekend Load Profile Boxplots'])
print("done.")