import pandas as pd
import numpy as np
import os
import ast
from matplotlib import pyplot as plt


MAX_NUM_MISSING_VALS = 6 * 3    # 2 hours * 6 10min intervals
MAX_NUM_ZEROS = 144 # allow MAX_NUM_ZEROS zero values before data is removed

# Used to turn quarter numbers into quarter names
quarter_season_dict = {
    1: 'Jan-Mar',
    2: 'Apr-Jun',
    3: 'Jul-Sep',
    4: 'Oct-Dec'
}


# Create a dictionary of folders while also creating new directories in the
# host computer's file directory if those folders are not already present.
def create_folders():
    folders = {}
    folders['Energy_Data_New'] = "Energy_Data_New"
    folders['Battery_Analysis'] = "Battery_Analysis"
    folders['Load Profile Boxplots'] = "Load Profile Boxplots"
    folders['Weekday-Weekend Load Profile Boxplots'] = "Weekday-Weekend Load Profile Boxplots"
    folders['Seasonal Boxplots'] = "Seasonal Boxplots"
    folders['Aggregated Load Profiles'] = "Aggregated Load Profiles"

    for folder_path in folders.values():
        if not os.path.exists(folder_path):
            os.mkdir(folder_path)

    return folders


# Reads in the battery data for each filepath provided
# Create a triple-nested dictionary that can be accessed in the form of
# dict[location][year][data] where location is the location number, year is the
# year 
def extract_battery_data(filepaths):
    filepaths = [filepath for filepath in filepaths if '.xlsx' in filepath]
    energy_balance_dict = {}
    count = 0   # used to keep track of progress
    for filepath in filepaths:
        # extract the metadata from the battery data file
        location, meta_dict = extract_metadata(filepath)

        # print progress
        count += 1
        print(f"Extracting Battery Data: {location} | Progress = {int(count/len(filepaths)*100)}%")

        # read battery data file
        energy_balance_df = pd.read_excel(filepath, sheet_name='Energy Balance', index_col=[0], header=None)

        # Parse battery DataFrame into a dictionary of metadata for each
        # subtable within the DataFrame. Store this new dict in the nested
        # dicts for the location and year.
        try:
            energy_balance_dict[location][meta_dict['yr_start']] = parse_other_data(energy_balance_df)
        except KeyError:
            energy_balance_dict[location] = {}
            energy_balance_dict[location][meta_dict['yr_start']] = parse_other_data(energy_balance_df)

    return energy_balance_dict


# Read in the location and other metadata (such as start and end date) from
# Sheet 1 of the Excel file at the filepath in the input.
def extract_metadata(filepath):
    meta_df = pd.read_excel(filepath, index_col=[0], sheet_name='Meta')
    location = int(meta_df.loc[0].values[0])
    meta_dict = ast.literal_eval(meta_df.loc[1].values[0])

    return location, meta_dict


# Return a DataFrame of the table of 10-minute power data from the Excel files
# with the energy usage data
def extract_power_data(filepath):
    excel_df = pd.read_excel(filepath, index_col=[0], sheet_name='Sheet1')
    data_df = excel_df.iloc[:144]
    data_df.index = data_df.index.rename('Ten-Minute Index')
    data_df = data_df.astype(float)

    other_data = parse_other_data(excel_df)

    return data_df, other_data
    

# Get the semi-meta data from the subtables underneath the main data table,
# store these data as seperate pandas Series' in other_data where the keys of
# other_data are the names of the subtables and the values of other_data are
# the Series' themselves.
def parse_other_data(excel_df):
    other_data = {}
    # 145 == first row of metadata
    # for i in range(145, len(excel_df.index)):
    for i in range(len(excel_df.index)):
        # if value is a string and not int or float, which indicates a label.
        table_name = excel_df.iloc[i].name
        if isinstance(table_name, str) and table_name != 'Ten-Minute Average Power (W)':
            other_data[table_name] = pd.Series(excel_df.iloc[i+1].values,
                                               index=excel_df.iloc[i].values,
                                               name=table_name)

    return other_data


# For a DataFrame that is indexed by hour of the day, shift the index value
# by hour_offset
def shift_hourly(df, hour_offset=-6):
    # add offset
    df.index = df.index.map(
        lambda x: x + hour_offset + 24 if x + hour_offset < 0 else x + hour_offset)
    df = df.sort_index()

    return df


# Create a boxplot of the load profile for each ten-minute period of the day,
# then save the figure to the specified save_folder.
def boxplot_load_profile(data_df, location, yearnum, save_folder):
    savefig_filepath = f"{save_folder}\\{location} Load Profile Boxplot {yearnum}.png"
    print(f"Creating and saving boxplot for {savefig_filepath}...", end='')

    xtick_freq = 4
    xtick_locs = np.asarray([x + 1 for x in list(data_df.index)[::xtick_freq]])
    data_df.T.boxplot(whis=(5, 95), sym='r+', figsize=(25, 10)).set_xticks(xtick_locs);
    plt.title(f"{location}_{yearnum}", fontsize='xx-large');
    plt.xlabel('Ten-Minute Index', fontsize='x-large');
    plt.ylabel('Load Profile (W)', fontsize='x-large');
    plt.savefig(savefig_filepath);
    plt.clf();

    print("Saved!")


# Create a boxplot of the load profile for each ten-minute period of the day,
# then save the figure to the specified save_folder.
def boxplot_load_profile_weekday_weekend(data_df, location, yearnum, save_folder):
    weekday_vals = pd.to_datetime([f"{yearnum}{x.split(' ')[0]}" for x in data_df.columns], format='%Y%j').weekday < 5

    # Series where index is df column name, values are True if that index is a weekday
    weekday_info = pd.Series(weekday_vals, name='Weekday Index', index=data_df.columns)

    weekday_columns = weekday_info.loc[weekday_info == True].index
    weekend_columns = weekday_info.loc[weekday_info == False].index

    weekday_df = data_df[data_df.columns.intersection(weekday_columns)]
    weekend_df = data_df[data_df.columns.intersection(weekend_columns)]

    for day_type, df in zip(['Weekday', 'Weekend'], [weekday_df, weekend_df]):
        print(df)
        savefig_filepath = f"{save_folder}\\{location} Load Profile Boxplot {yearnum} {day_type}.png"
        print(f"Creating and saving boxplot for {savefig_filepath}...", end='')

        xtick_freq = 4
        xtick_locs = np.asarray([x + 1 for x in list(df.index)[::xtick_freq]])
        df.T.boxplot(whis=(5, 95), sym='r+', figsize=(25, 10)).set_xticks(xtick_locs);
        plt.title(f"{location}_{yearnum}_{day_type}", fontsize='xx-large');
        plt.xlabel('Ten-Minute Index', fontsize='x-large');
        plt.ylabel('Load Profile (W)', fontsize='x-large');
        plt.savefig(savefig_filepath);
        plt.clf();

        print("Saved!")


# Create a boxplot of the load profile for each ten-minute period of the day,
# then save the figure to the specified save_folder.
def boxplot_load_profile_by_season(data_df, location, yearnum, save_folder):
    quarter_season_dict = {
        1: 'Jan-Mar',
        2: 'Apr-Jun',
        3: 'Jul-Sep',
        4: 'Oct-Dec'
    }

    # Turn column headers into a series of datetime values, then convert that
    # of datetime values into a series of values 0-3 that represent which
    # quarter of the year the day is in
    doy_vals = pd.to_datetime([f"{yearnum}{x.split(' ')[0]}" for x in data_df.columns], format='%Y%j')
    quarter_vals = doy_vals.quarter

    # For each season...
    for quarter_num, quarter_name in quarter_season_dict.items():
        quarter_info = quarter_vals == quarter_num  # bool index of loc of cols to keep
        
        # Series where col name gets bool value
        quarter_columns = pd.Series(quarter_info, name='Quarter Index', index=data_df.columns)
        quarter_cols_info = quarter_columns.loc[quarter_columns == True].index

        quarter_df = data_df[data_df.columns.intersection(quarter_cols_info)]

        print(quarter_df)
        savefig_filepath = f"{save_folder}\\{location} Load Profile Boxplot {yearnum} {quarter_name}.png"
        print(f"Creating and saving boxplot for {savefig_filepath}...", end='')

        xtick_freq = 4
        xtick_locs = np.asarray([x + 1 for x in list(quarter_df.index)[::xtick_freq]])
        quarter_df.T.boxplot(whis=(5, 95), sym='r+', figsize=(25, 10)).set_xticks(xtick_locs);
        plt.title(f"{location}_{yearnum}_{quarter_name}", fontsize='xx-large');
        plt.xlabel('Ten-Minute Index', fontsize='x-large');
        plt.ylabel('Load Profile (W)', fontsize='x-large');
        plt.savefig(savefig_filepath);
        plt.clf();

        print("Saved!")
    print()


# Convert exclude days with too many missing values, convert -8888 to nan,
# exclude days with too many zeros, and exclude days where an outage occurred.
def preprocess_df(data_df, other_data, outage_data, max_num_missing_vals, max_num_zeros=144):
    # Remove days with more than 18 missing values.
    data_df.loc[:, other_data['Number of Missing Values'] > max_num_missing_vals] = np.nan

    # Convert -8888 to nan
    data_df = data_df.replace(-8888, np.nan)

    # Remove days with more zeros than max_num_zeros
    data_df.loc[:, (data_df == 0).sum() >= max_num_zeros] = np.nan

    # Remove days where an outage occurred
    data_df.loc[:, outage_data == 1] = np.nan

    return data_df


# Aggregate into averages for each file, ignoring columns with too many zeros
def agg_avg_load_profiles(data_filepaths, max_num_missing_vals, battery_data, max_num_zeros=144):
    print("Aggregating all average load profiles...")
    ten_min_pwr_s_list = [] # Temporary list to hold Series' with 10-min averages
    count = 0   # used to keep track of progress
    for filepath in data_filepaths[3:]:
        # print progress
        count += 1
        print(f"Aggregating All Data: Progress = {int(count/len(data_filepaths)*100)}% | {filepath}")
        try:
            location, metadata = extract_metadata(filepath)
            data_df, other_data = extract_power_data(filepath)
            year = metadata['yr_start']

            data_df = preprocess_df(data_df, other_data, battery_data[location][year]['Outage Flag'],
                                    max_num_missing_vals, max_num_zeros)

            ten_min_avg_pwr = data_df.mean(axis='columns')
            ten_min_avg_pwr.name = f"{location}_{year}"
            ten_min_pwr_s_list.append(ten_min_avg_pwr)
        except Exception as error:
            print(f"Error: {error}")

    full_avg_data_df = pd.concat(ten_min_pwr_s_list, axis='columns')
    full_avg_data_df.index.name = 'Ten-Minute Index'

    new_col_name_tuples = [(x.split('_')[0], x.split('_')[1]) for x in full_avg_data_df.columns]
    new_multiindex_cols = pd.MultiIndex.from_tuples(new_col_name_tuples, names=('Location', 'Year'))
    full_avg_data_df.columns = new_multiindex_cols
    print(full_avg_data_df)

    return full_avg_data_df


# Aggregate into averages for each file, split into each season
def agg_avg_seasonal_load_profiles(data_filepaths, max_num_missing_vals, battery_data, max_num_zeros=144):
    ten_min_pwr_s_list = [] # Temporary list to hold Series' with 10-min averages
    count = 0   # used to keep track of progress
    for filepath in data_filepaths:
        # print progress
        count += 1
        print(f"Aggregating Seasonal Data: Progress = {int(count/len(data_filepaths)*100)}% | {filepath}")
        try:
            location, metadata = extract_metadata(filepath)
            data_df, other_data = extract_power_data(filepath)
            year = metadata['yr_start']

            data_df = preprocess_df(data_df, other_data, battery_data[location][year]['Outage Flag'],
                                    max_num_missing_vals, max_num_zeros)

            # Turn column headers into a series of datetime values, then convert that
            # of datetime values into a series of values 0-3 that represent which
            # quarter of the year the day is in
            doy_vals = pd.to_datetime([f"{year}{x.split(' ')[0]}" for x in data_df.columns], format='%Y%j')
            quarter_vals = doy_vals.quarter

            for quarter_num, quarter_name in quarter_season_dict.items():
                quarter_info = quarter_vals == quarter_num  # bool index of loc of cols to keep
            
                # Series where col name gets bool value
                quarter_columns = pd.Series(quarter_info, name='Quarter Index', index=data_df.columns)
                quarter_cols_info = quarter_columns.loc[quarter_columns == True].index

                quarter_df = data_df[data_df.columns.intersection(quarter_cols_info)]

                ten_min_avg_pwr = quarter_df.mean(axis='columns')
                ten_min_avg_pwr.name = f"{location}_{year}_{quarter_name}"
                ten_min_pwr_s_list.append(ten_min_avg_pwr)
        except Exception as error:
            print(f"Error: {error}")

    full_avg_data_df = pd.concat(ten_min_pwr_s_list, axis='columns')
    full_avg_data_df.index.name = 'Ten-Minute Index'

    location_year_qtr_tuples = [(x.split('_')[0], x.split('_')[1], x.split('_')[2]) for x in full_avg_data_df.columns]
    new_multiindex_cols = pd.MultiIndex.from_tuples(location_year_qtr_tuples, names=('Location', 'Year', 'Season'))
    full_avg_data_df.columns = new_multiindex_cols
    print(full_avg_data_df)

    return full_avg_data_df


# Create a boxplot of the load profile for each ten-minute period of the day,
# then save the figure to the specified save_folder.
def calc_load_profile_weekday_weekend(data_filepaths, max_num_missing_vals, battery_data, max_num_zeros=144):
    weekday_df_list = []
    weekend_df_list = []
    count = 0   # used to keep track of progress
    for filepath in data_filepaths:
        # print progress
        count += 1
        print(f"Aggregating Weekday/Weekend Data: | Progress = {int(count/len(data_filepaths)*100)}% | {filepath}")
        try:
            location, metadata = extract_metadata(filepath)
            data_df, other_data = extract_power_data(filepath)
            year = metadata['yr_start']

            data_df = preprocess_df(data_df, other_data, battery_data[location][year]['Outage Flag'],
                                    max_num_missing_vals, max_num_zeros)

            # get locations where day was a weekday
            weekday_vals = pd.to_datetime([f"{year}{x.split(' ')[0]}" for x in data_df.columns], format='%Y%j').weekday < 5

            # Series where index is df column name, values are True if that index is a weekday
            weekday_info = pd.Series(weekday_vals, name='Weekday Index', index=data_df.columns)

            weekday_columns = weekday_info.loc[weekday_info == True].index
            weekend_columns = weekday_info.loc[weekday_info == False].index

            # extract weekday and weekend data into separate dataframes
            weekday_df = data_df[data_df.columns.intersection(weekday_columns)]
            weekend_df = data_df[data_df.columns.intersection(weekend_columns)]

            # convert to a series of average values
            weekday_df = weekday_df.mean(axis=1)
            weekend_df = weekend_df.mean(axis=1)

            # rename series of mean values
            weekday_df.name = f"{location}_{year}_weekday"
            weekend_df.name = f"{location}_{year}_weekend"

            weekday_df_list.append(weekday_df)
            weekend_df_list.append(weekend_df)
        except Exception as error:
            print(f"Error: {error}")

    full_weekday_df = pd.concat(weekday_df_list, axis=1)
    full_weekend_df = pd.concat(weekend_df_list, axis=1)

    for df in [full_weekday_df, full_weekend_df]:
        location_year_type_tuples = [(x.split('_')[0], x.split('_')[1]) for x in df.columns]
        new_multiindex_cols = pd.MultiIndex.from_tuples(location_year_type_tuples, names=('Location', 'Year'))
        df.columns = new_multiindex_cols
        df.index.name = 'Ten-Minute Index'
        print(df)

    return full_weekday_df, full_weekend_df


# Converts a 10-minutely-indexed DataFrame into an hourly-indexed DataFrame
def calc_hourly_df(avg_ten_min_df):
    hourly_avg_df = avg_ten_min_df.copy()
    hourly_avg_df.index = hourly_avg_df.index // 6
    hourly_avg_df = hourly_avg_df.groupby('Ten-Minute Index', axis=0).mean()
    hourly_avg_df.index.name = 'Hourly Index'
    
    return hourly_avg_df