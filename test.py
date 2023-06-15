# Set current working directory to the folder containing this file
import os
os.chdir(os.path.dirname(os.path.realpath(__file__)))

import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from sklearn.cluster import KMeans
import time

from Navajo_Load_Profiles_Functions import *


folders = create_folders()

data_filepaths = [f"{folders['Energy_Data_New']}\\{filename}" for filename \
    in os.listdir(folders['Energy_Data_New']) \
    if os.path.isfile(f"{folders['Energy_Data_New']}\\{filename}")]

battery_analysis_filepaths = [f"{folders['Battery_Analysis']}\\{filename}" for filename \
    in os.listdir(folders['Battery_Analysis']) \
    if os.path.isfile(f"{folders['Battery_Analysis']}\\{filename}") and '.xlsx' in filename]

start = time.time()
gen = (pd.read_excel(filepath, sheet_name='Energy Balance', index_col=[0], header=None) for filepath in battery_analysis_filepaths)
for x in gen:
    energy_balance_df = pd.read_excel(x, sheet_name='Energy Balance', index_col=[0], header=None)

end = time.time()

print(end - start)
print()

start = time.time()
df = pd.concat(pd.read_excel(filepath, sheet_name='Energy Balance', index_col=[0], header=None) for filepath in battery_analysis_filepaths)
end = time.time()

print(end - start)
print()