import pandas as pd
import numpy as np
import openpyxl as op
from datetime import datetime
import os
from warnings import filterwarnings
from nicegui import app, ui, run, html, native
import requests
import asyncio
import re

file = "sample.xlsx"
header_df = pd.read_excel(file, sheet_name="DTS Header", engine="openpyxl", header=None)
cl_df = pd.read_excel(file, sheet_name="Code Lists",engine="openpyxl", header=None).dropna(how="all")
pd.set_option('display.max_rows', None)

list_of_names = []
list_of_codes = []

for i in range(0,len(cl_df),4):
    name = cl_df.iloc[i + 1,:][1]
    item_codes = cl_df.iloc[i+3, :].dropna()[1:].values
    list_of_names.extend([name] * len(item_codes))
    list_of_codes.extend(item_codes.tolist())

df_len = len(list_of_codes)

dts_name = [header_df.iloc[0,1]] * df_len
study_name = [header_df.iloc[1,1]] * df_len
study_protocol = [header_df.iloc[2,1]] * df_len

df = pd.DataFrame(study_name, columns=['Study Name'])
df['Protocol ID'] = study_protocol
df['DTS Name'] = dts_name
df['DTS CL Name'] = list_of_names
df['Source Code'] = list_of_codes

# Create file path for output file.
folder_path = os.path.dirname(file)
filename = '/Import_DTS_CL_Template_' + dts_name + '_'
now = datetime.now()
now_string = now.strftime("_%b_%d_%Y_%H_%M_%S")
timestamp_string = folder_path + filename + now_string + '.xlsx'

df.to_csv(timestamp_string)