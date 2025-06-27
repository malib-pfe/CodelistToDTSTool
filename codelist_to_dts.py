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

version_num = "1.0"
file_url = "https://raw.githubusercontent.com/malib-pfe/MDRComparisonTool/refs/heads/main/version.txt"
rccfile = None

# Filter out warnings due to weird Workbook naming.
filterwarnings("ignore", message="Workbook contains no default style", category=UserWarning)
def read_file_from_github(raw_url):
    """
    Reads a file from GitHub using its raw URL.

    Args:
        raw_url (str): The raw URL of the file on GitHub.

    Returns:
        str: The content of the file, or None if an error occurred.
    """
    try:
        response = requests.get(raw_url)
        response.raise_for_status()  # Raise HTTPError for bad responses (4xx or 5xx)
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Error reading file from GitHub: {e}")
        return None

def transform_cl(rcc:str) -> pd.DataFrame:
    # Full print option used for development.
    file = rcc
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
    filename = '/Import_DTS_CL_Template_' + dts_name[0]
    now = datetime.now()
    now_string = now.strftime("_%b_%d_%Y_%H_%M_%S")
    timestamp_string = folder_path + filename + now_string + '.csv'

    df.to_csv(timestamp_string, index=False)
    return timestamp_string

async def choose_rcc_file():
    global rccfile
    file = await app.native.main_window.create_file_dialog(allow_multiple=False, file_types= ('Excel Files (*.xlsx)',))
    isFile = await run.cpu_bound(checkFile, file)
    if isFile:
        n3 = ui.notification("Checking RCC Metadata Export...", type='ongoing', timeout=None, spinner=True)
        isFileInSheet1 = await run.cpu_bound(check_file_for_sheet, 'DTS Header', file[0])
        isFileInSheet2 = await run.cpu_bound(check_file_for_sheet, 'Code Lists', file[0])
        if isFileInSheet1 and isFileInSheet2:
            n3.message = "Metadata export selected."
            n3.type = "positive"
            n3.timeout = 3
            n3.spinner = False
            rcc_filepath.set_text(file[0])
            rccfile = file[0]
        else:
            n3.message = "'DTS Header' or 'Code Lists' sheet not found. Please check file."
            n3.type = "negative"
            n3.timeout = 3
            n3.spinner = False         
    else:
        ui.notify('No file selected.')


def check_file_for_sheet(sheetname, filename):
    xl = pd.ExcelFile(filename, engine='openpyxl')
    return sheetname in xl.sheet_names

def checkFile(file):
    return file is not None

def check_file_for_col(colnames, filename, sheetname):
    df = pd.read_excel(filename, sheet_name=sheetname,engine="openpyxl")
    for colname in colnames:
        try:
            df_col = df[colname]
        except:
            return colname
    return True

async def handle_execute():
    n = ui.notification("Executing... Please Wait.", type='ongoing', timeout=None, spinner=True)
    executeBtn.disable()

    rcc = rcc_filepath.text

    global result
    result = await run.cpu_bound(transform_cl, rcc)
    n.message = "Table export located in " + result
    n.type = "positive"
    n.timeout = 3
    n.spinner = False
    clearBtn.enable()

async def reset_page():
    executeBtn.enable()
    clearBtn.disable()
    rcc_filepath.set_text('')

# Define the UI.
ui.add_css(
    """
    .my-sticky-header-table {
        /* height or max-height is important */
        max-height: 400px;
        /* this is when the loading indicator appears */
        /* prevent scrolling behind sticky top row on focus */
    }
    
    .my-sticky-header-table .q-table__top,
    .my-sticky-header-table .q-table__bottom,
    .my-sticky-header-table thead tr:first-child th {
        /* bg color is important for th; just specify one */
        background-color: #00b4ff;
    }
    
    .my-sticky-header-table thead tr th {
        position: sticky;
        z-index: 1;
    }
    
    .my-sticky-header-table thead tr:first-child th {
        top: 0;
    }
    
    .my-sticky-header-table.q-table--loading thead tr:last-child th {
        /* height of all previous header rows */
        top: 48px;
    }
    
    .my-sticky-header-table tbody {
        /* height of all previous header rows */
        scroll-margin-top: 48px;
    }  
    """
)

state = {}

with ui.header():
    ui.label('Codelist to DTS Export Tool').style('font-size: 200%; font-weight: bold').classes('absolute-center')
    
with ui.row():
    ui.label('RCC Codelist Import').style('font-weight:bold')
    rcc_filepath = ui.label()

ui.button('Select RCC Codelist Import',on_click=choose_rcc_file)

ui.space()

with ui.row():
    executeBtn = ui.button("Execute", on_click= lambda: handle_execute() if rcc_filepath.text != '' else ui.notify('Please select a file to proceed.'))
    clearBtn = ui.button("Clear File", on_click= reset_page)
    clearBtn.disable()

file_content = read_file_from_github(file_url)
if file_content != version_num:
    executeBtn.disable()
    ui.notification("This app is out of date. Please use newest version.", timeout=False, type = "negative")

try:
    ui.run(native=True, reload=False, title="Codelist to DTS Tool")
except asyncio.CancelledError as e:
    pass
except KeyboardInterrupt:
    pass
