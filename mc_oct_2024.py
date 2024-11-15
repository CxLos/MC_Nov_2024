# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
import datetime
import folium
import os
import sys
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component
# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
script_dir = os.path.dirname(os.path.abspath(__file__))
data_path1 = 'data/BMHC_Navigation_Outreach_Impact_Report.xlsx'
data_path2 = 'data/BMHC_MarCom_Impact_Report.xlsx'
file_path1 = os.path.join(script_dir, data_path1)
file_path2 = os.path.join(script_dir, data_path2)

# Sheet Names
responses = 'Responses'
marcom = 'MarCom'

# Read data from excel file
data_r = pd.read_excel(file_path1, sheet_name=responses)
data_m1 = pd.read_excel(file_path2, sheet_name=responses)
data_m = pd.read_excel(file_path2, sheet_name=marcom)

# df_r = data_r2.copy()
df_m = data_m.copy()

# Concatenate dataframes
# df = pd.concat([data1, data2], ignore_index=True)

# Trim leading and trailing whitespaces from column names
# df_r.columns = df_r.columns.str.strip()
df_m.columns = df_m.columns.str.strip()

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# print(df_m.head())
# print('Total entries: ', len(df))
# print('Column Names: \n', df_m.columns)
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())

# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

# Column Names: 
#  Index(['MarCom Reporting Month',
#        'BMHC Organizational Communications/Marketing Activities',
#        'BMHC Organizational Public Information',
#        'BMHC Organizational Products', 
#        'BMHC Organizational Event-Oriented',
#        'Care Network Enhancement Communications/Marketing',
#        'Care Network Enhancement Public Information',
#        'Care Network Enhancement Products',
#        'Care Network Enhancement Event-Oriented',
#        'Know Your Numbers Communications/Marketing',
#        'Know Your Numbers Public Information', 
#        'Know Your Numbers Products',
#        'Know Your Numbers Event Oriented',
#        'Health Awareness & ED Communications/Marketing',
#        'Health Awareness & ED Public Information',
#        'Health Awareness & ED Products',
#        'Health Awareness & ED Event-Oriented'],
#       dtype='object')

# ------------------------------- Missing Values ----------------------------------- #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])


# ============================== Data Preprocessing ========================== #

# Check for duplicate columns
# duplicate_columns = df.columns[df.columns.duplicated()].tolist()
# print(f"Duplicate columns found: {duplicate_columns}")
# if duplicate_columns:
#     print(f"Duplicate columns found: {duplicate_columns}")

# Fill Missing Values
# df['screener_id'] = df['screener_id'].fillna("N/A")
# df['referral_id'] = df['referral_id'].fillna("N/A")
# # df['created_at'] = df['created_at'].fillna("N/A")
# df['site'] = df['site'].fillna('google forms')
# # df['Timestamp'] = df['Timestamp'].interpolate()
# df['Last Name'] = df['Last Name'].fillna('N/A')
# df['Gender'] = df['Gender'].fillna('N/A')
# df['Age'] = df['Age'].fillna(0)
# df['Physical Appointment'] = df['Physical Appointment'].fillna('N/A')
# df['Coverage'] = df['Coverage'].fillna('N/A')
# df['Status'] = df['Status'].fillna('N/A')
# df['Service'] = df['Service'].fillna('N/A')
# df['Housing'] = df['Housing'].fillna('N/A')
# df['Income'] = df['Income'].fillna('N/A')
# df['BMHC Referrals'] = df['BMHC Referrals'].fillna('N/A')
# df['Mental Health'] = df['Mental Health'].fillna('N/A')
# df['Transportation'] = df['Transportation'].fillna('N/A')
# df['Diversion'] = df['Diversion'].fillna('N/A')
# df['Communication Type'] = df['Communication Type'].fillna('N/A')
# df['Race/Ethnicity'] = df['Race/Ethnicity'].fillna('N/A')
# df['Social Services'] = df['Social Services'].fillna('N/A')
# df['Rating'] =df['Rating'].fillna(0)
# df['Rating'] = df['Rating'].astype('Int64')
# df['Completed Survey'] =df['Completed Survey'].fillna('N/A')
# df['Veteran'] = df['Veteran'].fillna('N/A')
# df['Services Not Completed'] =df['Services Not Completed'].fillna('N/A')
# df['Reason for No Show'] =df['Reason for No Show'].fillna('N/A')
# df['Zip Code'].fillna(df['Zip Code'].mode()[0], inplace=True)
# df['Zip Code'] = df['Zip Code'].astype('Int64')
# df['Zip Code'] = df['Zip Code'].replace(-1, df['Zip Code'].mode()[0])

# print(df.dtypes)

# income_mode = df['Income'].mode()
# print('Income Mode:', income_mode)

# missing = df.isnull().sum()
# print('Columns with missing values after fillna: \n', missing[missing>0])

# value counts for 'Rating' column
# rating_counts = df['Rating'].value_counts()
# print('Rating Counts:\n', rating_counts)

# -----------------------------------------------------------------------------

# Get the distinct values in column

# distinct_service = df['What service did/did not complete?'].unique()
# print('Distinct:\n', distinct_service)

# ------------------------------------ SQL ---------------------------------------

# Connect to SQL
con = sqlite3.connect("bmhc_2024.db")
cur = con.cursor()

df_m.to_sql("bmhc_responses_q4_2024", con, if_exists='replace', index=False, method="multi")

# # Show list of all tables in db.
# # tables = pd.read_sql_query("""
# #   SELECT name 
# #   FROM sqlite_master 
# #   WHERE type = 'table';
# # """, con)
# # print("Tables in the database:\n", tables)

# # # Check if data is inserted correctly
# # df_check = pd.read_sql_query("SELECT * FROM bmhc_responses_q3_2024 LIMIT 5;", con)
# # print(df_check)

con.close()

# ========================= Filtered DataFrames ========================== #

# Organizational Events
organizational_columns = [
    'BMHC Organizational Communications/Marketing Activities',
    'BMHC Organizational Public Information',
    'BMHC Organizational Products',
    'BMHC Organizational Event-Oriented'
]

# Create a new DataFrame with only the specified columns
df_org = df_m[organizational_columns]

# Print the filtered DataFrame to verify the results
# print(df_org)

# Network Enhancement
network_columns = [
    'Care Network Enhancement Communications/Marketing',
    'Care Network Enhancement Public Information',
    'Care Network Enhancement Products',
    'Care Network Enhancement Event-Oriented'
]

# Create a new DataFrame with only the specified columns
df_network = df_m[network_columns]

# Print the filtered DataFrame to verify the results
# print(df_network)

# Know Your Numbers
numbers_columns = [
    'Know Your Numbers Communications/Marketing',
    'Know Your Numbers Public Information',
    'Know Your Numbers Products',
    'Know Your Numbers Event Oriented'
]

# Create a new DataFrame with only the specified columns
df_numbers = df_m[numbers_columns]

# Print the filtered DataFrame to verify the results
# print(df_numbers)

# Health Awareness & ED
health_columns = [
    'Health Awareness & ED Communications/Marketing',
    'Health Awareness & ED Public Information',
    'Health Awareness & ED Products',
    'Health Awareness & ED Event-Oriented'
]

# Create a new DataFrame with only the specified columns
df_health = df_m[health_columns]

# Print the filtered DataFrame to verify the results
# print(df_health)

# # ========================== DataFrame Table ========================== #

# Organizational Events Table
org_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_org.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_org[col] for col in df_org.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

org_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# Network Enhancement Table
network_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_network.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_network[col] for col in df_network.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

network_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# Know Your Numbers Table
numbers_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_numbers.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_numbers[col] for col in df_numbers.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

numbers_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# Health Awareness & ED Table
health_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_health.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_health[col] for col in df_health.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

health_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server

app.layout = html.Div(children=[ 

    html.Div(className='divv', children=[ 
        
        html.H1('BMHC MarCom November 2024 Report', 
        className='title'),

        html.A(
        'Repo',
        href='https://github.com/CxLos/MC_Impact_11_2024',
        className='btn')
    ]),    

# Data Table Organizational
html.Div(
    className='row0',
    children=[

        html.Div(
            className='table',
            children=[
                html.H1(
                    className='table-title1',
                    children='Organizational Events'
                )
            ]
        ),
        html.Div(
            className='highlights',
            children=[
                html.Div(
                    className='highs-activity',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Activities:']
                        ),
                        html.H1(
                            className='high2',
                            children=['10']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-public',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Public Info:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-product',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Products:']
                        ),
                        html.H1(
                            className='high2',
                            children=['2']
                        ),
                    ],
                ),
                html.Div(
                    className='highs',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Events:']
                        ),
                        html.H1(
                            className='high2',
                            children=['3']
                        ),
                    ],
                ),

            ]
        ),
        html.Div(
            className='table2', 
            children=[
                dcc.Graph(
                    className='data',
                    figure=org_table
                )
            ]
        )
    ]
),

# Data Table Network Enhancement
html.Div(
    className='row0',
    children=[
        html.Div(
            className='table',
            children=[
                html.H1(
                    className='table-title',
                    children='Care Network Enhancement'
                )
            ]
        ),
                html.Div(
            className='highlights',
            children=[
                html.Div(
                    className='highs-activity',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Activities:']
                        ),
                        html.H1(
                            className='high2',
                            children=['6']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-public',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Public Info:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-product',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Products:']
                        ),
                        html.H1(
                            className='high2',
                            children=['5']
                        ),
                    ],
                ),
                html.Div(
                    className='highs',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Events:']
                        ),
                        html.H1(
                            className='high2',
                            children=['0']
                        ),
                    ],
                ),

            ]
        ),
        html.Div(
            className='table2', 
            children=[
                dcc.Graph(
                    className='data',
                    figure=network_table
                )
            ]
        )
    ]
),

# Data Table Know Your Numbers
html.Div(
    className='row0',
    children=[
        html.Div(
            className='table',
            children=[
                html.H1(
                    className='table-title',
                    children='Know Your Numbers'
                )
            ]
        ),
                html.Div(
            className='highlights',
            children=[
                html.Div(
                    className='highs-activity',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Activities:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-public',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Public Info:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-product',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Products:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Events:']
                        ),
                        html.H1(
                            className='high2',
                            children=['0']
                        ),
                    ],
                ),

            ]
        ),
        html.Div(
            className='table2', 
            children=[
                dcc.Graph(
                    className='data',
                    figure=numbers_table
                )
            ]
        )
    ]
),

# Data Table Health Awareness & ED
html.Div(
    className='row00',
    children=[
        html.Div(
            className='table',
            children=[
                html.H1(
                    className='table-title',
                    children='Health Awareness & ED'
                )
            ]
        ),
                html.Div(
            className='highlights',
            children=[
                html.Div(
                    className='highs-activity',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Activities:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-public',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Public Info:']
                        ),
                        html.H1(
                            className='high2',
                            children=['10']
                        ),
                    ],
                ),
                html.Div(
                    className='highs-product',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Products:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),
                html.Div(
                    className='highs',
                    children=[
                        html.H1(
                            className='high1',
                            children=['Events:']
                        ),
                        html.H1(
                            className='high2',
                            children=['1']
                        ),
                    ],
                ),

            ]
        ),
        html.Div(
            className='table2', 
            children=[
                dcc.Graph(
                    className='data',
                    figure=health_table
                )
            ]
        )
    ]
),
])

# Callback function
# @app.callback(
#     Output('', 'figure'),
#     [Input('', 'value')]
# )

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# ----------------------------------------------- Updated Database ----------------------------------------

# updated_path = 'data/bmhc_q4_2024_cleaned.xlsx'
# data_path = os.path.join(script_dir, updated_path)
# df.to_excel(data_path, index=False)
# print(f"DataFrame saved to {data_path}")

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx