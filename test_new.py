import os
from datetime import date
from tabulate import tabulate as tab

import pandas as pd

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from mpl_toolkits.basemap import Basemap

import numpy as np

from PIL import Image, ImageOps

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

today = date.today()

shp_simple_countries = r'C:/Users/Georges/PycharmProjects/data/simple_countries/simple_countries'
workDirectory = r'C:/Users/Georges/Downloads/'

WebinarFileName = '20210127_Webinar_TheArtAndScience_Karkatzoulis'

outputExcelFile = workDirectory+WebinarFileName+'_Report.xlsx'


# Excel import
inputExcelFile = workDirectory+WebinarFileName+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='export', engine='openpyxl',
                   usecols=['ID', 'Honorary title', 'First name', 'Last name', 'Email', 'Live Location:Country',
                            'Industries:Industries', 'Business/Professional sector', 'How did you hear about us?'
                            ])

participants = df.shape[0]


# JOO_ACYMAILING_SUBSCRIBER IMPORT
df_subscriber = pd.read_csv(workDirectory+'joo_acymailing_subscriber.csv', sep=';', usecols=['subid', 'source', 'email'])
# SOURCES HARMONIZATION
df_subscriber['source'] = df_subscriber['source'].replace({'EXTERN: ': ''}, regex=True)
df_subscriber['source'] = df_subscriber['source'].replace({'PROSPECT: ': ''}, regex=True)

# COUNT SOURCES

# NEW EMAIL SUBSCRIBERS
df_WebinarNew = pd.DataFrame(df[~df['Email'].isin(df_subscriber ['email'])])
newWebinar = df_WebinarNew.shape[0]


print(tab(df_WebinarNew, headers='keys', tablefmt='psql', showindex=False))
print(newWebinar)

