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

workDirectory = r'C:/Users/Georges/Downloads/Webinar/'

WebinarFileName = '20210330_Webinar_Russian'

outputExcelFile = workDirectory+WebinarFileName+'_Report.xlsx'


# WEBINAR EXCEL IMPORT
inputExcelFile = workDirectory+WebinarFileName+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='export', engine='openpyxl',
                   usecols=['ID', 'Honorary title', 'First name', 'Last name', 'Email', 'Live Location:Country',
                            'Industries:Industries', 'Business/Professional sector', 'How did you hear about AMS?'
                            ])


# JOO_ACYMAILING_SUBSCRIBER IMPORT
df_subscriber = pd.read_csv(workDirectory+'joo_acymailing_subscriber.csv', sep=',', usecols=['source', 'email'])
# SOURCES HARMONIZATION
df_subscriber['source'] = df_subscriber['source'].replace({'EXTERN: ': ''}, regex=True)
df_subscriber['source'] = df_subscriber['source'].replace({'PROSPECT: ': ''}, regex=True)


# NEW EMAILS
df_WebinarNew = pd.DataFrame(df[~df['Email'].isin(df_subscriber['email'])])
newWebinar = df_WebinarNew.shape[0]




# COUNT NEW EXPERTISE & INTERESTS (CUSTOM FIELD Industries:Industries)
df_NewtempIndustries = pd.DataFrame(pd.melt(df_WebinarNew['Industries:Industries'].str.split(',', expand=True))['value'])
df_NewIndustries_count = pd.DataFrame(df_NewtempIndustries.groupby(['value'], dropna=False).size(), columns=['Total'])\
    .reset_index()
df_NewIndustries_count = df_NewIndustries_count.fillna('AZERTY')

df_NewIndustries_count['Percent'] = (df_NewIndustries_count['Total'] / df_WebinarNew.shape[0]) * 100
df_NewIndustries_count['Percent'] = df_NewIndustries_count['Percent'].round(decimals=2)

# EMPTY VALUES
if newWebinar > 0 :
    industriesEmpty = df_WebinarNew['Industries:Industries'].isna().sum()
    industriesEmptyPercent = round((industriesEmpty / df_WebinarNew.shape[0]) * 100, 2)





print(tab(df_NewIndustries_count, headers='keys', tablefmt='psql', showindex=False))

