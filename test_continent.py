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
shp_simple_areas = r'C:/Users/Georges/PycharmProjects/data/simple_areas/simple_areas'
inputCountryConversion = r'C:/Users/Georges/PycharmProjects/data/countries_conversion.xlsx'

workDirectory = r'C:/Users/Georges/Downloads/Webinar/'

WebinarFileName = '20210127_Webinar_TheArtAndScience_Karkatzoulis'

ReportExcelFile = workDirectory + WebinarFileName + '_Report.xlsx'
NewAddThenDeleteExcelFile = workDirectory + WebinarFileName + '_NewAddJooThenDelete.xlsx'
NewCollectExcelFile = workDirectory + WebinarFileName + '_NewToCollect.xlsx'


# WEBINAR EXCEL IMPORT
inputExcelFile = workDirectory+WebinarFileName+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='export', engine='openpyxl',
                   usecols=['ID', 'Honorary title', 'First name', 'Last name', 'Email', 'Live Location:Country',
                            'Industries:Industries', 'Business/Professional sector', 'How did you hear about us?'
                            ])

participants = df.shape[0]


# JOO_ACYMAILING_SUBSCRIBER IMPORT
df_subscriber = pd.read_csv(workDirectory+'joo_acymailing_subscriber.csv', sep=',', usecols=['source', 'email'])
# SOURCES HARMONIZATION
df_subscriber['source'] = df_subscriber['source'].replace({'EXTERN: ': ''}, regex=True)
df_subscriber['source'] = df_subscriber['source'].replace({'PROSPECT: ': ''}, regex=True)


# COUNTRY CONVERSION IMPORT
df_CountryConversion = pd.read_excel(inputCountryConversion, sheet_name='countries', engine='openpyxl',
                   usecols=['COUNTRY_HB', 'continent_stat'])


# NEW EMAILS
df_WebinarNew = pd.DataFrame(df[~df['Email'].isin(df_subscriber['email'])])
newWebinar = df_WebinarNew.shape[0]


# COUNT NEW AREAS
# JOIN LEFT WITH COUNTRY CONVERSION
df_WebinarNewAreas = pd.merge(df_WebinarNew, df_CountryConversion, left_on='Live Location:Country', right_on='COUNTRY_HB', how='left')\
    [['Email', 'continent_stat']]

df_NewAreasCount = pd.DataFrame(df_WebinarNewAreas.groupby(['continent_stat'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_NewAreasCount = df_NewAreasCount.fillna('Unknow')

df_NewAreasCount['Percent'] = (df_NewAreasCount['Total'] / df_NewAreasCount['Total'].sum()) * 100
df_NewAreasCount['Percent'] = df_NewAreasCount['Percent'].round(decimals=1)





# TERMINAL OUTPUTS AND TESTS
print(tab(df_NewAreasCount.head(35), headers='keys', tablefmt='psql'))