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

# REMOVE ADMIN
index_drop1 = df[df['Email'].apply(lambda x: x.endswith('@informa.com'))].index
df = df.drop(index_drop1)
index_drop2 = df[df['Email'].apply(lambda x: x.endswith('@euromedicom.com'))].index
df = df.drop(index_drop2)
index_drop3 = df[df['Email'].apply(lambda x: x.endswith('@eurogin.com'))].index
df = df.drop(index_drop3)
index_drop4 = df[df['Email'].apply(lambda x: x.endswith('@multispecialtysociety.com'))].index
df = df.drop(index_drop4)
index_drop5 = df[df['Email'].apply(lambda x: x.endswith('@ce.com.co'))].index
df = df.drop(index_drop5)
index_drop6 = df[df['Email'].apply(lambda x: x == ('max.carter11@yahoo.com'))].index
df = df.drop(index_drop6)

# TERMINAL OUTPUTS AND TESTS
print(df)