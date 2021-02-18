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


# COUNT HOW DID YOU HEAR ABOUT US (CUSTOM FIELD How did you hear about us?)
df_HowDidYouHearAboutUs_count = pd.DataFrame(df.groupby(['How did you hear about us?'], dropna=True).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_HowDidYouHearAboutUs_count = df_HowDidYouHearAboutUs_count.fillna('Unknow')

df_HowDidYouHearAboutUs_count['Percent'] = (df_HowDidYouHearAboutUs_count['Total'] / df_HowDidYouHearAboutUs_count['Total'].sum()) * 100
df_HowDidYouHearAboutUs_count['Percent'] = df_HowDidYouHearAboutUs_count['Percent'].round(decimals=1)

# REPLACE SOME VALUES
df_HowDidYouHearAboutUs_count['How did you hear about us?'] = df_HowDidYouHearAboutUs_count['How did you hear about us?'].replace(['Other: please specify'],'Other')



# CHART HOW DID YOU HEAR ABOUT US (CUSTOM FIELD How did you hear about us?)
chartLabel = df_HowDidYouHearAboutUs_count['How did you hear about us?'].tolist()
chartLegendLabel = df_HowDidYouHearAboutUs_count['How did you hear about us?'].tolist()
chartValue = df_HowDidYouHearAboutUs_count['Total'].tolist()
chartLegendPercent = df_HowDidYouHearAboutUs_count['Percent'].tolist()

legendLabels = []
for i, j in zip(chartLegendLabel, map(str, chartLegendPercent)):
    legendLabels.append(i + ' (' + j + ' %)')

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig3 = plt.figure(figsize=(14,10))
plt.pie(chartValue, labels=chartLabel, colors=colors, autopct='%1.1f%%', shadow=False, startangle=90)

plt.axis('equal')
plt.title('How did you hear about us (known)', pad=20, fontsize=15)

plt.legend(legendLabels, loc='best', fontsize=8)

fig3.savefig(workDirectory+'myplot3.png', dpi=100)
plt.show()
plt.clf()