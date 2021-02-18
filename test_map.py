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

# COUNT COUNTRY
df_Country_count = pd.DataFrame(df.groupby(['Live Location:Country'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Country_count = df_Country_count.fillna('Unknow')

df_Country_count['Percent'] = (df_Country_count['Total'] / df_Country_count['Total'].sum()) * 100
df_Country_count['Percent'] = df_Country_count['Percent'].round(decimals=1)

print(tab(df_Country_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(participants)


# MAP COUNTRIES
df_Country_count.set_index('Live Location:Country', inplace=True)

my_values = df_Country_count['Percent']

num_colors = 30
cm = plt.get_cmap('Blues')
scheme = [cm(i / num_colors) for i in range(num_colors)]

my_range = np.linspace(my_values.min(), my_values.max(), num_colors)

df_Country_count['Percent'] = np.digitize(my_values, my_range) - 1

map1 = plt.figure(figsize=(14, 8))

ax = map1.add_subplot(111, frame_on=False)

m = Basemap(lon_0=0, projection='robin')
m.drawmapboundary(color='w')

m.readshapefile(shp_simple_countries, 'units', color='#444444', linewidth=.2, default_encoding='iso-8859-15')

for info, shape in zip(m.units_info, m.units):
    shp_ctry = info['COUNTRY_HB']
    if shp_ctry not in df_Country_count.index:
        color = '#dddddd'
    else:
        color = scheme[df_Country_count.loc[shp_ctry]['Percent']]

    patches = [Polygon(np.array(shape), True)]
    pc = PatchCollection(patches)
    pc.set_facecolor(color)
    ax.add_collection(pc)

# Cover up Antarctica
ax.axhspan(0, 1000 * 1800, facecolor='w', edgecolor='w', zorder=2)

# Draw color legend
ax_legend = map1.add_axes([0.2, 0.14, 0.6, 0.03], zorder=3)
cmap = mpl.colors.ListedColormap(scheme)
cb = mpl.colorbar.ColorbarBase(ax_legend, cmap=cmap, ticks=my_range, boundaries=my_range, orientation='horizontal')

plt.figtext(0.2, 0.15, WebinarFileName.replace('_', ' '), ha="left", fontsize=13, weight='bold')
plt.figtext(0.2, 0.12, 'Participants: '+str(participants), ha="left", fontsize=11)

cb.remove()

plt.show()
plt.clf()

