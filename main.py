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


# COUNT COUNTRY
df_Country_count = pd.DataFrame(df.groupby(['Live Location:Country'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Country_count = df_Country_count.fillna('Unknow')

df_Country_count['Percent'] = (df_Country_count['Total'] / df_Country_count['Total'].sum()) * 100
df_Country_count['Percent'] = df_Country_count['Percent'].round(decimals=1)


# COUNT CATEGORIES (CUSTOM FIELD Business/Professional sector)
df['Categories'] = df['Business/Professional sector'].str.split(': ').str[0]
df_Categories_count = pd.DataFrame(df.groupby(['Categories'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Categories_count = df_Categories_count.fillna('Unknow')

df_Categories_count['Percent'] = (df_Categories_count['Total'] / df_Categories_count['Total'].sum()) * 100
df_Categories_count['Percent'] = df_Categories_count['Percent'].round(decimals=1)


# COUNT SPECIALTIES (CUSTOM FIELD Business/Professional sector)
df['Specialties'] = df['Business/Professional sector'].str.split(': ').str[1]
df_Specialties_count = pd.DataFrame(df.groupby(['Specialties'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Specialties_count = df_Specialties_count.fillna('Unknow')

df_Specialties_count['Percent'] = (df_Specialties_count['Total'] / df_Specialties_count['Total'].sum()) * 100
df_Specialties_count['Percent'] = df_Specialties_count['Percent'].round(decimals=1)


# COUNT SPECIALTIES PER COUNTRY
df_SpecialtiesPerCountry_count = pd.DataFrame(df.groupby(['Live Location:Country', 'Specialties'], dropna=False)\
    .size(), columns=['Total']).sort_values(['Live Location:Country', 'Total'], ascending=[True, False]).reset_index()
df_SpecialtiesPerCountry_count = df_SpecialtiesPerCountry_count.fillna('Unknow')

df_SpecialtiesPerCountry_count['Percent'] = (df_SpecialtiesPerCountry_count['Total'] / df_SpecialtiesPerCountry_count['Total'].sum()) * 100
df_SpecialtiesPerCountry_count['Percent'] = df_SpecialtiesPerCountry_count['Percent'].round(decimals=2)


# COUNT EXPERTISE & INTERESTS (CUSTOM FIELD Industries:Industries)
df_tempIndustries = pd.DataFrame(pd.melt(df['Industries:Industries'].str.split(',', expand=True))['value'])
df_Industries_count = pd.DataFrame(df_tempIndustries.groupby(['value'], dropna=False).size(), columns=['Total'])\
    .reset_index()
df_Industries_count = df_Industries_count.fillna('AZERTY')

df_Industries_count['Percent'] = (df_Industries_count['Total'] / df.shape[0]) * 100
df_Industries_count['Percent'] = df_Industries_count['Percent'].round(decimals=2)

# EMPTY VALUES
industriesEmpty = df['Industries:Industries'].isna().sum()
industriesEmptyPercent = round((industriesEmpty / df.shape[0]) * 100, 2)

# REPLACE EMPTY VALUES AND SORT
df_Industries_count.loc[(df_Industries_count['value'] == 'AZERTY')] = [['Unknow', industriesEmpty, industriesEmptyPercent]]
df_Industries_count = df_Industries_count.sort_values(['Total'], ascending=False)


# COUNT HOW DID YOU HEAR ABOUT US (CUSTOM FIELD How did you hear about us?)
df_HowDidYouHearAboutUs_count = pd.DataFrame(df.groupby(['How did you hear about us?'], dropna=True).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_HowDidYouHearAboutUs_count = df_HowDidYouHearAboutUs_count.fillna('Unknow')

df_HowDidYouHearAboutUs_count['Percent'] = (df_HowDidYouHearAboutUs_count['Total'] / df_HowDidYouHearAboutUs_count['Total'].sum()) * 100
df_HowDidYouHearAboutUs_count['Percent'] = df_HowDidYouHearAboutUs_count['Percent'].round(decimals=1)

# REPLACE SOME VALUES
# df_HowDidYouHearAboutUs_count['How did you hear about us?'] = df_HowDidYouHearAboutUs_count['How did you hear about us?'].replace(['Email from partners (AMWC, VCS, FACE etc)'],'Email from partners')
df_HowDidYouHearAboutUs_count['How did you hear about us?'] = df_HowDidYouHearAboutUs_count['How did you hear about us?'].replace(['Other: please specify'],'Other')


# COUNT SOURCES
# JOIN LEFT WITH SUBSCRIBERS
df_WebinarSubscriber = pd.merge(df, df_subscriber, left_on='Email', right_on='email', how='left')\
    [['Email', 'source']]

df_Sources = pd.DataFrame(df_WebinarSubscriber.groupby(['source'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Sources = df_Sources.fillna('')

df_Sources['Percent'] = (df_Sources['Total'] / df_Sources['Total'].sum()) * 100
df_Sources['Percent'] = df_Sources['Percent'].round(decimals=1)

# NEW EMAIL SUBSCRIBERS
df_WebinarNew = pd.DataFrame(df[~df['Email'].isin(df_subscriber ['email'])])
newWebinar = df_WebinarNew.shape[0]

# EXCEL FILE
writer = pd.ExcelWriter(outputExcelFile, engine='xlsxwriter')

df_Country_count.to_excel(writer, index=False, sheet_name='Countries', header=['Country', 'Total', '%'])
df_Categories_count.to_excel(writer, index=False, sheet_name='Categories', header=['Category', 'Total', '%'])
df_Specialties_count.to_excel(writer, index=False, sheet_name='Specialties', header=['Specialty', 'Total', '%'])
df_SpecialtiesPerCountry_count.to_excel(writer, index=False, sheet_name='Specialties per country', header=['Country', 'Specialty', 'Total', '%'])
df_Industries_count.to_excel(writer, index=False, sheet_name='Expertise & Interests', header=['Expertise or Interest', 'Total', '%'])
df_HowDidYouHearAboutUs_count.to_excel(writer, index=False, sheet_name='How Did You Hear', header=['How did you hear about us (known)', 'Total', '%'])
df_Sources.to_excel(writer, index=False, sheet_name='Sources', header=['Source', 'Total', '%'])

writer.save()

# EXCEL FILTERS
workbook = openpyxl.load_workbook(outputExcelFile)
sheetsLits = workbook.sheetnames

for sheet in sheetsLits:
    worksheet = workbook[sheet]
    FullRange = 'A1:' + get_column_letter(worksheet.max_column) + str(worksheet.max_row)
    worksheet.auto_filter.ref = FullRange
    workbook.save(outputExcelFile)

# EXCEL COLORS
for sheet in sheetsLits:
    worksheet = workbook[sheet]
    for cell in workbook[sheet][1]:
        worksheet[cell.coordinate].fill = PatternFill(fgColor = 'FFC6C1C1', fill_type = 'solid')
        workbook.save(outputExcelFile)

# EXCEL COLUMN SIZE
for sheet in sheetsLits:
    for cell in workbook[sheet][1]:
        if get_column_letter(cell.column) == 'A':
            workbook[sheet].column_dimensions[get_column_letter(cell.column)].width = 30
        else:
            workbook[sheet].column_dimensions[get_column_letter(cell.column)].width = 10
        workbook.save(outputExcelFile)


# CHART CATEGORIES
chartLabel = df_Categories_count['Categories'].tolist()
chartLegendLabel = df_Categories_count['Categories'].tolist()
chartValue = df_Categories_count['Total'].tolist()
chartLegendPercent = df_Categories_count['Percent'].tolist()

chartLabel[-1] = ''

legendLabels = []
for i, j in zip(chartLegendLabel, map(str, chartLegendPercent)):
    legendLabels.append(i + ' (' + j + ' %)')

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig2 = plt.figure()
plt.pie(chartValue, labels=chartLabel, colors=colors, autopct=None, shadow=False, startangle=90)
plt.axis('equal')
plt.title('Categories', pad=20, fontsize=15)

plt.legend(legendLabels, loc='best', fontsize=8)

fig2.savefig(workDirectory+'myplot2.png', dpi=100)
plt.clf()

im = Image.open(workDirectory+'myplot2.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot2.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot2.png')
img.anchor = 'E4'

workbook['Categories'].add_image(img)
workbook.save(outputExcelFile)


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

fig3.savefig(workDirectory+'myplot3.png', dpi=75)
plt.clf()

im = Image.open(workDirectory+'myplot3.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot3.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot3.png')
img.anchor = 'E4'

workbook['How Did You Hear'].add_image(img)
workbook.save(outputExcelFile)


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

# Footer
plt.figtext(0.2, 0.17, WebinarFileName.replace('_', ' '), ha="left", fontsize=13, weight='bold')
plt.figtext(0.2, 0.14, 'Participants: '+str(participants)+' - New emails: '+str(newWebinar), ha="left", fontsize=11)

cb.remove()

map1.savefig(workDirectory+'mymap1.png', dpi=110, bbox_inches='tight')
plt.clf()

im = Image.open(workDirectory+'mymap1.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'mymap1.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'mymap1.png')
img.anchor = 'E2'

workbook['Countries'].add_image(img)
workbook.save(outputExcelFile)


# REMOVE PICTURES
os.remove(workDirectory+'myplot2.png')
os.remove(workDirectory+'myplot3.png')
os.remove(workDirectory+'mymap1.png')


# TERMINAL OUTPUTS AND TESTS
print("OK, export done!")