# -*- coding: utf-8 -*-
"""
Created on Thu Mar  2 13:24:43 2023

@author: agheorghitoiu
"""

import re
import os
import nltk
import pandas as pd
import numpy as np
#import xlsxwriter
pd.options.display.max_colwidth = 200


#############################################################################################################
#Set proces parameters
#############################################################################################################
#os.chdir('C:\\aJB\\Community\\2019\\Python\\LDA\\NPS_LDA\\VoC_Sites')
os.chdir('C:\\Users\\agheorghitoiu\\Desktop\\Ideas monthly snapshots')
file_name = 'ideas_statusses_Mar2023.xlsx'

#read the input excel file
data_raw_df = pd.read_excel(file_name)

#transform the status_update_date into an integer. This interger will be used to identify the 
#latest status for an idea in each snapshot month-year. 
#see snapshot_df.loc[snapshot_df.groupby('IdeaId')['status_update_date_int'].idxmax()] below
from datetime import datetime
#data_raw_df['status_update_date'] = pd.datetime(data_raw_df['status_update_date'])
#data_raw_df['status_update_date_int']=data_raw_df['status_update_date'].apply(lambda x:int(x.strftime('%Y%m%d%H%M%S')))
data_raw_df['status_update_date_int'] = pd.to_datetime(data_raw_df["status_update_date"], errors="coerce")

data_raw_df['status_update_date_int'] = data_raw_df["status_update_date_int"].dt.strftime('%Y%m%d%H%M%S')
data_raw_df[['status_update_date_int']] =  data_raw_df[['status_update_date_int']].astype(float)

#set range of months: start at april 2017, Community ideas was lanched end Feb 2017 (1st of March)
from pandas.tseries.offsets import MonthEnd
rng_end = rng_end = (data_raw_df['status_update_date'].max()+MonthEnd(1)).strftime('%m/%d/%Y') #MonthEnd(1) is required to include current month
rng_start = '4/1/2017'
rng = pd.date_range(rng_start,rng_end, freq='M')  #this generates set of months-year from 1 Apr 2017 until Currrent

#filter to select all rows where status_update_date is after 1st of March 2017
start_date_filter = data_raw_df['status_update_date'].dt.date >= datetime.strptime('01/3/17', '%d/%m/%y').date()


#Generate monthly snapshots
#logic: For each month-year select all ideas that exist. Then pull out latest status for each idea
all_monthly_snapshots_df = pd.DataFrame() #create empty dataframe
for i in rng:
    snapshot_df = data_raw_df.loc[(data_raw_df['status_update_date'] <= i) & start_date_filter]
    snapshot_df = snapshot_df.loc[snapshot_df.groupby('IdeaId')['status_update_date_int'].idxmax()]
    snapshot_df['snap_month_year']=i
    snapshot_df['snap_year']=i.strftime('%Y')
    all_monthly_snapshots_df = all_monthly_snapshots_df.append(snapshot_df, ignore_index=True)

#################################################################################################
#Write detailed output to xlsx in output destination
#################################################################################################
#
# https://xlsxwriter.readthedocs.io/example_pandas_header_format.html
# Create a Pandas Excel writer using XlsxWriter as the engine.
#
# To addess warning "with link or location/anchor > 2079 characters since it exceeds Excel's limit for URLS"
# see https://stackoverflow.com/questions/35440528/how-to-save-in-xlsx-long-url-in-cell-using-pandas
# To addess warning "with link or location/anchor > 2079 characters since it exceeds Excel's limit for URLS"
# see https://stackoverflow.com/questions/35440528/how-to-save-in-xlsx-long-url-in-cell-using-pandas

out_monthly_snapshots_df = all_monthly_snapshots_df[['snap_year','snap_month_year','IdeaId','WkReportIdeaStatus','IdeaStatus','GroupName','ChallengeName']]
#out_monthly_snapshots_df = all_monthly_snapshots_df
writer = pd.ExcelWriter('ideas_monthly_snapshots.xlsx', engine='xlsxwriter',options={'strings_to_urls': False})
#writer = pd.ExcelWriter('Cluster output '+ str(k) + " " + file_name, engine='xlsxwriter')
out_monthly_snapshots_df.to_excel(writer, sheet_name='monthly_snaps',index=False)
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book

################################################################################
#Select worksheet with Summary
################################################################################
worksheet = writer.sheets['monthly_snaps']

# Add a header format.
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})

# Write the column headers with the defined format.
for col_num, value in enumerate(out_monthly_snapshots_df.columns.values):
    worksheet.write(0, col_num, value, header_format)

# Apply the autofilter based on the dimensions of the dataframe.
worksheet.autofilter(0, 0, out_monthly_snapshots_df.shape[0], out_monthly_snapshots_df.shape[1]-1)
#Freeze the Row-header
worksheet.freeze_panes(1, 1)
# Close the Pandas Excel writer and output the Excel file.
writer.save()
