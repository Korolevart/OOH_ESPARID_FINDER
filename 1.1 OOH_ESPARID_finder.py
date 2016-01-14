import os
os.chdir('J:\MEC\Analytics and Insight\Korolev\Python\OOH')
import pyodbc
import pandas as pd
import sqlalchemy
import openpyxl
import xlrd
import numpy
import math
#Loading OOH functions
from OOH_functions import *


#Reading all files
#Reading file from group
df_working=pd.read_excel('C:/Users/artem.korolev/Desktop/MEGA/All d2.xlsx','Sheet1',index_col=None, na_values=['NA'])
#Looking for coordinates in Google
df_working=Google_search(df_working)

#Reading SQL database
print("Getting data from SQL")
engine = sqlalchemy.create_engine("mssql+pyodbc://espar:espar@MSKSQLP01110/Odaplan?driver=SQL Server")
df_SQL=pd.read_sql("""SELECT  distinct b.[CITY3],a.[ESPAR_ID],a.[OWNER], b.[X1], b.[Y1],b.[TYPE_GENER], b.[SIZES],b.[ABC]
                FROM [Odaplan].[dbo].[OP] a join [Odaplan].[dbo].[OSDATA] b on a.[ESPAR_ID]=b.[ESPAR_ID] where a.[PROJECT_ID]=(SELECT [PROJECT_ID]
                FROM [Odaplan].[dbo].[PROJECTS] where [NAME]='2015_1') and b.X1 is not null""",engine)
#Reading City_index and making changes
print("Reading City_index and making changes")
df_working = dictionary_check(df_working,'J:/MEC/Analytics and Insight/Korolev/Python/OOH/CITY_index.xlsx','City')
#Reading ABC_index and making changes
print("Reading ABC_index and making changes")
df_working = dictionary_check(df_working,'J:/MEC/Analytics and Insight/Korolev/Python/OOH/ABC_index.xlsx','ABC')
#Reading Owner_index and making changes
print("Reading Owner_index and making changes")
df_working = dictionary_check(df_working,'J:/MEC/Analytics and Insight/Korolev/Python/OOH/OWNER_index.xlsx','Owner')
#Reading Sizes_index and making changes
print("Reading Sizes_index and making changes")
df_working = dictionary_check(df_working,'J:/MEC/Analytics and Insight/Korolev/Python/OOH/SIZES_index.xlsx','Sizes')

#Creating Key columns in data from group
#df_working['CITY'] + df_working['OWNER']+ df_working['SIZES']+ df_working['ABC']
df_working['Key1'] = df_working['City'] + df_working['Owner']+ df_working['Sizes']+ df_working['ABC']
#df_working['CITY'] + df_working['SIZES']+ df_working['ABC']
df_working['Key2'] = df_working['City'] + df_working['Sizes']+ df_working['ABC']
#df_working['CITY'] + df_working['SIZES']
df_working['Key3'] = df_working['City'] + df_working['Sizes']


#Creating Key column in data from SQL
#df_SQL['CITY'] + df_SQL['OWNER']+ df_SQL['TYPE_GENER']+ df_SQL['SIZES']+ df_SQL['ABC']
df_SQL['Key1'] = df_SQL['CITY3'] + df_SQL['OWNER']+ df_SQL['SIZES']+ df_SQL['ABC']
#df_SQL['CITY'] + df_SQL['TYPE_GENER']+ df_SQL['SIZES']+ df_SQL['ABC']
df_SQL['Key2'] = df_SQL['CITY3'] + df_SQL['SIZES']+ df_SQL['ABC']
#df_SQL['CITY'] + df_SQL['TYPE_GENER']+ df_SQL['SIZES']
df_SQL['Key3'] = df_SQL['CITY3'] + df_SQL['SIZES']


#Running loop with KEY1
df_working=KEY_loop(df_working, df_SQL, 'Key1', 'Distan1', 'ESPAR_ID1')
df_working2=df_working[df_working['Distan1'].apply(numpy.isnan)]
df_working2.reset_index(drop=True,inplace =True)
#Running loop with KEY2
df_working2=KEY_loop(df_working2, df_SQL, 'Key2', 'Distan2', 'ESPAR_ID2')
df_working3=df_working2[df_working2['Distan2'].apply(numpy.isnan)]
df_working3.reset_index(drop=True,inplace =True)
#Running loop with KEY3
df_working3=KEY_loop(df_working3, df_SQL, 'Key3', 'Distan3', 'ESPAR_ID3')
df_working2=df_working2[['Adress1','Distan2','ESPAR_ID2']]
df_working3=df_working3[['Adress1','Distan3','ESPAR_ID3']]

print('1')
merged_left = pd.merge(left=df_working,right=df_working2, how='left', left_on='Adress1', right_on='Adress1')
print('2')
merged_left = pd.merge(left=merged_left,right=df_working3, how='left', left_on='Adress1', right_on='Adress1')
print('3')
writer = pd.ExcelWriter('C:/Users/artem.korolev/Desktop/MEGA/df_working.xlsx')
merged_left.to_excel(writer,'Sheet1')
writer.save()


