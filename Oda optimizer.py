import os
os.chdir('C:/Users/artem.korolev/Desktop/Planeta/ESP')
import pyodbc
import pandas as pd
import sqlalchemy
import openpyxl
import xlrd
import numpy
import math
from os import listdir
#Loading OOH functions
#from OOH_functions import *


file_list = os.listdir('C:/Users/artem.korolev/Desktop/Planeta/ESP')
for files in file_list:
    print(files)
    duplicates=[]
    ESP=[]

    #Reading ESP file
    ESP=[line.strip() for line in open(files, 'r')]

    ESP = pd.Series(ESP, name='ESPAR_ID')
    ESP=pd.concat([ESP], axis=1)

    #Reading SQL database
    print("Getting data from SQL")


    engine = sqlalchemy.create_engine("mssql+pyodbc://espar:espar@MSKSQLP01110/Odaplan?driver=SQL Server")

    #period1 = 2015
    df_SQL_period1=pd.read_sql("""SELECT  distinct b.[CITY3],a.[ESPAR_ID],a.[OWNER], b.[X1], b.[Y1],b.[TYPE_GENER], b.[SIZES],b.[ABC]
                    FROM [Odaplan].[dbo].[OP] a join [Odaplan].[dbo].[OSDATA] b on a.[ESPAR_ID]=b.[ESPAR_ID] where a.[PROJECT_ID]=(SELECT [PROJECT_ID]
                    FROM [Odaplan].[dbo].[PROJECTS] where [NAME]='2015_1' ) and b.X1 is not null and b.[CITY3]='UFA'""",engine)

    #period2 = 2014
    df_SQL_period2=pd.read_sql("""SELECT  distinct b.[CITY3],a.[ESPAR_ID],a.[OWNER], b.[X1], b.[Y1],b.[TYPE_GENER], b.[SIZES],b.[ABC]
                    FROM [Odaplan].[dbo].[OP] a join [Odaplan].[dbo].[OSDATA] b on a.[ESPAR_ID]=b.[ESPAR_ID] where a.[PROJECT_ID]=(SELECT [PROJECT_ID]
                    FROM [Odaplan].[dbo].[PROJECTS] where [NAME]='2014_1' ) and b.X1 is not null and b.[CITY3]='UFA'""",engine)




    time2=[]
    time1=[]

    time2=location(df_SQL_period2, ESP, time2, 'lat2', 'lng2', 'Owner2', 'Sizes2', 'ABC2', 'City2')
    time1=location(df_SQL_period1,ESP, time1,'lat1','lng1','Owner1','Sizes1','ABC1','City1')

    #Creating Key columns in data from group
    #df_working['CITY'] + df_working['OWNER']+ df_working['SIZES']+ df_working['ABC']
    #df_working['CITY'] + df_working['SIZES']+ df_working['ABC']
    time1['Key2'] = time1['City2'] + time1['Sizes2']+ time1['ABC2']
    #df_working['CITY'] + df_working['SIZES']
    time1['Key3'] = time1['City2'] + time1['Sizes2']



    df_SQL=df_SQL_period1


    #Creating Key column in data from SQL
    #df_SQL['CITY'] + df_SQL['OWNER']+ df_SQL['TYPE_GENER']+ df_SQL['SIZES']+ df_SQL['ABC']
    #df_SQL['CITY'] + df_SQL['TYPE_GENER']+ df_SQL['SIZES']+ df_SQL['ABC']
    df_SQL['Key2'] = df_SQL['CITY3'] + df_SQL['SIZES']+ df_SQL['ABC']
    #df_SQL['CITY'] + df_SQL['TYPE_GENER']+ df_SQL['SIZES']
    df_SQL['Key3'] = df_SQL['CITY3'] + df_SQL['SIZES']
    print(False in list(time1.lng1.notnull()))

    if False in list(time1.lng1.notnull()):

        df_working=time1[time1.lng1.notnull() == False]

        df_working=df_working[['ESPAR_ID','lat2','lng2','Owner2','Sizes2','ABC2','City2','Key2','Key3']]
        df_working.columns = ['ESPAR_ID','lat','lng','Owner','Sizes','ABC','City','Key2','Key3']
        df_working.reset_index(drop=True,inplace =True)



        #Running loop with KEY1

        #Running loop with KEY2
        df_working=KEY_loop(df_working, df_SQL, 'Key2', 'Distan2', 'ESPAR_ID2')

        #duplicates2=KEY_loop_replacing_uniq(df_working2, df_SQL, 'Key2', 'Distan2', 'ESPAR_ID2')
        df_working3=KEY_loop(df_working, df_SQL, 'Key3', 'Distan3', 'ESPAR_ID3')
        #duplicates3=KEY_loop_replacing_uniq(df_working3, df_SQL, 'Key3', 'Distan3', 'ESPAR_ID3')

        df_working3['Gatherer']=''
        df_working3['Gatherer']=df_working3['ESPAR_ID2']
        #df_working3['Gatherer']=df_working3['Gatherer'].fillna(df_working3['ESPAR_ID2'])
        df_working3['Gatherer']=df_working3['Gatherer'].fillna(df_working3['ESPAR_ID3'])
        df_working3=df_working3[['Gatherer','ESPAR_ID']]

        New = pd.merge(left=time1,right=df_working3, how='outer', left_on='ESPAR_ID', right_on='ESPAR_ID')
        New['Gatherer']=New['Gatherer'].fillna(New['ESPAR_ID'])
        New=New[['Gatherer','lat2','lng2','Owner2','Sizes2','ABC2','City2','Key2','Key3']]
        New.columns = ['Gatherer','lat','lng','Owner2','Sizes2','ABC2','City2','Key2','Key3']


        Distance = 0.00622166666
        #Distance = 0.02
        duplicates=KEY_loop_replacing_uniq(New, df_SQL, 'Key3', 'Distan3', 'Gatherer',Distance)
        duplicates=duplicates[['Gatherer']]

        #writer = pd.ExcelWriter('C:/Users/artem.korolev/Desktop/MEGA/Dupl1.xlsx')
        #duplicates.to_excel(writer,'Sheet1')
        #writer.save()
        path = 'C:/Users/artem.korolev/Desktop/Planeta/Planeta/' + files
        f = open(path,'w')
        for esparids in duplicates['Gatherer']:
            f.write(esparids+'\n')
        f.close()
    else:
        path = 'C:/Users/artem.korolev/Desktop/Planeta/Planeta/' + files
        f = open(path,'w')
        for esparids in ESP['ESPAR_ID']:
            f.write(esparids+'\n')
        f.close()




