import xlrd
import openpyxl
import numpy
import math
import openpyxl
import xlrd
import urllib
import urllib.request
import re
from openpyxl import load_workbook
from time import sleep
import pandas as pd

#Checking dictionaries and replacing names in working_file
def dictionary_check(working_file, dictionary_add, name):
    #Reading City_index
    dictionary_add = xlrd.open_workbook(dictionary_add)
    sheet = dictionary_add.sheet_by_name('Sheet1')
    # read City Index from Excel and replace cities in EXCEL and SQL
    keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]
    dict_list = []
    for row_index in range(1, sheet.nrows):
        d = {keys[col_index]: sheet.cell(row_index, col_index).value for col_index in range(sheet.ncols)}
        a=d.get(name)
        b=d.get('abc')
        working_file[name]=working_file[name].replace(a, b)
        #df_SQL['CITY']=df_SQL['CITY'].replace(a, b)
        dict_list.append(d)
    return working_file




#Looking for ESPAR_ids using KEYs and coordinates
def KEY_loop(df_working, df_SQL, KEY, Distan, ESPAR_IDs):
    #Running loop with KEY1
    print(KEY)
    a=0
    df_working[Distan] = numpy.nan
    df_working[ESPAR_IDs] = ''
    for all_SQL_rows in df_SQL[KEY]:
        b=0
        for all_working_rows in df_working[KEY]:
            if (all_SQL_rows == all_working_rows and math.isnan(df_working.iloc[b]['lat']) != True):
                Distance=numpy.sqrt((df_SQL.iloc[a]['Y1']-df_working.iloc[b]['lat'])**2+(df_SQL.iloc[a]['X1']-df_working.iloc[b]['lng'])**2)
                if math.isnan(df_working[Distan][b]):
                    df_working.set_value(b, Distan, Distance)
                    df_working.set_value(b, ESPAR_IDs, df_SQL.iloc[a]['ESPAR_ID'])
                elif df_working[Distan][b]>Distance:
                    df_working.set_value(b, Distan, Distance)
                    df_working.set_value(b, ESPAR_IDs, df_SQL.iloc[a]['ESPAR_ID'])
            b=b+1
        a=a+1
    return df_working


#searching for address in google and retrieve coordinates
def Google_search(df_working):
    df_working['Search'] =df_working['City'] +", " + df_working['Adress2']
    df_working['lat'] = numpy.nan
    df_working['lng'] = numpy.nan
    a=0
    for i in df_working['Search']:
        print(a)
        url = 'http://maps.google.com/maps/api/geocode/json?address='
        word=str(i)
        full_url = url+word
        print(full_url)
        if full_url != 'http://maps.google.com/maps/api/geocode/json?address=nan':
            url = urllib.parse.urlsplit(full_url)
            #print(url)
            russian = ''.join(re.findall('[А-Яа-я,.\/\\-\d]', url[3]))
            #print(russian)
            urlnew = urllib.parse.quote(russian)
            url_last=url[0]+"://"+url[1]+url[2]+"?address="+urlnew
            proxy_support = urllib.request.ProxyHandler({})
            opener = urllib.request.build_opener(proxy_support)
            urllib.request.install_opener(opener)
            with urllib.request.urlopen(url_last) as response:
                html = response.read()
            address = re.search('("formatted_address" : ")([^"]+)', str(html))
            lat = re.search('("lat" : )([0-9]{1,}[,.][0-9]{1,})', str(html))
            lng = re.search('("lng" : )([0-9]{1,}[,.][0-9]{1,})', str(html))
            print(word+' '+'address: ' + address.group(2)+' '+'lat: '+lat.group(2)+' '+'lng: '+lng.group(2))
            df_working.set_value(a, 'lat', lat.group(2))
            df_working.set_value(a, 'lng', lng.group(2))
        a=a+1
        sleep(0.3)
    return df_working