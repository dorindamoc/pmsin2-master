import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
from classes import specie, pm_doc
import math
import logging
from theone_con import from_theone
logging.basicConfig(level=logging.INFO) # DEBUG, INFO, WARN



#Storing the links on gdrive for every site
gsheets = {
    'ROSPA0004' : 'https://docs.google.com/spreadsheets/d/1FRh6F8wh9iEcAOI_7gt-h7YDm38YaraEsMl7BCs9c_c/edit#gid=0',
    'ROSPA0012' : 'https://docs.google.com/spreadsheets/d/1oeUT7zVmWmNeIxOuCNX2xNZ8KNxlGA-_Mt7QNYW_Z1o/edit#gid=0',
    'ROSPA0112' : 'https://docs.google.com/spreadsheets/d/1nM5UR7hykwL0mffkkd6YQIvpZzHUPfJBEfrfxf_llHI/edit#gid=0',
    'ROSPA0111' : 'https://docs.google.com/spreadsheets/d/1LrmO_aMmvEyCzuAtsiejMjtSzADsAkQsIGuQPX6hxug/edit#gid=0',
    'ROSPA0109' : 'https://docs.google.com/spreadsheets/d/1H4epQ_iR9jZGugIlo7i9pGrP5Hmf8IocBiuRBUscKLA/edit#gid=0',
    'ROSPA0101' : 'https://docs.google.com/spreadsheets/d/1XNYEqTLr4EuL9Ghb3x8OhWno9VYqRgmtsNpV4icCX5o/edit#gid=0',
    'ROSPA0064' : 'https://docs.google.com/spreadsheets/d/1jxUQQU15eAe6HnLVuNGHDYbi-xFYuc5UbG2BT7eKKdA/edit#gid=0',
    'ROSPA0061' : 'https://docs.google.com/spreadsheets/d/1awwKOi1lY8KTQeRPp3D7BS_mLB3oBXrb98VgyJyvWeY/edit#gid=0',
    'ROSPA0051' : 'https://docs.google.com/spreadsheets/d/1ck3Q4dv55TdI7Uo98qZ4-E0-2-xxaWzq9WBL5X0SSKk/edit#gid=0',
    'ROSPA0042' : 'https://docs.google.com/spreadsheets/d/1A_NPum_EEw5eHNO7AAqm4ZFkezFq07ZRQj32UMQp-YE/edit#gid=0',
    'descrieri' : 'https://docs.google.com/spreadsheets/d/13qN_1-CXxNzcUNV52s8tERspjqxij4Rn0uyHSbPSebg/edit#gid=839134727'}

#Setting the authentification
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('Driveacc-017ff6c04252.json', scope)
gc = gspread.authorize(credentials)

#Getting the site
the_site = 'ROSPA0'+input('What are the last 3 numbers on your site code?  ')

#Getting the table for the site from db
bf = from_theone("select * from poim.sit_table('"+the_site+"')")
bf.fillna('na', inplace=True)

#Getting the master table from gdrive
mst = gc.open_by_url(gsheets[the_site]).worksheet('master')
master = mst.get_all_records()
headers_m = master.pop(0)
df_master = pd.DataFrame(master, columns = headers_m)
df_master = df_master[df_master.id != '']
df_master.fillna('na', inplace=True)

#Getting the impacts table from gdrive
imp = gc.open_by_url(gsheets[the_site]).worksheet('presiuni_sit')
impacts = imp.get_all_records()
df_impacts = pd.DataFrame(impacts)
df_impacts = df_impacts[df_impacts.impact != '']
df_impacts.fillna('na', inplace=True)

#Getting the measures table from gdrive
mas = gc.open_by_url(gsheets[the_site]).worksheet('masuri_sit')
masuri = mas.get_all_records()
df_masuri = pd.DataFrame(masuri)
df_masuri = df_masuri[df_masuri.masura != '']
df_masuri.fillna('na', inplace=True)

#Getting the descriptions table from gdrive
desc = gc.open_by_url(gsheets['descrieri']).worksheet('descrieri_tst')
descrieri = desc.get_all_records()
df_desc = pd.DataFrame(descrieri)
df_desc.dropna(inplace=True)
df_desc.fillna('na', inplace=True)

#Getting the list of the rows from the master table
df_rows = list(df_master.index)


#Creation of the document
doc = pm_doc()



                #-------CREATION OF THE POPULATION TABLES------

#Itering through rows and creating species objects&tables
for row in df_rows:
    #Initialize sp object 
    sp = specie(row, df_master, bf, df_desc, df_impacts, df_masuri)

    #Put the species header
    doc.sp_header(sp)

    #Create the description tabel (tabels A)
    doc.ftblA_sp()
    #Fill the description tabel (tabels A)
    doc.ftblA_sp_fill(sp)
    #Insert the image with the species
    doc.fimg(sp) #Add the image with the species

    #Create the specific tabels (tabels B)
    doc.ftblB_sp()
    #Fill the specific tabels (tabels B)
    doc.ftblB_sp_fill(sp)
    #Insert the map with the species
    doc.fmap(sp)


#Salvation ^^ of the document
doc.save(the_site+'_export_population'+'.docx')


