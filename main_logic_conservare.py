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


#Open the text file for gis
f=open(the_site + '_gis_final.txt','w+', encoding="utf-8")

                #-------CREATION OF THE CONSERVATION TABLES------

#Itering through rows and creating species objects&tables
for row in df_rows:
    #Initialize sp object 
    sp = specie(row, df_master, bf, df_desc, df_impacts, df_masuri)
    f.write(str(sp.cod_sp) +','+ str(sp.lat_sp) + ',' + str(sp.cod_n)+'-'+str(sp.feno) + ',' + str(sp.d3)+'\n')

    #Put the species header
    doc.sp_header(sp)

    #Create the population based conservation chapter
    doc.chapter_h2('Evaluarea stării de conservare a speciei din punctul de vedere al populației speciei')
    #Add the table title
    doc.chapter_h3('Tabelul A: Parametri pentru evaluarea stării de conservare a speciei din punct de vedere al populației')
    #Create the population based conservation tabels (tabels A)
    doc.ftblA_cons()
    doc.ftblA_cons_fill(sp)
    #Adding matrices tables
    doc.chapter_h3('Matricea 1) Matricea de evaluare a stării de conservare a speciei din punct de vedere al populației speciei')  
    doc.ftblM_1()
    doc.ftblM_1_fill(sp)

    #Create the habitat based conservation chapter
    doc.chapter_h2('Evaluarea stării de conservare a speciei din punctul de vedere al habitatului speciei')
    #Add the table title
    doc.chapter_h3('Tabelul B: Parametri pentru evaluarea stării de conservare a speciei din punct de vedere al habitatului speciei')
    #Create the habitat based conservation tabels (tabels B)
    doc.ftblB_cons()
    doc.ftblB_cons_fill(sp)
    #Adding matrices tables
    doc.chapter_h3('Matricea 2) Matricea pentru evaluarea tendinței globale a habitatului speciei')  
    doc.ftblM_2()
    doc.ftblM_2_fill(sp)   
    doc.chapter_h3('Matricea 3) Matricea de evaluare a stării de conservare a speciei din punct de vedere al habitatului speciei')  
    doc.ftblM_3()
    doc.ftblM_3_fill(sp)

    #Create the perspectives based conservation chapter
    doc.chapter_h2('Evaluarea stării de conservare a speciei din punctul de vedere al perspectivelor speciei')
    #Add the table title
    doc.chapter_h3('Tabelul C: Parametri pentru evaluarea stării de conservare a speciei din punct de vedere al perspectivelor speciei în viitor')
    #Create the perspectives based conservation tabels (tabels C)
    doc.ftblC_cons()
    doc.ftblC_cons_fill(sp)
    #Adding matrices tables
    doc.chapter_h3('Matricea 4) Matricea pentru evaluarea perspectivelor speciei din punct de vedere al populației speciei')  
    doc.ftblM_4()
    doc.ftblM_4_fill(sp)   
    doc.chapter_h3('Matricea 5) Perspectivele speciei în viitor, după implementarea planului de management actual')  
    doc.ftblM_5()
    doc.ftblM_5_fill(sp)
    doc.chapter_h3('Matricea 6) Matricea evaluării stării de conservare a speciei din punct de vedere al perspectivelor speciei în viitor, după implementarea planului de management actual')  
    doc.ftblM_6()
    doc.ftblM_6_fill(sp)

    #Create the global conservation chapter
    doc.chapter_h2('Evaluarea globală a speciei')
    #Add the table title
    doc.chapter_h3('Tabelul D: Parametri pentru evaluarea stării globale de conservare a speciei în cadrul ariei naturale protejate')
    #Create the global conservation tabels (tabels D)
    doc.ftblD_cons()
    doc.ftblD_cons_fill(sp)
    #Adding matrices tables
    doc.chapter_h3('Matricea 7) Evaluarea stării globale de conservare a speciei')  
    doc.ftblM_7()
    doc.ftblM_7_fill(sp)  

f.close()
                #-------CREATION OF THE MEASURES TABLES------


doc.chapter_h1('Măsurile de conservare pentru menținerea sau îmbunătățirea stării de conservare a speciilor de păsări')

doc.measures_ch(df_masuri, df_impacts)




#Salvation ^^ of the document
doc.save(the_site+'_export_final'+'.docx')


