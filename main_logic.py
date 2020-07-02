import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
from classes import specie, pm_doc
import math
import logging
from theone_con import from_theone
logging.basicConfig(level=logging.INFO) # DEBUG, INFO, WARN

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
    'ROSPA0042' : 'https://docs.google.com/spreadsheets/d/1A_NPum_EEw5eHNO7AAqm4ZFkezFq07ZRQj32UMQp-YE/edit#gid=0'}



#Function for adding a new site
def add_site(name, link):
    gsheets[name] = link
    logging.info(name + ' was included on the list')

#Setting the authentification
def set_json(kf):
    kfl = kf
    logging.info('The new json file was set')



class site():
    def __init__(self, the_site):
        self.the_site = the_site
        #Authentification in gdrive
        self.kfl = 'Driveacc-017ff6c04252.json'
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(self.kfl, self.scope)
        self.gc = gspread.authorize(self.credentials)

        #Getting the table for the site from db
    def bf(self):
        self.bf = from_theone("select * from poim.sit_table('"+self.the_site+"')")
        self.bf.fillna('na', inplace=True)
        return self.bf

        #Getting the master table from gdrive
    def master(self):
        self.mst = self.gc.open_by_url(gsheets[self.the_site]).worksheet('master')
        self.master = self.mst.get_all_records()
        self.headers_m = self.master.pop(0)
        self.df_master = pd.DataFrame(self.master, columns = self.headers_m)
        self.df_master = self.df_master[self.df_master.id != '']
        self.df_master.fillna('na', inplace=True)
        return self.df_master

        #Getting the impacts table from gdrive
    def impacts(self):
        self.imp = self.gc.open_by_url(gsheets[the_site]).worksheet('presiuni_sit')
        self.impacts = self.imp.get_all_records()
        self.df_impacts = pd.DataFrame(impacts)
        self.df_impacts = self.df_impacts[self.df_impacts.impact != '']
        self.df_impacts.fillna('na', inplace=True)
        return self.df_impacts

        #Getting the measures table from gdrive
    def masuri(self):
        self.mas = self.gc.open_by_url(gsheets[the_site]).worksheet('masuri_sit')
        self.masuri = self.mas.get_all_records()
        self.df_masuri = pd.DataFrame(self.masuri)
        self.df_masuri = self.df_masuri[self.df_masuri.masura != '']
        self.df_masuri.fillna('na', inplace=True)
        return self.df_masuri

        #Getting the descriptions table from gdrive
    def descrieri(self):
        self.desc = self.gc.open_by_url(gsheets[self.the_site]).worksheet('descrieri')
        self.descrieri = self.desc.get_all_records()
        self.df_desc = pd.DataFrame(descrieri)
        self.df_desc.dropna(inplace=True)
        self.df_desc.fillna('na', inplace=True)
        return self.df_desc



doc_parts = {
    "Species header": 'sp_header',
    "Descriptive table A": 'ftblA_sp', 
    "Descriptive table B": 'ftblB_sp',
    "Conservation table A": 'ftblA_cons',
    "Matrix 1 table": 'ftblM_1',
    "Conservation table B": 'ftblB_cons',
    "Matrix 2 table": 'ftblM_2',
    "Matrix 3 table": 'ftblM_3',
    "Conservation table C": 'ftblC_cons',
    "Matrix 4 table": 'ftblM_4',
    "Matrix 5 table": 'ftblM_5',
    "Matrix 6 table": 'ftblM_6',
    "Conservation table D": 'ftblD_cons',
    "Matrix 7 table": 'ftblM_7',
    "Measures chapter": 'oss',
    "Measures chapter heading": 'chapter_h1',
    "Descriptive chapter heading": 'chapter_h1',
    "Conservation chapter heading": 'chapter_h1'
    }

