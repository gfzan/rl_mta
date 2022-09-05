# -*- coding: utf-8 -*-
"""
Created on Fri Aug 26 10:13:28 2022

@author: gfzan
"""

# Flusso Informativo Sanitario RL-MTA per invio su portale SMAF

# importo librerie 
from datetime import datetime, date

import pandas as pd
import numpy as np
#import bamboolib as bam
import os, glob
import xlrd

# imposto caratteristiche display

#pd.options.display.max_colwidth = 100
#pd.options.display.max_colwidth=30
pd.set_option('display.max_rows', 50)
#pd.set_option('display.max_columns', 10)
pd.set_option('display.width', 300)
pd.set_option('display.max_colwidth', None)


## definisco alcune strutture dati utili

# MAPPA CODICE E DESCRIZIONE EROGATORE
dict_Erogatore = {'726003482':'Poli Crema',
                  '726003485':'Poli Rivolta',
                  '726003489':'Poli Soncino',
                  '726003491':'Poli Castelleone'}

# DEFINISCE I LIMITI DELLE CLASSI DI PRIORITA
dict_Parametri_Priorita = {'U':3,
                           'B':10,
                           'D':{'visite':30,
                               'strumentali':60},
                           'P':120
                            }
## variabili globali

intColumns = ['Chiave', 'N_utenti_inizio', 'N_prenotanti', 'N_follow_screening', 'N_URGENTI', 'N_Prima_Diagnosi', 'Flag_Virtuale']
dataColumns = ['Data_Prescrizione', 'Data_rilevazione', 'Data_ins_agenda', 'Prima_data_prospettata', 'Data_assegnata', 'Data_Rilevazione']
flusso_D1_cols = ['Chiave', 'Codice_Fiscale', 'Identificativo_PAI']
flusso_D2_cols = ['Chiave', 'Codice_Erogatore', 'Codice_Prestazione', 'Codice_Prescrizione', 'Codice_IUP', 'Codice_Priorita', 'Data_Prescrizione', 'Data_rilevazione', 'Data_ins_agenda', 'N_prenotazione', 'Flag_Cronicità', 'Prima_data_prospettata', 'Data_assegnata', 'Flag_Tolleranza', 'Flag_Approfondimento',
       'Flag_Virtuale', 'Garanzia_TMAX']
flusso_M_cols = ['Codice_ATS', 'Codice_Erogatore', 'Codice_Prestazione', 'Data_Rilevazione', 'N_utenti_inizio', 'N_prenotanti', 'N_follow_screening', 'N_URGENTI', 'N_Prima_Diagnosi']


class RL_MTA:
        
    def __init__(self, path, type="csv"):
        # type può avere valori ["csv", "xls"] 
        if type == "csv":
            os.chdir(path)
            result = glob.glob('*.{}'.format(type))
            for file in result:
                if 'D1' in file:
                    self.Flusso_D1 = pd.read_csv(file, sep=';', dtype=object, keep_default_na=False)
                elif 'D2' in file:
                    self.Flusso_D2 = pd.read_csv(file, sep=';', dtype=object, keep_default_na=False)
                    for col in dataColumns:
                        if col in self.Flusso_D2.columns:             
                            try:
                                self.Flusso_D2[col] = self.Flusso_D2[col].apply(lambda x: '' if x == 'nan' else x[4:]+'-'+x[2:4]+'-'+x[0:2]).replace('--','')
                            except:
                                print('campi date non modificati come XX-XX-XXXX; verificare la causa')
                                print('colonna non trovata in D2,csv : ', col)
                                pass
                        else:
                            pass
                        
                elif 'M' in file:
                    self.Flusso_M = pd.read_csv(file, sep=';', dtype=object, keep_default_na=False)
                    for col in dataColumns:
                        if col in self.Flusso_M.columns:
                            try:
                                self.Flusso_M[col] = self.Flusso_M[col].apply(lambda x: '' if x == 'nan' else x[4:]+'-'+x[2:4]+'-'+x[0:2]).replace('--','')
                            except:
                                print('campi date non modificati come XX-XX-XXXX; verificare la causa')
                                print('colonna non trovata in M.csv : ', col)
                                pass
                        else:
                            pass
                else:
                    print('NON sono stati caricati tutti i dati dei file D1, D2, M; verificare la causa')
                    try:
                        print('shape D1   =  D1.shape')
                    except:
                        print('file D1 non caricato')
                    try:
                        print('shape D2   =  D2.shape')
                    except:
                        print('file D2 non caricato')
                    try:
                        print('shape M    =  M.shape')
                    except:
                        print('file M non caricato')
                    pass
                
                print('file elaborato : ', file)
                
        elif type =="xls":
            os.chdir(path)
            result = glob.glob('*.{}'.format(type))
            for file in result:
                #Dati = pd.ExcelFile(file).parse() #use r before absolute file path
                xls = xlrd.open_workbook(file, on_demand=True)
                #print (xls.sheet_names()) # <- remeber: xlrd sheet_names is a function, not a property                
                for sheet in xls.sheet_names():
                    if 'D1' in sheet:
                        self.Flusso_D1 = pd.read_excel(file, sheet_name=sheet, header=0, index_col=None, dtype=object, na_filter = False)
                    elif 'D2' in sheet:
                        self.Flusso_D2 = pd.read_excel(file, sheet_name=sheet, header=0, index_col=None, dtype=object, na_filter = False)
                    elif 'M' in sheet:
                        self.Flusso_M = pd.read_excel(file, sheet_name=sheet, header=0, index_col=None, dtype=object, na_filter = False)
                    else:
                        pass
                    print(sheet, file)
                    
        else:
            pass
        
        
        ## COSTRUISCO FLUSSO IN UNICA TABELLA
        
        self.Flusso_D = pd.merge(self.Flusso_D1, self.Flusso_D2, on=["Chiave"])
        self.Flusso_D.drop(self.Flusso_D.filter(regex='Unname').columns, axis=1, inplace=True)
        self.Flusso_M.drop(self.Flusso_M.filter(regex='Unname').columns, axis=1, inplace=True)
        
        self.Flusso = pd.merge(self.Flusso_D, self.Flusso_M, on=["Codice_Erogatore", "Codice_Prestazione"])
        #self.Flusso.drop("Data_rilevazione", axis=1, inplace=True)
        
        
        # INSERISCO ULTIMA STRUTTURA DATI: NOMENCLATORE PRESTAZIONE DA FILE REG. LOMB.
        # seguo path relativo rispetto a quello caricato nel costruttore
    
        self.Nomenclatore = pd.read_excel('../DatiAccessori/prest_AMB_da+PUBBLICARE+marzo+2022+.xlsx',\
                   sheet_name='Nomenclatore_Tariffario',\
                   usecols=['codice_senza_punto', 'descr_prestaz breve'],\
                   header=0, dtype=str).rename(columns = {'codice_senza_punto':'Codice_Prestazione', 'descr_prestaz breve':'Descrizione Prestazione'})
            
        self.presidio = pd.DataFrame(dict_Erogatore.items(), columns=['Codice_Erogatore', 'Presidio'])
        #self.presidio = pd.DataFrame(dict_Erogatore.items(), columns=['Codice_Erogatore', 'Presidio']).astype({'Codice_Erogatore': 'int64'})
        self.priorita = pd.DataFrame.from_dict(dict_Parametri_Priorita, orient='index', columns=['GG'])
        self.FlussoAll = (self.Flusso.merge(self.Nomenclatore, on=["Codice_Prestazione"])).merge(self.presidio, on='Codice_Erogatore')
        
        
        ## AGGIORNO I TIPI SULLE COLONNE "int" E "data"
        
        #intColumns = ['Chiave', 'N_utenti_inizio', 'N_prenotanti', 'N_follow_screening', 'N_URGENTI', 'N_Prima_Diagnosi']
        self.FlussoAll[intColumns] = self.FlussoAll[intColumns].astype(int)
        
        #dataColumns = ['Data_Prescrizione', 'Data_rilevazione', 'Data_ins_agenda', 'Prima_data_prospettata', 'Data_assegnata', 'Data_Rilevazione']

        # problema overflow (out of bound) formato data int64[ns] -> elabora i giorni di 584 anni da 1677 a 2262
        for col in dataColumns:
            try:
                self.FlussoAll[col] = self.FlussoAll[col].astype(str).str.replace('2999-12-31', '1799-12-31')
                #self.FlussoAll[col] = self.FlussoAll[col].apply(lambda x: datetime.strptime(x,'%Y-%m-%d').date())
                self.FlussoAll[col] = pd.to_datetime(self.FlussoAll[col], format='%Y-%m-%d')
            except:
                pass
        
        # aggiungo la colonna con il calcolo dei giorni di attesa
        # CONDIZIONE SULLA SCELTA DELLE DATE DA SOTTRARRE PER LA LISTA ATTESA
        
        self.FlussoAll['Attesa(GG)'] = np.where(self.FlussoAll.Prima_data_prospettata.isna(),\
                                    self.FlussoAll.Data_assegnata - self.FlussoAll.Data_Rilevazione,\
                                    self.FlussoAll.Prima_data_prospettata - self.FlussoAll.Data_Rilevazione)
        self.FlussoAll['Attesa(GG)'] = self.FlussoAll['Attesa(GG)'].astype('timedelta64[D]').astype(int)
        
        ## INSERT BOUND COLUMN
        
        self.FlussoAll['Bound'] = ''

        for index, row in self.FlussoAll.iterrows():
            if row.Codice_Priorita=='U':
                self.FlussoAll['Bound'][index]=3
                #print(index, 'URG')
            elif row.Codice_Priorita=='B':
                self.FlussoAll['Bound'][index]=10
                #print(index,'breve')
            elif row.Codice_Priorita=='P':
                self.FlussoAll['Bound'][index]=120
                #print(index,'PROGRAMMABILE')
            elif row.Codice_Priorita=='D' and 'VISITA' in row['Descrizione Prestazione']:
                self.FlussoAll['Bound'][index]=30
                #print(index,'VIsita')
            else:
                self.FlussoAll['Bound'][index]=60        
                #print(index,'Strumentale')
        
        print('il "costruttore" ha implementato le seg. strutture dati:')
        print(' - Flussi: [M, D1, D2]')
        print(' - Flusso_D [merge]')
        print(' - FlussoAll dal merge di [D + M + Presidio + Nomenclatore]')
        print(' - Aggiunte le colonne sui t. di attesa e gli "out of bound"')


    def out_of_bound(self, df):
        df['out_of_bound'] = np.where((df['Codice_Priorita']=='P') & (df['Attesa(GG)']>120), 1, 0)
        df['out_of_bound'] = np.where(((df['Codice_Priorita']=='D') & (df['Descrizione Prestazione'].str.contains('VISITA'))) & (df['Attesa(GG)']>30), 1, df['out_of_bound'])
        df['out_of_bound'] = np.where(((df['Codice_Priorita']=='D') & (df['Descrizione Prestazione'].str.contains('VISITA')==False)) & (df['Attesa(GG)']>60), 1, df['out_of_bound'])
        df['out_of_bound'] = np.where((df['Codice_Priorita']=='B') & (df['Attesa(GG)']>10), 1, df['out_of_bound'])
        df['out_of_bound'] = np.where((df['Codice_Priorita']=='U') & (df['Attesa(GG)']>3), 1, df['out_of_bound'])
        return df
    
    def removeDataND(self, df, data='1799-12-31'):
        print('dimensioni iniziali data_frame  = ', df.shape)
        selToRemove = df.loc[df['Data_assegnata'].astype('string').isin(['1799-12-31'])]
        F_mod = df.drop(selToRemove.index)
        print('dimensioni finali data_frame  =   ', F_mod.shape)
        
        ## ... e modifico il file M per invio
        
        selToRemoveGroupby = selToRemove.groupby(['Codice_Erogatore', 'Codice_Prestazione']).agg(count=('Codice_Prestazione', 'size'))
        #print(selToRemoveGroupby)
        
        M_mod = self.Flusso_M.copy()
        for index, row in selToRemoveGroupby.iterrows():
            # mappo le righe da modificare sulla selezione dei record cancellati DF Flusso
            val = self.Flusso_M['N_prenotanti'][(self.Flusso_M['Codice_Erogatore'] == row.name[0]) & (self.Flusso_M['Codice_Prestazione'] == row.name[1])]
            #print(self.Flusso_M.loc[val.index]['N_prenotanti'])
            if int(val)- row[0]==0:
                #print((self.Flusso_M_mod.loc[val.index, ['Codice_Erogatore','Codice_Prestazione']]),'  indice eliminato')
                M_mod.drop(val.index, inplace=True)
            else:
                M_mod.at[val.index, 'N_prenotanti'] = str(int(val)- row[0])
            try:
                print(M_mod.loc[val.index]['N_prenotanti']+'  indice aggiornato')
            except:
                pass
                
        return (F_mod, M_mod, selToRemove)
    
    def saveCSVfile(self, flussoF_mod, flussoM_mod):
        ## devo passare le strutture dati D ed M modificate
        
        #path =input('inserire path file dati per flusso:')
        path = '../DatiInvioSMAF'
        if not os.path.exists(path):
            os.makedirs(path)
        
        # cols = flussoD_mod.columns.values.tolist()
        # cols_D1 = cols[0:3]
        # cols_D2 = cols[3:]
        # cols_D2.insert(0, cols[0])

        flussoF_mod[flusso_D1_cols].to_csv(path +'/'+ 'D1.csv',sep=';', index=False)
        flussoF_mod[flusso_D2_cols].to_csv(path +'/'+ 'D2.csv',sep=';', index=False)
        flussoM_mod.to_csv(path +'/'+ 'M.csv',sep=';', index=False)

## ORA FACCIO I REPORT PER IL RUA E LA SEGRETERIA DELLA DMP

    def reportRUA(self, df_F_mod):
        ## prima versione tabella pivot
        df_F_mod_groupby = df_F_mod.groupby(['Codice_Prestazione', 'Descrizione Prestazione', 'Presidio', 'Codice_Priorita']).agg(count=('Codice_Prestazione', 'size'),
                                        mean_Attesa=('Attesa(GG)', 'mean'),
                                        n_Flag_V=('Flag_Virtuale', 'sum'),
                                        out=('out_of_bound', 'sum')).astype(int)
        
        #AGGIUNGO colonna "out%"
        df_F_mod_groupby['out %'] = (100*df_F_mod_groupby['out']/df_F_mod_groupby['count']).map('{:,.1f}%'.format)
        
        ########################################################
        # CAMBIO VALORE -81273 in ND sul valore medio di attesa
        df_F_mod_groupby['mean_Attesa'] = df_F_mod_groupby['mean_Attesa'].replace(-10, 'ND')
        ########################################################
        
        ## RINOMINO LE COLONNE DEL DF PER AVERE INTESTAZIONI DI COLONNE MIGLIORI NEL REPORT
        
        df_Mod_Renamed = df_F_mod.rename(columns={'Codice_Priorita':'Priorità', 'Codice_Prestazione':'Codice Prestazione', 
                                         'Flag_Virtuale':'n° Flag Virtuale', 'out_of_bound':'n° out of bound'})
        
        ## seconda versione tabella pivot (migliore)
        pivot = pd.pivot_table(df_Mod_Renamed, values=['Attesa(GG)', 'n° Flag Virtuale', 'n° out of bound'], 
                       index=['Codice Prestazione', 'Descrizione Prestazione','Presidio', 'Priorità'],
                       aggfunc={'Attesa(GG)':[len, np.mean],
                                'n° Flag Virtuale':sum,
                                'n° out of bound':sum})
        pivot['Attesa(GG)'] = pivot['Attesa(GG)'].astype(int).replace(-10, 'ND')
        pivot['out %'] = (100*(pivot['n° out of bound']['sum'])/(pivot['Attesa(GG)']['len'])).map('{:,.1f}%'.format)
        
        #pivot.reset_index(inplace=True)        ## non si fa su tabella multi indice
       
        
        ## metodo "style apply" su dataframe groupby per formattazione con codice colore dei valori di attesa
        ## metodo: color_cells_df()
        
        df_F_mod_groupby = df_F_mod_groupby.reset_index()
        df_F_mod_groupby.rename(columns={'Codice_Priorita':'Priorità', 'mean_Attesa':'Attesa media(gg)', 'Codice_Prestazione':'Codice Prestazione', 
                                         'n_Flag_V':'n° Flag Virtuale', 'out':'n° out of bound'}, inplace=True)
        df_F_mod_groupby_styleApply = df_F_mod_groupby.style.apply(self.color_cells_df, subset=['Descrizione Prestazione', 'Priorità', 'Attesa media(gg)', 'out %'], axis=1)
        
        pivot_styleApply = pivot.style.set_table_styles([{'selector' : 'th',
                                       'props' : [('border','2px solid black')],
                                       'selector' : 'tr',
                                       'props' : [('border','2px solid green')]}], overwrite=True)\
                                        .apply(self.color_cells_pivot, axis=1)
         
        ## HTLM
        
        #path =input('inserire path file Report HTML:')
        path = '../ReportHTML'
        if not os.path.exists(path):
            os.makedirs(path)
        
        textFile = open(path +'/' + 'Report_Interno.html', 'w')

        textFile.write(
            "<html>\n \
                <head>\n \
                    <title>RL-MTA Liste di Attesa</title>\n \
                    <link rel='stylesheet' href='Style_Pivot.css'>\n \
                </head>\n \
                <body>\n \
                    <h1>RL-MTA Liste di Attesa</h1>\n \
                    <h2>MONITORAGGIO INTERNO</h2>\n \
                    <p> Rilevazione del :  " + str(df_F_mod['Data_Rilevazione'][3].date().strftime('%d/%m/%Y')) + "</p>\n")
        
        textFile.write(pivot_styleApply.to_html())
        
        textFile.write("</body></html>")
        
        textFile.close()          
        
        return(df_F_mod_groupby, df_F_mod_groupby_styleApply, pivot, pivot_styleApply)
    
    
    def reportDMP(self, df_F_mod):    ## uso il report prodotto per il RUA

        reportDmp = df_F_mod.groupby(['Presidio', 'Descrizione Prestazione', 'Codice_Priorita', 'Bound']).agg(mean_Attesa=('Attesa(GG)', 'mean')).astype(int)
        reportDmp = reportDmp.reset_index()
        reportDmp = reportDmp[reportDmp['Codice_Priorita'].isin(['B', 'D'])]
        reportDmp = reportDmp[['Presidio', 'Descrizione Prestazione', 'Bound', 'mean_Attesa']].sort_values(by=['Presidio','Descrizione Prestazione'], ascending=[True, True])
        
        reportDmp.rename(columns={'mean_Attesa': 'Attesa Media (classe B e D)', 'Bound':'Attesa Standard'}, inplace=True)
        
        ## HTML
        #path =input('inserire path file Report HTML:')
        path = '../ReportHTML'
        if not os.path.exists(path):
            os.makedirs(path)
            
        textFile = open(path +'/' + 'Report_Portale.html', 'w')

        textFile.write(
            "<html>\n \
                <head>\n \
                    <title>RL-MTA Liste di Attesa</title>\n \
                    <link rel='stylesheet' href='Style_Segreteria.css'>\n \
                </head>\n \
                <body>\n \
                    <h1>AMMINISTRAZIONE TRASPARENTE: Liste di Attesa</h1>\n \
                    <h2>PRESTAZIONI AMBULATORIALI e di RICOVERO</h2>\n \
                    <h3>art 32, CO;2 Lett B - D.Lgs 33/2013</h3>\n \
                    <p> Rilevazione del " + str(df_F_mod['Data_Rilevazione'][3].date().strftime('%d/%m/%Y')) + "</p>\n")
        
        textFile.write(reportDmp.to_html())
        
        textFile.write("\n<script type='text/javascript'>\n \
            window.onload = function(){\n \
                var rows = document.querySelectorAll('tbody tr'), i;\n \
                for(var row of rows){\n \
                    if (parseInt(row.cells[4].innerHTML) > parseInt(row.cells[3].innerHTML)){\n \
                        row.cells[4].style.backgroundColor = 'yellow';\n \
                        cell = row.insertCell();\n\
                        cell.innerHTML = 'X'\n\
                        }}}</script>")
        textFile.write("</body></html>")
        
        textFile.close() 
        
        return(reportDmp)
    
    
    def color_cells_df(self, x):
        color=''
        if x[2]=='ND':
            color = 'yellow'
            
        if x[3] !='0.0%':
            color = 'lightgreen'
            pass
        
        if x[1] == 'P' and x[2]>120:
             color = 'lightblue'
        elif x[1] == 'D':
            if 'VISITA' in x[0] and x[2]>30:
                color = 'lightblue'
            elif ~('VISITA' in x[0]):
                if x[2]>60:
                    color = 'lightblue'
                else: pass
        elif x[1] == 'B' and x[2]>10:
            color = 'lightblue'
        elif x[1] == 'U' and x[2]>3:
            color = 'lightblue'
        
        else: pass

        return ('','background-color: %s' % color,'background-color: %s' % color, 'background-color: %s' % color)
    
    
    def color_cells_pivot(self, x):
        color=''
        if x[1]=='ND':
            pass
        
        if x[4] !='0.0%':
            color = 'lightgreen'
            pass
        
        if x.name[3] == 'P' and x[1]>120:
            color = 'lightblue'
            
        elif x.name[3] == 'B' and x[1]>10:
            color = 'lightblue'
            
        elif x.name[3] == 'U' and x[1]>3:
            color = "lightblue"
            
        elif x.name[3] == 'D':
            if 'VISITA' in x.name[1] and x[1]>30:
                color = 'lightblue'
                
            elif ~('VISITA' in x.name[1]):
                if x[1]>60:
                    color = 'lightblue'
                else: pass 
    
        return ('','background-color: %s' % color, '','' ,'background-color: %s' % color)
        
        
        
if __name__ == "__main__":
    path =input('inserire path file dati per flusso:')
    tipoFile = input('inserire estensione file da leggere:')
    try: 
        datiStruct = RL_MTA(path, tipoFile)
    except:
        datiStruct = RL_MTA(r'./DatiLoadFlussoMTA', 'xls')
    print('Fatto!')
    
    #datiStruct = RL_MTA(path, tipoFile)
    
                
                

            

