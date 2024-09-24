import pandas as pd
import numpy as np
from tkinter.filedialog import *

open = askopenfilename(filetypes=[("xlsx files", "*.xlsx")])
df = pd.read_excel(open)
df.columns = df.iloc[5]
df = df.dropna(axis='columns',how='all')
df = df.drop(columns=df.columns[1],axis=1)
df.drop(labels=[0,1,2,3,4,5,6], axis=0, inplace=True)
convert = df[df['МОЛ'].notnull()]
index_df_os = 7

po_list = ['Windows', 'Excel', 'Antivirus', 'Office', 'Server CAL', 'Rmt Dsctp']
df.loc[:, po_list] = 'Нет'
def konvert(enum):
    global index_df_os
    po_list_name = ''
    if len(df[df['Инв номер'] == enum].index) != 0:
        index_df = df[df['Инв номер'] == enum].index[0]
        if index_df in convert['МОЛ'].index:
            index_df_os = index_df
        else:
            name = df.loc[index_df]['Инв номер']
            for i in po_list:
                if name.find(i) != -1:
                    po_list_name = i
            if not po_list_name in po_list:
                if not name in df.columns:
                    df.loc[:, name] = 'Нет'
                df.loc[index_df_os][name] = 'Да'
                df.drop(labels=[index_df], axis=0, inplace=True)
            else:
                df.loc[index_df_os][po_list_name] = name
                df.drop(labels=[index_df], axis=0, inplace=True)

def sverka_glpi(invent):
    print(invent)
    invent = str(invent)
    for i in df_glpi['Наименование']:
        if invent.lower().find(i[:2].lower()) != -1:
            if invent.lower().find(i[-6:-2]) != -1:
                return 'Yes'

df['Инв номер'].apply(konvert)
open = askopenfilename(filetypes=[("csv files", "*.csv")])
df_glpi = pd.read_csv(open, delimiter=';')
df['GLPI'] = df['Инв номер'].apply(sverka_glpi)


save = asksaveasfilename(filetypes=[("xlsx files", "*.xlsx")])
df.to_excel(f'{save}.xlsx',index=False)