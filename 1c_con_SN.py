import pandas as pd
#import numpy as np
from tkinter.filedialog import *

open = askopenfilename(filetypes=[("xlsx files", "*.xlsx")])
df = pd.read_excel(open)
work = pd.DataFrame()
pd.options.mode.chained_assignment = None
stolb = ['Филиал в городе Шымкент', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4']

def otdele(colum_1): # Получить список отделений
    if 'Отделение' in colum_1 or 'УК ШФ' in colum_1 :
        return True
    else:
        return False


def add_otdel(colum): # Присвоить отделения к списку Инвентарников
    otdelen = df[df['Филиал в городе Шымкент'].apply(otdele)]
    otdelen = otdelen['Филиал в городе Шымкент']
    i = work.index[work['Инвентарный'] == colum].tolist()[0]
    end_index = len(otdelen.index)
    for start_index in range(0,end_index + 1):
        if start_index != end_index:
            if i < otdelen.index[start_index]:
                name_otdel = otdelen[otdelen.index[start_index - 1]]
                return str(name_otdel)
        else:
            name_otdel = otdelen[otdelen.index[end_index - 1]]
            return str(name_otdel)


def MOL(name): # Получить список ФИО
    i = work.index[work['Инвентарный'] == name].tolist()[0] + 2
    fio = str(df[df.index == i]['Филиал в городе Шымкент'])
    fio = fio.split('\n')[0].split(str(i))[1].strip()
    return fio

df.drop(labels=[0,1,2,3,4,5], axis=0, inplace=True) # Удалить первые 5 строк
work = df[df['Unnamed: 1'].notnull()] # фильтрация списка по ОС
work.drop(labels=['Unnamed: 4', 'Unnamed: 2'], axis=1, inplace=True) # Удалить столбцы 2,3
work.columns = ['Инвентарный', 'ОС', 'SN']
#test1 = pd.core.series.Series(data=['1'], index=[2])
work['ФИО'] = work['Инвентарный'].apply(MOL)
work['Отделение'] = work['Инвентарный'].apply(add_otdel)
#save = '321.xlsx'
save = asksaveasfilename(filetypes=[("xlsx files", "*.xlsx")])
work.to_excel(f'{save}.xlsx',index=False)

