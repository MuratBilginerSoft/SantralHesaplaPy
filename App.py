from itertools import product
from datetime import datetime
import pandas as pd
import numpy as np
import os

min = float(input("Minimum Değer: "))
max = float(input("Maximum Değer: "))

# Excel dosyasının tam yolu
rootPath = "C:\\Users\\DMR\\Desktop\\Hesaplama\\"

fileName = "Mert.xlsx"

sourcePath = rootPath + fileName

df = pd.read_excel(sourcePath)

df = df.fillna(0)

print("İşlem Başladı")

columns = df.columns

# Tüm kombinasyonları hesapla
combinations = list(product(*[df[col] for col in columns]))

# Kombinasyonların çarpımlarını hesapla ve bir sözlüğe ekle
product_dict = {comb: round(np.prod(comb), 2) for comb in combinations}

new_df = pd.DataFrame(list(product_dict.items()), columns=['Combination', 'Product'])

# Orijinal DataFrame'e yeni DataFrame'i ekle
df = pd.concat([df, new_df], axis=1)

df.drop(df[df['Product'] == 0].index, inplace=True)

df.drop(df[df['Product'] <= min].index, inplace=True)
df.drop(df[df['Product'] >= max].index, inplace=True)


now = datetime.now()
date_string = now.strftime("%Y-%m-%d_%H-%M-%S")

resultExcelName =  "MertSonuc_" + date_string +".xlsx"

resultPath = rootPath + resultExcelName 

df.to_excel(resultPath, index=False)

excelExePath = '"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"'

os.system(f'start {excelExePath} {resultPath}')

print("İşlem Tamamlandı")