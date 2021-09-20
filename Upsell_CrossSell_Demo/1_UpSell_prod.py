import pandas as pd
import numpy as np
from sklearn.preprocessing import MinMaxScaler

Input_product = 'PA-5220'

print("Reading Data Ticket Dump...")
sophos = pd.read_excel("Case Data Dump from Jan'21 to Jul'21.xlsx")

print("Reading Products Class file...")
product_group = pd.read_excel("Upsell-Data.xlsx")

# print("Reading Config file...")
# Input_upsell = pd.read_csv("Input_upsell.csv", delimiter= '\t')

#table = pd.pivot_table(sophos, values ='Case Number', index =['Account Name: Account Name', 'Ass_years'], columns =['Product'], aggfunc = "count")
table = pd.pivot_table(sophos, values ='Case Number', index =['Account Name: Account Name'], columns =['Product'], aggfunc = "count")
df1 = pd.DataFrame(table.to_records())

print(f"Number of unique Accounts: {len(df1)}")
prod_type = product_group[product_group['Product '] == Input_product]['Type'].tolist()[0]
pro_1 = product_group[product_group['Type'] == prod_type]
# pro_1 = product_group[product_group['Type'] == Input_upsell['Input Values'][0]]
pro_2 = pro_1[pro_1['Product '] == Input_product]
Rank_Target = pro_2['Rank'].tolist()[0]

Target_products = pro_1['Product '][pro_1['Rank'] < Rank_Target].tolist()

Upsell_products = pro_1['Product '][pro_1['Rank'] > Rank_Target].tolist()

Upsell_products1 = []
Target_products1 = []

for prods in Upsell_products:
    if prods in df1.columns.tolist():
        Upsell_products1.append(prods)

for prods in Target_products:
    if prods in df1.columns.tolist():
        Target_products1.append(prods)

print(f"\nInput product: {Input_product}")
print(f"\nProducts to Upsell: {Upsell_products1}")
print(f"\nTarget Accounts with Products: {Target_products1}")

dfx = df1.copy()
for Upsell in Upsell_products1:
    dfx = dfx[dfx[f'{Upsell}'].isnull()]
    print(f"Number of unique Accounts after removing {Upsell}: {len(dfx)}")
    
    
dfy = dfx[Target_products1]
dfx['#Targets'] = dfy.apply(lambda x: x.sum(), axis='columns')
dfx = dfx[dfx['#Targets'] > 0]
print(f"Number of unique Accounts using Target products: {len(dfx)}")


dfx['Pro_diversity'] = dfx.apply(lambda x: x.notnull().sum()-3, axis='columns')
dfx = dfx.sort_values('#Targets', ascending=False)

dfx = dfx.replace(np.nan, 0, regex=True)
#dfx1 = dfx[['Account Name: Account Name','#Targets','Pro_diversity', 'Ass_years']]
dfx1 = dfx[['Account Name: Account Name','#Targets','Pro_diversity']]

scaler=MinMaxScaler(feature_range=(1,100))
dfx1['#Targets'] = scaler.fit_transform(dfx1[["#Targets"]])
dfx1['Pro_diversity'] = scaler.fit_transform(dfx1[["Pro_diversity"]])

#dfx1["Rank"] = dfx1[['#Targets', 'Pro_diversity','Ass_years']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)
dfx1["Rank"] = dfx1[['#Targets', 'Pro_diversity']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)

dfx1 = dfx1.sort_values("Rank")

print(f"\nThe top 5 Recommended Accounts to Upsell:{Upsell_products1}")
print(dfx1.head(5))

dfx1.to_csv("1_UpSell_prod_output.csv", index = False)