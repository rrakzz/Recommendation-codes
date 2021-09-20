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
#print(Input_upsell.columns.tolist())

df2 = df1[df1[f'{Input_product}'].notnull()]
print(f"Number of Accounts with {Input_product}: {len(df2)}")

Products = []
Count = []
for i,prod in enumerate(df2.columns.tolist()):
    if i>1:
        Products.append(prod)
        Count.append(len(df2[df2[f'{prod}'].notnull()]))
#         print(f"{prod}", len(df2[df2[f'{prod}'].notnull()]))

result = pd.DataFrame({"Product": Products, "Count": Count})
result = result.sort_values(f'Count', ascending=False)

scaler=MinMaxScaler(feature_range=(1,100))
result['%Count'] = scaler.fit_transform(result[["Count"]])

types, EOS = [], []
for val in result.Product.tolist():
    if val in product_group['Product '].tolist():
        types.append(product_group[product_group['Product '] == val]['Type'].tolist()[0])
        EOS.append(product_group[product_group['Product '] == val]['EOS'].tolist()[0])
    else:
        types.append('')
        EOS.append('') 
result['types'] = types
result['EOS'] = EOS 

prod_type = product_group[product_group['Product '] == Input_product]['Type'].tolist()[0]
result = result[result['types'] != prod_type]

result = result[result['EOS'] != 'Yes'] 

print(f"\nThe top Recommended Products for Cross selling to {Input_product}")
# print('Result: ',result['Product'].tolist()[:10])

result["Rank"] = result[['%Count']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)

print(result.sort_values("Rank").head(5))
result = result.sort_values("Rank")
print(result.head(5))

result.to_csv("2_CrossSell_prod_output.csv", index = False)