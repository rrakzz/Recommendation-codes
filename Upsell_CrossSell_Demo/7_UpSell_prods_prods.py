import pandas as pd
import numpy as np
from sklearn.preprocessing import MinMaxScaler
from collections import Counter

Input_products = ['Panorama', 'PA-5060'] # 1 product or many

Data = pd.read_excel("Case Data Dump from Jan'21 to Jul'21.xlsx")
product_group = pd.read_excel("Upsell-Data.xlsx")

if len(Input_products) > 0:
    for prods in Input_products:
        print(f'Upsell options for product: {prods}')
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        pro_1 = product_group[product_group['Type'] == prod_type]
        # pro_2 = pro_1[pro_1['Product '] == Input_upsell['Input Values'][1]]
        Rank_Target = product_group[product_group['Product '] == prods]['Rank'].tolist()[0]
        Capacity = product_group[product_group['Product '] == prods]['Capacity_Gbps'].tolist()[0]
           
    print(f'\nChecking for Upsell options..')

    product_list = Data['Product'].tolist()
    
    prods_li = []
    counts_li = []
    for key,val in dict(Counter(product_list)).items():
        if key not in Input_products:
            prods_li.append(key)
            counts_li.append(val)

    Up_sells = pd.DataFrame({'products':prods_li , 'count': counts_li})
    Up_sells = Up_sells.sort_values(f'count', ascending=False)

    types, EOS = [], []
    for val in Up_sells.products.tolist():
        if val in product_group['Product '].tolist():
            types.append(product_group[product_group['Product '] == val]['Type'].tolist()[0])
            EOS.append(product_group[product_group['Product '] == val]['EOS'].tolist()[0])
        else:
            types.append('')
            EOS.append('') 
    Up_sells['types'] = types
    Up_sells['EOS'] = EOS 


    for prods in Input_products:
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        Up_sells1 = Up_sells[Up_sells['types'] == prod_type]

    Up_sells1 = Up_sells1[Up_sells1['EOS'] != 'Yes'] 

    Up_sells1["Rank"] = Up_sells1[['count']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)
    
    Up_sells1 = Up_sells1.sort_values("Rank")
    print(Up_sells1.head(5))

    Up_sells1.to_csv("7_UpSell_prods_prods_output.csv", index = False)