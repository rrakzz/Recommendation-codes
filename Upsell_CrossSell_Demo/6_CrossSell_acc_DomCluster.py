import pandas as pd
import numpy as np
from sklearn.preprocessing import MinMaxScaler
from collections import Counter

Input_account = 'Blue Cross and Blue Shield of Louisiana'

Data = pd.read_excel("Case Data Dump from Jan'21 to Jul'21.xlsx")
product_group = pd.read_excel("Upsell-Data.xlsx")
Accounts = pd.read_excel("Account Details-Palo Alto.xlsx")

Account_cliserID = Accounts[Accounts['Account Name'] == Input_account]['ClusterID'].tolist()[0]
Accounts_list = Accounts[Accounts['ClusterID'] == Account_cliserID]['Account Name'].tolist()

acc_prods = list(set(Data[Data['Account Name: Account Name'] == Input_account]['Product'].tolist()))
Cross_sell_p = []
Cross_sell_c = []

print(f"Account: {Input_account}")
print(f"Identified {len(acc_prods)} unique products.")
print(acc_prods)
print('')

def Cross_sell(Data,product_):
    table = pd.pivot_table(Data, values ='Case Number', index =['Account Name: Account Name'], columns =['Product'], aggfunc = "count")
    df1 = pd.DataFrame(table.to_records())
    df2 = df1[df1[f'{product_}'].notnull()]
#     print(f"Number of Accounts with {product_}: {len(df2)}")

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
    return result['Product'].tolist()[1:],result['%Count'].tolist()[1:]
    
if len(acc_prods) > 0:
    for prods in acc_prods:
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        pro_1 = product_group[product_group['Type'] == prod_type]
        # pro_2 = pro_1[pro_1['Product '] == Input_upsell['Input Values'][1]]
        Rank_Target = product_group[product_group['Product '] == prods]['Upsell Rank'].tolist()[0]
        Capacity = product_group[product_group['Product '] == prods]['Capacity_Gbps'].tolist()[0]
        
        Cross_sel_p, Cross_sel_c = Cross_sell(Data,prods)
        Cross_sell_p.append(Cross_sel_p)
        Cross_sell_c.append(Cross_sel_c)
        print('')
    
    print(f'\nChecking for Cross sell options based on Similar Accounts..')
        
    product_list = []
    for acc in Accounts_list:
        for i,val in enumerate(Data['Account Name: Account Name']):
            if acc == val:
                product_list.append(Data['Product'][i])

    prods_li = []
    counts_li = []
    for key,val in dict(Counter(product_list)).items():
        if key not in acc_prods and type(key) == str:
            prods_li.append(key)
            counts_li.append(val)

    cross_sells1 = pd.DataFrame({'products':prods_li , 'count': counts_li})
    cross_sells1 = cross_sells1.sort_values(f'count', ascending=False)

    types, EOS = [], []
    for val in cross_sells1.products.tolist():
        if val in product_group['Product '].tolist():
            types.append(product_group[product_group['Product '] == val]['Type'].tolist()[0])
            EOS.append(product_group[product_group['Product '] == val]['EOS'].tolist()[0])
        else:
            types.append('')
            EOS.append('') 
    cross_sells1['types'] = types
    cross_sells1['EOS'] = EOS 


    for prods in acc_prods:
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        cross_sells1 = cross_sells1[cross_sells1['types'] != prod_type]

    cross_sells1 = cross_sells1[cross_sells1['EOS'] != 'Yes'] 
    
    cross_sells1["Rank"] = cross_sells1[['count']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)

    cross_sells1 = cross_sells1.sort_values("Rank")
    print(cross_sells1.head(5))

    cross_sells1.to_csv("6_CrossSell_acc_DomCluster_output.csv", index = False)