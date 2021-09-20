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

#     print(f"Number of unique Accounts: {len(df1)}")
    #print(Input_upsell.columns.tolist())

#     product_ = Input_upsell["Input Values"][1]

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

#     print(f"\nThe top 10 Recommended Products for Cross selling to Accounts with {product_}")
#     print(result.head(11))
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
    
    print(f'Checking for Cross sell options based on Accounts with Similar products..')
    
    cross_sell_prods = []
    cross_sell_count = []

    for x,v in enumerate(Cross_sell_p):
        for i,val in enumerate(v):
            if val not in acc_prods:
                cross_sell_prods.append(val)
                cross_sell_count.append(Cross_sell_c[x][i])

    cross_sells = pd.DataFrame({'products':cross_sell_prods , '%count': cross_sell_count})
    
    table = pd.pivot_table(cross_sells, values ='%count', index =['products'], aggfunc = np.mean) #, columns =['Product']
    df1 = pd.DataFrame(table.to_records())
    df1 = df1.sort_values(f'%count', ascending=False)

    Ranks, types, capacities, EOS = [], [], [], []
    for val in df1.products.tolist():
        if val in product_group['Product '].tolist():
            Ranks.append(product_group[product_group['Product '] == val]['Upsell Rank'].tolist()[0])
            types.append(product_group[product_group['Product '] == val]['Type'].tolist()[0])
            capacities.append(product_group[product_group['Product '] == val]['Capacity_Gbps'].tolist()[0])
            EOS.append(product_group[product_group['Product '] == val]['EOS'].tolist()[0])
        else:
            Ranks.append('')
            types.append('')
            capacities.append('')
            EOS.append('') 
    df1['Ranks'] = Ranks
    df1['types'] = types
    df1['capacities'] = capacities  
    df1['EOS'] = EOS 

    for prods in acc_prods:
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        df1 = df1[df1['types'] != prod_type]
            
    df1 = df1[df1['EOS'] != 'Yes'] 

    df1 = df1.drop(['Ranks'], axis = 1)
    df1["Rank"] = df1[['%count']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)

    df1 = df1.sort_values("Rank")
    print(df1.head(5))

    df1.to_csv("4_CrossSell_acc_ProCluster_output.csv", index = False)