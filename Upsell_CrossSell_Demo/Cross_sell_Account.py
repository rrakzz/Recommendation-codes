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

print(f"----------------Upsell and Cross sell---------------")
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
        print(f'Upsell options for product: {prods}')
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        pro_1 = product_group[product_group['Type'] == prod_type]
        # pro_2 = pro_1[pro_1['Product '] == Input_upsell['Input Values'][1]]
        Rank_Target = product_group[product_group['Product '] == prods]['Upsell Rank'].tolist()[0]
        Capacity = product_group[product_group['Product '] == prods]['Capacity_Gbps'].tolist()[0]
           
        Upsell_products = pro_1['Product '][pro_1['Upsell Rank'] > Rank_Target].tolist()
#         print(f"1. Clustering Higher Variant Products than {prods}")   #(prods, ":", Upsell_products)
        
        if len(Upsell_products) > 0:
            for prod in acc_prods:
                if prod in Upsell_products:
                    Upsell_products.remove(prod)
#             print(prods, ":", Upsell_products)
#             print(f"2. Removing duplicates and clustering the Unique products")
            
            Upsell_products = pro_1['Product '][(pro_1['Upsell Rank'] > Rank_Target) & (pro_1['EOS'] != 'Yes')].tolist()
#             print(prods, ":", Upsell_products)
#             print(f"3. Clustering the Products with EOS as 'NO'")
            
            Upsell_products = pro_1['Product '][(pro_1['Upsell Rank'] > Rank_Target) & 
                                                (pro_1['EOS'] != 'Yes') & 
                                                (pro_1['Capacity_Gbps'] > Capacity)].tolist()
#             print(prods, "Upsell:", Upsell_products)
#             print(f"4. Clustering the Products with higher Capaciy than {prods}")
            print("Result: ", Upsell_products)
        else:
            print("Result: No producs to Upsell")
        
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
    
    print(f"\nThe top Recommended Products for Cross selling to {Input_account}")
    print('Result: ', df1['products'].tolist()[:10])
    
    print(f'\nChecking for Cross sell options based on Similar Accounts..')
        
    product_list = []
    for acc in Accounts_list:
        for i,val in enumerate(Data['Account Name: Account Name']):
            if acc == val:
                product_list.append(Data['Product'][i])

    prods_li = []
    counts_li = []
    for key,val in dict(Counter(product_list)).items():
        if key not in acc_prods:
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

    print(f"\nThe top Recommended Products for Cross selling to {Input_account}")
    print('Result: ',cross_sells1['products'].tolist()[:10])
            
else:
    print("No existing products found.")