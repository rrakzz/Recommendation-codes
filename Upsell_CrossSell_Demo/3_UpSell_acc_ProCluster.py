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

products_up = pd.DataFrame()

if len(acc_prods) > 0:
    for prods in acc_prods:
        print(f'Upsell options for product: {prods}')
        prod_type = product_group[product_group['Product '] == prods]['Type'].tolist()[0]
        pro_1 = product_group[product_group['Type'] == prod_type]
        # pro_2 = pro_1[pro_1['Product '] == Input_upsell['Input Values'][1]]
        Rank_Target = product_group[product_group['Product '] == prods]['Rank'].tolist()[0]
        Capacity = product_group[product_group['Product '] == prods]['Capacity_Gbps'].tolist()[0]
           
        Upsell_products = pro_1['Product '][pro_1['Rank'] > Rank_Target].tolist()
#         print(f"1. Clustering Higher Variant Products than {prods}")   #(prods, ":", Upsell_products)
        
        if len(Upsell_products) > 0:
            for prod in acc_prods:
                if prod in Upsell_products:
                    Upsell_products.remove(prod)
#             print(prods, ":", Upsell_products)
#             print(f"2. Removing duplicates and clustering the Unique products")
            
            Upsell_products = pro_1['Product '][(pro_1['Rank'] > Rank_Target) & (pro_1['EOS'] != 'Yes')].tolist()
#             print(prods, ":", Upsell_products)
#             print(f"3. Clustering the Products with EOS as 'NO'")
            
            Upsell_products = pro_1['Product '][(pro_1['Rank'] > Rank_Target) & 
                                                (pro_1['EOS'] != 'Yes') & 
                                                (pro_1['Capacity_Gbps'] > Capacity)].tolist()
            pro_2 = pro_1[(pro_1['Rank'] > Rank_Target) & 
                                        (pro_1['EOS'] != 'Yes') & 
                                        (pro_1['Capacity_Gbps'] > Capacity)]
#             print(prods, "Upsell:", Upsell_products)
#             print(f"4. Clustering the Products with higher Capaciy than {prods}")
            pro_2['Target'] = prods
            print("Result: ", Upsell_products)
#             print(pro_2)
            print('')
            products_up = products_up.append(pro_2)
        else:
            print("Result: No producs to Upsell")
            print('')
            
    products_up = products_up.drop(['Upsell Rank' , 'Capacity1', 'Rank', 'Up-sell'], axis = 1)
    products_up["Rank"] = products_up[['Capacity_Gbps']].apply(tuple,axis=1).rank(method='dense',ascending=False).astype(int)

    products_up = products_up.sort_values("Rank")
    print(products_up.head(5))

    products_up.to_csv("3_UpSell_acc_ProCluster_output.csv", index = False)