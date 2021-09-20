import pandas as pd
import numpy as np

sophos = pd.read_excel("Case Dump_6th oct.xlsx")
sophos['merged'] = sophos['Product Category'].astype(str) + " " + sophos['Description'].astype(str) + " " + sophos['Subject'].astype(str) + " " + sophos['Case Reason'].astype(str) + " " + sophos['Skill'].astype(str)
sophos_desc = sophos['merged'].tolist()
sophos_desc_clean = []
for i in sophos_desc:
    sophos_desc_clean.append(i.replace('\n', ' ').replace('\r', ' ').replace('\\', ' ').replace('nan', ' '))

Case_Number_list = sophos['Case Number'].tolist()
product_category_list = sophos['Product Category'].tolist()
Subject_list = sophos['Subject'].tolist()
Description = sophos['Description'].astype(str).tolist()

try:
    input_product = input("Enter Case number: ")
    for i, val in enumerate(Case_Number_list):
        if int(input_product) == int(val):
            print("Retrieving similar tickets, please wait...")
            product_index = i
            break

    from nltk.corpus import stopwords 
    from nltk.tokenize import word_tokenize 

    X = sophos_desc_clean[product_index]

    similarity_index = []
    for sent in sophos_desc_clean:
        X_list = word_tokenize(X.lower())  
        Y_list = word_tokenize(sent.lower()) 

        # sw contains the list of stopwords 
        sw = stopwords.words('english')  
        l1 =[];l2 =[] 

        # remove stop words from the string 
        X_set = {w for w in X_list if not w in sw}  
        Y_set = {w for w in Y_list if not w in sw} 

        # form a set containing keywords of both strings  
        rvector = X_set.union(Y_set)  
        for w in rvector: 
            if w in X_set: l1.append(1) # create a vector 
            else: l1.append(0) 
            if w in Y_set: l2.append(1) 
            else: l2.append(0) 
        c = 0

        # cosine formula  
        for i in range(len(rvector)): 
                c+= l1[i]*l2[i] 
        cosine = c / float((sum(l1)*sum(l2))**0.5) 
        similarity_index.append(round(cosine, 2))

    simil_pred = pd.DataFrame({'Case_Number':Case_Number_list, 'product_category': product_category_list, 
                               'Subject': Subject_list, 'Description': Description,
                               'similarity_index' : similarity_index})
    simil_pred = simil_pred.sort_values(by = 'similarity_index',ascending = False)
    index_names = simil_pred[simil_pred['Case_Number'] == int(input_product) ].index 
    simil_pred = simil_pred.drop(index_names)
    print(simil_pred[['Case_Number', 'similarity_index']].head(5))

except:
    print("Sorry, Case number not found.")