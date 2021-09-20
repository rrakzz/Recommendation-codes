import pandas as pd
import numpy as np
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize 
from nltk.stem.porter import PorterStemmer
stemmer = PorterStemmer()
#from translate import Translator
#translator= Translator(from_lang="spanish", to_lang="English")
from textblob import TextBlob


print("Generating Item Similarity Matrix...")
Description = pd.read_excel("Quiz_questions_template_Mexico.xlsx", sheet_name = "Item_description_Ponds")
Description['merged'] = Description['Description'].astype(str) + " " + Description['Ingredients'].astype(str) + " " + Description['Work'].astype(str) #+ " " + Description['Tutorial'].astype(str) + Description['Title_1'].astype(str) + ", " + Description['Title_2'].astype(str) + ", " + Description['Title_3'].astype(str) + ", " + Description['Title_4'].astype(str)  + ", " + Description['Title_5'].astype(str) 

Description_name = Description['Name'].tolist()
Description_desc = Description['merged'].tolist()
Description_clean = []
for i in Description_desc:
    Description_clean.append(i.replace('\n', ' ').replace('\r', ' ').replace('\\', ' ').replace('nan', ' ').replace('/', ' '))

similar_products = pd.DataFrame({'Name': Description_name})
for ii,X in enumerate(Description_clean):
    similarity_index = []
    for j,Y in enumerate(Description_clean):
    
        #Translating from Spanish to English
        #X = translator.translate(X)
        #Y = translator.translate(Y)
        #OR using TextBlob
        X = TextBlob(X).translate(to='en')
        Y = TextBlob(Y).translate(to='en')
        
        #Removing symbols
        symbols = "!\"#$%&()*+-./:;<=>?@[\]^_`{|}~\n"
        for i in symbols:
            Y = Y.replace(i, ' ')
            X = X.replace(i, ' ')

        X_list = word_tokenize(X.lower())  
        Y_list = word_tokenize(Y.lower()) 

        # sw contains the list of stopwords 
        sw = stopwords.words('english')   #spanish
        l1 =[];l2 =[] 

        # remove stop words from the string 
        X_set = {w for w in X_list if not w in sw}  
        Y_set = {w for w in Y_list if not w in sw} 

        #Stemming and remove words with length 1 and 2 letters
        X_set_new, Y_set_new = {''}, {''}
        for w in X_set:
            if len(w) > 2:
                X_set_new.add(stemmer.stem(w))
        for w in Y_set:
            if len(w) > 2:
                Y_set_new.add(stemmer.stem(w))    

        # form a set containing keywords of both strings  
        rvector = X_set_new.union(Y_set_new)  
        for w in rvector: 
            if w in X_set_new: l1.append(1) # create a vector 
            else: l1.append(0) 
            if w in Y_set_new: l2.append(1) 
            else: l2.append(0) 
        c = 0

        # cosine formula  
        for i in range(len(rvector)): 
                c+= l1[i]*l2[i] 
        cosine = c / float((sum(l1)*sum(l2))**0.5) 
        similarity_index.append(round(cosine, 2))
    similar_products[f'{Description_name[ii]}'] = similarity_index
similar_products.to_excel("Similarity_matrix_output_Ponds_Mexico_TB.xlsx", index=False)
print(similar_products)