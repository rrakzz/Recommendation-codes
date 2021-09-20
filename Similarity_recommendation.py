import pandas as pd
import numpy as np
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize 
from nltk.stem.porter import PorterStemmer
stemmer = PorterStemmer()

open_ticket_count = 10
extract_similar_tickets = 5

sophos = pd.read_excel("Case_Tickets.xlsx")

#Merging Description and Subject for similarity check
sophos['merged'] = sophos['Description'].astype(str) + " " + sophos['Subject'].astype(str)
print("Total tickets found:", len(sophos))

#Selection of Closed tickets into DataFrame
sophos_open = sophos[sophos['Status'] == 'Closed']
sophos_open = sophos_open.drop_duplicates(['Case Number'], keep='last')
print("Total unique open tickets found:", len(sophos_open))

#Selection of Open tickets into DataFrame
sophos_close = sophos[sophos['Status'] != 'Closed']
sophos_close = sophos_close.drop_duplicates(['Case Number'], keep='last')
print("Total unique closed tickets found:", len(sophos_close))

#Cleaning open tickets
sophos_desc_o = sophos_open['merged'].tolist()
sophos_desc_o_clean = []
for i in sophos_desc_o:
    sophos_desc_o_clean.append(i.replace('\n', ' ').replace('\r', ' ').replace('\\', ' ').replace('nan', ' '))

Case_Number_o = sophos_open['Case Number'].tolist()
Status_o = sophos_open['Status'].tolist()
Sbj_o = sophos_open['Subject'].tolist()
Product_Category_o = sophos_open['Product Category'].tolist()

#Looping through each open ticket
similar_tickets = pd.DataFrame()
if len(sophos_open) > 1:
    for ii,X in enumerate(sophos_desc_o_clean[:open_ticket_count]):
        print(f'Checking #{ii+1}/{len(sophos_open)}:', Case_Number_o[ii])
        try:        
            #Filtering based on Product_Category
            if type(Product_Category_o[ii]) is str:
                sophos_c = sophos_close[sophos_close['Product Category'] == Product_Category_o[ii]]  
            else:
                sophos_c = sophos_close
            #print('')
            #print("Product Category:", Product_Category_o[ii])
            #print("Subject:", Sbj_o[ii])
            #print('Number of tickets with matching product category:', len(sophos_c))
                                
            Case_Number_c = sophos_c['Case Number'].tolist()
            sophos_desc_c = sophos_c['merged'].tolist()
            Sbj_c = sophos_c['Subject'].tolist()
            Date_closed_c = sophos_c['Case Last Modified Date'].tolist()
            Case_Owner_c = sophos_c['Case Owner'].tolist()
            sophos_desc_c_clean = []
            for i in sophos_desc_c:
                sophos_desc_c_clean.append(i.replace('\n', ' ').replace('\r', ' ').replace('\\', ' ').replace('nan', ' '))

            #Looping through each closed ticket
            similarity_index = []
            for j,Y in enumerate(sophos_desc_c_clean):
                #Removing symbols
                symbols = "!\"#$%&()*+-./:;<=>?@[\]^_`{|}~\n"
                for i in symbols:
                    Y = Y.replace(i, ' ')
                    X = X.replace(i, ' ')

                X_list = word_tokenize(X.lower())  
                Y_list = word_tokenize(Y.lower()) 

                # sw contains the list of stopwords 
                sw = stopwords.words('english')  
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

            simil_pred = pd.DataFrame({'Sl.no.':ii, 'Open_case': Case_Number_o[ii], 'Open_Ticket_subject': Sbj_o[ii], 
                                       "Product Category": Product_Category_o[ii], 'Similar_Cases':Case_Number_c, 
                                       'Closed_Ticket_Subject': Sbj_c, 'similarity_index' : similarity_index,
                                       'Date_closed': Date_closed_c, 'Case_Owner':Case_Owner_c})
            simil_pred = simil_pred.sort_values(by = 'similarity_index',ascending = False)
            simil_pred = simil_pred.reset_index(inplace = False) 
#             print(simil_pred.head(5))
            similar_tickets = similar_tickets.append(simil_pred.iloc[:extract_similar_tickets])
        except:
            print("Sorry, no matching closed tickets found.")
else:
    print("Sorry, no open tickets found.")
    
similar_tickets =  similar_tickets.drop('index', axis = 1).reset_index().drop('index', axis = 1)
similar_tickets.to_csv('Similar_tickets_output.csv', index = False)