import pandas as pd
import numpy as np
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize 
from nltk.stem.porter import PorterStemmer
stemmer = PorterStemmer()

Quiz_output_limit = 3
Similarity_output_limit = 10

print('TAKE A QUIZ...')
Mapping = pd.read_excel("Quiz_questions_template.xlsx", sheet_name = "Ponds_mapping")
Description = pd.read_excel("Quiz_questions_template.xlsx", sheet_name = "Ponds_desc")
Quiz = pd.read_excel("Quiz_questions_template.xlsx", sheet_name = "Ponds_quiz")

selection_list = Quiz.Slection.tolist()

option_list = []
for i, val in enumerate(Quiz.Question):
    options = Quiz.iloc[i].values.tolist()[2:]
    options_clean = [x for x in options if str(x) != 'nan']
    print(i+1, val, options_clean)
    
    if selection_list[i] != 'Multi':
        option_selected = input("Enter your option here: ")
        while option_selected.strip() not in options_clean:
            print("Please type from the list of options above.")
            option_selected = input("Type your option here: ")

        print("")
        if option_selected.strip() != "Skip" and option_selected.strip() != "Quit":
            option_list.append(option_selected.strip())
        if option_selected.strip() == "Quit":
            break
    else:
        option_selected_list = input("Enter your multiple options here (',' seperated): ")
        option_selected = option_selected_list.split(",")
        option_selected = [i.strip() for i in option_selected]
        while (set(option_selected).issubset(set(options_clean))) is False:
            print(f"Please type options from the above list with ',' seperated.")
            option_selected_list = input("Enter your multiple options here (',' seperated): ")
            option_selected = option_selected_list.split(",")
            
        print("")
        for opt in option_selected:
            opt = opt.strip()
            if opt != "Skip" and opt != "Quit":
                option_list.append(opt)
            if opt == "Quit":
                break
        if opt == "Quit":
            break        
print("Selected attributes: ", option_list)

Mapping_score = Mapping[option_list]
Mapping['Score'] = Mapping_score.sum(axis=1) 
Mapping = Mapping.sort_values(by = 'Score',ascending = False).reset_index().drop('index', axis = 1)
print("\nTop Products based on Quiz:")
print(Mapping[['Name', 'Score']].head(Quiz_output_limit))
Mapping[['Name', 'Score']].head(Quiz_output_limit).to_csv("products_quiz_output.csv")

Quiz_output = Mapping['Name'].head(Quiz_output_limit).tolist()
Description['merged'] = Description['Description'].astype(str) + " " + Description['Ingredients'].astype(str) + " " + Description['Work'].astype(str) 

Description_name = Description['Name'].tolist()
Description_desc = Description['merged'].tolist()
Description_clean = []
for i in Description_desc:
    Description_clean.append(i.replace('\n', ' ').replace('\r', ' ').replace('\\', ' ').replace('nan', ' ').replace('/', ' '))

Quiz_output_desc = []
for i in Quiz_output:
    if i in Description_name:
        Quiz_output_desc.append(Description_clean[Description_name.index(i)])

similar_products = pd.DataFrame()
for ii,X in enumerate(Quiz_output_desc):
    similarity_index = []
    for j,Y in enumerate(Description_clean):
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
        
    simil_pred = pd.DataFrame({'Name': Description_name, 'SI' : similarity_index})
    simil_pred = simil_pred.sort_values(by = 'SI',ascending = False)
    simil_pred = simil_pred.reset_index(inplace = False).drop('index', axis = 1) 
    simil_pred = simil_pred.drop([0])
    similar_products = similar_products.append(simil_pred.iloc[:10])
#     print(f'\nTop {Similarity_output_limit} Similar products of "{Quiz_output[ii]}"')
#     print(simil_pred.head(Similarity_output_limit))
    
similar_products = similar_products.sort_values(by = 'SI',ascending = False)
similar_products = similar_products.drop_duplicates(['Name'], keep='first')
similar_products =  similar_products.reset_index().drop('index', axis = 1)

print("\nTop Products based on Similarity:")
print(similar_products.head(Similarity_output_limit))
similar_products.head(Similarity_output_limit).to_csv("products_similarity_output.csv")