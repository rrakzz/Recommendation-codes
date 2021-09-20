import pandas as pd
import numpy as np

Quiz_output_limit = 3
Similarity_output_limit = 5

print('TAKE A QUIZ...')
Mapping = pd.read_excel("Quiz_questions_template.xlsx", sheet_name = "ATH_mapping")
Description = pd.read_excel("Quiz_questions_template.xlsx", sheet_name = "ATH_desc")
Quiz = pd.read_excel("Quiz_questions_template.xlsx", sheet_name = "ATH_quiz")
Matrix = pd.read_excel("Similarity_matrix_output_ATH.xlsx")

selection_list = Quiz.Type.tolist()

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
            option_selected = [i.strip() for i in option_selected]
            
        print("")
        for opt in option_selected:
            opt = opt.strip()
            if opt != "Skip" and opt != "Quit":
                option_list.append(opt)
            if opt == "Quit":
                break
        if opt == "Quit":
            break        

print("-"*60)
print("Selected attributes: ", option_list)
print("-"*60)
Mapping_score = Mapping[option_list]
Mapping['Score'] = Mapping_score.sum(axis=1) 
Mapping = Mapping.sort_values(by = 'Score',ascending = False).reset_index().drop('index', axis = 1)
print("\nRecommendation based on Quiz:")
print(Mapping[['Name', 'Score']].head(Quiz_output_limit))
Mapping[['Name', 'Score']].head(Quiz_output_limit).to_csv("products_quiz_output.csv")
print("-"*60)
for i in Mapping['Name'].head(Quiz_output_limit).tolist():
    print('\nRecommendation based on Similarity:')
    print(Matrix[['Name', i]].sort_values(by = i, ascending = False).reset_index(inplace = False).drop('index', axis = 1).drop([0]).head(Similarity_output_limit))
    
    Description['Products_merged'] = Description['Title_1'].astype(str) + ", " + Description['Title_2'].astype(str) + ", " + Description['Title_3'].astype(str) + ", " + Description['Title_4'].astype(str)  + ", " + Description['Title_5'].astype(str)  
Description_products = Description['Products_merged'].tolist()
Description_tutorials = Description['Tutorial'].tolist()
Description_name = Description['Name'].tolist()

print("-"*60)
print('\nTutorials and Products Recommendation:')

Quiz_output_products = []
Quiz_output_tutorials = []
for i in Mapping['Name'].head(Quiz_output_limit).tolist():
    if i in Description_name:
        Quiz_output_tutorials.append(Description_tutorials[Description_name.index(i)])
        Quiz_output_products.append(Description_products[Description_name.index(i)])

grouping = pd.DataFrame({"Name": Mapping['Name'].head(Quiz_output_limit).tolist(), "Tutorials": Quiz_output_tutorials,
                        "Products": Quiz_output_products})
print(grouping)