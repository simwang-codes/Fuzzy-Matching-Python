# Fuzzy Matching Script: Matching names from two Excel files using fuzzy matching techniques. Great for B2B or marketing research tasks

from thefuzz import fuzz, process
import pandas as pd

# Load data from Excel files
a_list = pd.read_excel("A_list.xlsx", usecols=[0], header=None)
b_list = pd.read_excel("B_list.xlsx", usecols=[0], header=None)

# Normalize text data in both files for more consistent matching
def preprocess_data(column):
    return column.str.lower()

a_list[0] = preprocess_data(a_list[0])
b_list[0] = preprocess_data(b_list[0])


matches = []

# Skipping the first row of a_list since it normally stores the title
for _, name in a_list.iloc[1:].iterrows():
    # This script uses "fuzz.token_sort_ratio" as its scorer because it helps detect substring matches, such as "Target" in "Target Corp"
    match = process.extractOne(name[0], b_list.iloc[1:, 0], scorer=fuzz.token_sort_ratio)
    
# If a match is found, append the matched information (A_list name, B_list name, and score) to the matches list
    if match:
        matches.append({
            "A_list_name": name[0],
            "B_list_name": match[0],
            "Score": match[1]
        })
    
results = pd.DataFrame(matches)

# Filter results by score threshold
threshold = 50
results = results[results["Score"] >= threshold]

# Store the results in a new file
results.to_excel("fuzzy_matching_results.xlsx", index=False)
print("Matching results saved to fuzzy_matching_results.xlsx")