import pandas as pd
import re
import os

# Create a folder to save the results
result_folder = 'Result/Istatistic'
os.makedirs(result_folder, exist_ok=True)  # Create the folder if it does not exist
print(f"Result folder created at: {result_folder}")

# Read the keywords from a single-column text file
keywords_file_path = 'Terms/Surveying_Methods.txt'
with open(keywords_file_path, 'r', encoding='utf-8') as f:
    keywords_full = [line.strip().lower() for line in f if line.strip()]  # Read non-empty lines and add them to the keywords list
keywords_full = list(set(keywords_full))  # Make the keywords unique to remove duplicates
print(f"Keywords loaded: {len(keywords_full)} keywords read from {keywords_file_path}")

# Read the keywords to be excluded from a single-column text file
exclude_file_path = 'Terms/Surveying_Methods_Exclude.txt'
with open(exclude_file_path, 'r', encoding='utf-8') as f:
    exclude_keywords = [line.strip().lower() for line in f if line.strip()]  # Read non-empty lines and add them to the exclude list
print(f"Exclude keywords loaded: {len(exclude_keywords)} keywords read from {exclude_file_path}")

# Load the Excel file
file_path = 'Data/WOS+SCP_Raw.xlsx'
xls = pd.ExcelFile(file_path)  # Load the Excel file
print(f"Excel file loaded: {file_path}")

# Load the first sheet
df = pd.read_excel(xls, sheet_name='Data')  # Load the first sheet into a DataFrame
print(f"Sheet loaded: {df.shape[0]} rows and {df.shape[1]} columns")

# Filtering based on keywords (search in Title, Abstract, and Keywords columns)
pattern_partial = '|'.join(r'(?<!\w)' + re.escape(keyword) + r'(?!\w)' for keyword in keywords_full)  # Create partial match pattern for keywords

print(f"Search pattern created with {len(keywords_full)} keywords")
df_filtered = df[df[['TI', 'AB', 'DE']].apply(lambda x: x.str.contains(pattern_partial, case=False, na=False)).any(axis=1)]  # Partial match search for TI, AB, and DE
print(f"Filtered DataFrame: {df_filtered.shape[0]} rows match the keywords")

# Determine match types and collect statistics
match_stats = {keyword: {'TI': 0, 'AB': 0, 'DE': 0} for keyword in keywords_full}  # Create dictionary for keyword statistics
match_details = []
for idx, row in df_filtered.iterrows():  # Iterate over filtered rows
    for column in ['TI', 'AB', 'DE']:  # Check TI, AB, and DE columns
        if pd.notna(row[column]):  # If the column value is not empty
            for keyword in keywords_full:
                if re.search(r'\b' + re.escape(keyword) + r'\b', row[column], re.IGNORECASE):  # If there's an exact match
                    match_stats[keyword][column] += 1  # Increment the match count for the relevant column
                    highlighted_text = re.sub(f'({re.escape(keyword)})', r'**\1**', row[column], flags=re.IGNORECASE)  # Highlight the keyword
                    match_details.append({'Index': idx,'DI': row['DI'], 'Column': column,'Keyword': keyword, 'Highlighted Text': highlighted_text})  # Save match details
print(f"Match details collected for {len(match_details)} entries")

# Remove duplicates based on the DI column
df_filtered_unique = df_filtered.drop_duplicates(subset='DI')  # Remove duplicates based on the DI column
removed_duplicates = len(df_filtered) - len(df_filtered_unique)  # Calculate the number of removed duplicate rows
print(f"Removed {removed_duplicates} duplicate entries based on 'DI' column")

# Normalize the SC column and aggregate article counts by category
df_filtered_unique.loc[:, 'SC'] = df_filtered_unique['SC'].str.lower().str.strip()  # Normalize the SC column
category_counts = df_filtered_unique['SC'].value_counts()  # Get unique categories and their counts in the SC column

# Select all columns and save the filtered data to a new Excel file
output_file_path = os.path.join(result_folder, 'filtered_keywords_data_full.xlsx')
df_filtered_unique.to_excel(output_file_path, index=False)  # Save the filtered data to an Excel file
print(f"Filtered data saved to {output_file_path}")

# Save match details and statistics to Excel files
match_details_df = pd.DataFrame(match_details)  # Create a DataFrame for match details
match_details_output_path = os.path.join(result_folder, 'match_details.xlsx')
match_details_df.to_excel(match_details_output_path, index=False)  # Save match details to an Excel file
print(f"Match details saved to {match_details_output_path}")

# Save statistics to an Excel file
stats_output_path = os.path.join(result_folder, 'match_statistics.xlsx')
stats_df = pd.DataFrame.from_dict(match_stats, orient='index')  # Create a DataFrame for keyword statistics
stats_df.to_excel(stats_output_path)  # Save statistics to an Excel file
print(f"Match statistics saved to {stats_output_path}")

# Identify other unique keywords in the DE column excluding the exclude list
de_keywords = df_filtered_unique['DE'].dropna().str.split(';').explode().str.strip().str.lower()  # Split, clean, and normalize keywords in the DE column
de_keywords_filtered = de_keywords[~de_keywords.isin(exclude_keywords)].value_counts()  # Count keywords excluding the exclude list
print(f"Found {len(de_keywords_filtered)} unique keywords in DE column excluding exclude list")

de_keywords_df = de_keywords_filtered.reset_index()  # Create a DataFrame for keywords and their counts
de_keywords_df.columns = ['Keyword', 'Count']  # Set column names
de_keywords_output_path = os.path.join(result_folder, 'de_keywords_statistics.xlsx')
de_keywords_df.to_excel(de_keywords_output_path, index=False)  # Save DE column keywords to an Excel file
print(f"DE keywords statistics saved to {de_keywords_output_path}")

# Save the category counts in the SC column to an Excel file
category_counts_output_path = os.path.join(result_folder, 'category_counts_statistics.xlsx')
category_counts_df = category_counts.reset_index()
category_counts_df.columns = ['Category', 'Count']
category_counts_df.to_excel(category_counts_output_path, index=False)
print(f"Category counts saved to {category_counts_output_path}")

# Print summary of matches
print(f"Number of matches in Title: {df['TI'].str.contains(pattern_partial, case=False, na=False).sum()}")
print(f"Number of matches in Abstract: {df['AB'].str.contains(pattern_partial, case=False, na=False).sum()}")
print(f"Number of matches in Keywords: {df['DE'].str.contains(pattern_partial, case=False, na=False).sum()}")
print(f"Number of duplicate entries removed: {removed_duplicates}")

# Advanced statistics
print("\n--- Advanced Statistics ---")
print(f"Total records in original dataset: {df.shape[0]}")
print(f"Total records after filtering: {df_filtered.shape[0]}")
print(f"Total unique records after removing duplicates: {df_filtered_unique.shape[0]}")
print(f"Percentage of records filtered: {df_filtered.shape[0] / df.shape[0] * 100:.2f}%")
print(f"Percentage of unique records after removing duplicates: {df_filtered_unique.shape[0] / df.shape[0] * 100:.2f}%")
print(f"Number of unique keywords found in DE column (excluding exclude list): {len(de_keywords_filtered)}")
print(f"Number of unique categories in SC column: {len(category_counts)}")
print(f"Categories and their counts (number of articles in each category):\n{category_counts}")
