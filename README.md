#Search Pattern

pattern_partial = '|'.join(r'(?<!\w)' + re.escape(keyword) + r'(?!\w)' for keyword in keywords_full)

print(f"Search pattern created with {len(keywords_full)} keywords")
df_filtered = df[df[['TI', 'AB', 'DE']].apply(lambda x: x.str.contains(pattern_partial, case=False, na=False)).any(axis=1)]
print(f"Filtered DataFrame: {df_filtered.shape[0]} rows match the keywords")
