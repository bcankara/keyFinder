# KeyFiltering.py

This Python script is designed for filtering and performing statistical analysis on scientific articles. It processes a given Excel file to filter text based on specified keywords and generates comprehensive statistical reports.

---

## Features
- **Keyword Filtering**:
  - Filters the *Title* (TI), *Abstract* (AB), and *Keywords* (DE) columns based on predefined keywords.
  - Allows specifying a list of keywords to exclude during the filtering process.

- **Data Cleaning and Categorization**:
  - Removes duplicate entries based on the `DI` column.
  - Normalizes and categorizes the `SC` column, reporting the number of articles in each unique category.

- **Detailed Match Analysis**:
  - Tracks which columns and rows match specific keywords.
  - Highlights matching keywords and saves detailed match data.

- **Statistical Reporting**:
  - Saves filtered data, match details, keyword statistics, and category counts in separate Excel files for further analysis.

- **Advanced Analytics**:
  - Calculates the percentage of filtered and unique records.
  - Provides insights into unique keywords in the `DE` column, excluding specified keywords.

---

## Input Files
1. **Terms/Surveying_Methods.txt**: Contains the list of keywords for filtering.
2. **Terms/Surveying_Methods_Exclude.txt**: Specifies keywords to exclude from the filtering process.
3. **Data/WOS+SCP_Raw.xlsx**: The Excel file containing the raw dataset for processing.

---

## Output Files
Generated outputs are saved in the `Result/Istatistic` folder:
1. **filtered_keywords_data_full.xlsx**: Filtered dataset after applying keyword filters.
2. **match_details.xlsx**: Detailed keyword matches, including highlighted text.
3. **match_statistics.xlsx**: Statistical breakdown of keyword matches in each column.
4. **de_keywords_statistics.xlsx**: Unique keywords from the `DE` column, excluding specified keywords.
5. **category_counts_statistics.xlsx**: Counts of articles in each unique category from the `SC` column.

---

## How to Use
1. Place the required input files in the specified directories.
2. Run the script: `python KeyFiltering.py`.
3. Check the `Result/Istatistic` folder for the output files.

---

## Dependencies
- Python 3.x
- Libraries: 
  - `pandas`
  - `re`
  - `os`

Install dependencies via pip:
```bash
pip install pandas
