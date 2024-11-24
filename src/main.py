import pandas as pd

# File Paths
input_file = r'spreadsheets\Plant_Data_BaseMasterfinal(1)REVISED3plant description.xlsx'
output_file = r'spreadsheets\New_Spreadsheet.xlsx'

# Load client spreadsheet
client_df = pd.read_excel(input_file)

# Strip leading/trailing spaces from column names
client_df.columns = client_df.columns.str.strip()

# Map "X" to TRUE and blanks to FALSE for boolean columns
def map_boolean(cell):
    return 'TRUE' if cell == 'X' else 'FALSE'

# Apply boolean mapping to specific columns
for col in ['Evergreen', 'Deciduous', 'Perennial', 'Deer']:
    client_df[col] = client_df[col].apply(map_boolean)

# Combine style columns into one
def map_styles(row):
    styles = []
    if row.get('English') == 'X':
        styles.append('english')
    if row.get('Formal') == 'X':
        styles.append('formal')
    if row.get('Tropical') == 'X':
        styles.append('tropical')
    if row.get('Waterwise') == 'X':
        styles.append('waterwise')
    return ', '.join(styles)

client_df['Style'] = client_df.apply(map_styles, axis=1)

# Combine sunlight columns into one
def map_sunlight(row):
    sunlight = []
    if row.get('Full-sun') == 'X':
        sunlight.append('full_sun')
    if row.get('Full-shade') == 'X':
        sunlight.append('full_shade')
    if row.get('Filter-sun') == 'X':
        sunlight.append('filter_sun')
    return ', '.join(sunlight)

client_df['Sunlight'] = client_df.apply(map_sunlight, axis=1)

# Remove original style and sunlight columns
columns_to_remove = ['English', 'Formal', 'Tropical', 'Waterwise', 'Full-sun', 'Full-shade', 'Filter-sun']
client_df = client_df.drop(columns=columns_to_remove, errors='ignore')

# Create the ID column
category_letter_to_id_start = {chr(i): 1000 * (i - 64) for i in range(65, 77)}  # Map A-L to 1000, 2000, ..., 12000

def assign_ids(df):
    ids = []
    grouped = df.groupby('Category')
    for category, group in grouped:
        base_id = category_letter_to_id_start.get(category.upper(), 1000)  # Default to 1000 if no match
        ids.extend(range(base_id, base_id + len(group)))
    return ids

client_df['ID'] = assign_ids(client_df)

# Ensure proper column order
final_columns = ['ID', 'Category', 'Common', 'Botanical', 'Height', 'Width', 'Style', 'Sunlight', 'Zone', 
                 'Evergreen', 'Deciduous', 'Perennial', 'Deer', 'Notes', 'ImageURL']

# Add missing columns (if any) with empty values and reorder
for col in final_columns:
    if col not in client_df.columns:
        client_df[col] = ''
client_df = client_df[final_columns]

# Save the formatted spreadsheet
client_df.to_excel(output_file, index=False)

print(f"Formatted data saved to {output_file}")