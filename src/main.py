import pandas as pd

# File Paths
source_file = r'spreadsheets\Plant_Data_BaseMasterfinal(1)REVISED2 plant description.xlsx'
output_file = r'spreadsheets\processed_Plant_Data_Base.xlsx'

# Load data from source file
source_df = pd.read_excel(source_file)

# Strip leading/trailing spaces from source column names
source_df.columns = source_df.columns.str.strip()

# Print columns for debugging
print("Source Columns:", source_df.columns.tolist())

# Check for and remove duplicate columns in source_df
source_df = source_df.loc[:, ~source_df.columns.duplicated()]

# Verify columns again
print("Cleaned Source Columns:", source_df.columns.tolist())

# Mapping source columns to target columns
column_mapping = {
    'Plant': 'Category',
    'Common Name': 'Common',
    'Botanical Name': 'Botanical',
    'Height': 'Height',
    'Width': 'Width',
    'Evergreen': 'Evergreen',
    'Deciduous': 'Deciduous',
    'Perennial': 'Perennial',
    'Deer': 'Deer',
    'Notes': 'Notes'
}

# Function to map "X" values to the corresponding style names
def map_styles(row):
    styles = []
    if row.get('English') == 'X':
        styles.append('english')
    if row.get('Formal') == 'X':
        styles.append('formal')
    if row.get('Tropical') == 'X':
        styles.append('tropical')
    if row.get('Xeriscape') == 'X':
        styles.append('xeriscape')
    return ', '.join(styles)

# Apply the function to create the 'Style' column
source_df['Style'] = source_df.apply(map_styles, axis=1)

# Function to map "X" values to the corresponding sunlight names
def map_sunlight(row):
    sunlight = []
    if row.get('Full-sun') == 'X':
        sunlight.append('full_sun')
    if row.get('Full-shade') == 'X':
        sunlight.append('full_shade')
    if row.get('Filter-sun') == 'X':
        sunlight.append('filter_shade')
    return ', '.join(sunlight)

# Apply the function to create the 'Sunlight' column
source_df['Sunlight'] = source_df.apply(map_sunlight, axis=1)

# Function to map "X" values to the corresponding zone names
def map_zone(row):
    zones = []
    if row.get('ZN') == 'X':
        zones.append('TX_ZN')
    if row.get('ZS') == 'X':
        zones.append('TX_ZS')
    if row.get('ZE') == 'X':
        zones.append('TX_ZE')
    if row.get('ZW') == 'X':
        zones.append('TX_ZW')
    if row.get('ZC') == 'X':
        zones.append('TX_ZC')
    return ', '.join(zones)

# Apply the function to create the 'Zone' column
source_df['Zone'] = source_df.apply(map_zone, axis=1)

# Function to map 'X' to TRUE and empty to FALSE for specific columns
def map_boolean(cell):
    return 'TRUE' if cell == 'X' else 'FALSE'

# Apply the map_boolean function to the relevant columns
source_df['Evergreen'] = source_df['Evergreen'].apply(map_boolean)
source_df['Deciduous'] = source_df['Deciduous'].apply(map_boolean)
source_df['Perennial'] = source_df['Perennial'].apply(map_boolean)
source_df['Deer'] = source_df['Deer'].apply(map_boolean)

# Add newly mapped columns to the mapping dictionary
column_mapping.update({
    'Style': 'Style',
    'Sunlight': 'Sunlight',
    'Zone': 'Zone',
    'Evergreen': 'Evergreen',
    'Deciduous': 'Deciduous',
    'Perennial': 'Perennial',
    'Deer': 'Deer'
})

# Rename the source columns to match the target columns
source_df = source_df.rename(columns=column_mapping)

# Reset the index for the dataframe and remove duplicates
source_df = source_df.reset_index(drop=True).drop_duplicates()

# Sort the dataframe by 'Category' in ascending order
source_df = source_df.sort_values(by='Category').reset_index(drop=True)

# Create a dictionary to map the first letter of the 'Category' to the starting ID
category_letter_to_id_start = {chr(i): 1000 * (i - 64) for i in range(65, 91)}

# Function to create the 'ID' column based on 'Category'
def assign_ids(df):
    ids = []
    for _, group in df.groupby('Category'):
        base_id = category_letter_to_id_start.get(group['Category'].iloc[0][0].upper(), 1000)
        ids.extend(range(base_id, base_id + len(group)))
    return ids

# Apply the function to generate the 'ID' column
source_df['ID'] = assign_ids(source_df)

# Reorder the columns with 'ID' as the first column
source_df = source_df[['ID'] + [col for col in source_df.columns if col != 'ID']]

# Write the processed data to a new Excel file
source_df.to_excel(output_file, index=False)

print(f"Data processed and saved to {output_file}")
