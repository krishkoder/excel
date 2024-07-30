import pandas as pd
import os

def remove_duplicates(file_path):
    # Load the Excel file
    data = pd.read_excel(file_path)
    
    # Remove duplicates
    data_cleaned = data.drop_duplicates()
    
    # Save the cleaned data back to an Excel file
    cleaned_file_path = file_path.replace('.xlsx', '_cleaned.xlsx')
    data_cleaned.to_excel(cleaned_file_path, index=False)
    return cleaned_file_path

# List of files to process
files = [
    'Long Beach, CA.xlsx',
    'Los Angeles, CA.xlsx',
    'Napa, CA.xlsx',
    'Newport Beach, CA.xlsx',
    'Oakland, CA.xlsx',
    'Pasadena, CA.xlsx',
    'Perris Valley, CA.xlsx',
    'Rosamond, CA.xlsx',
    'Sacramento, CA.xlsx',
    'San Francisco, CA.xlsx'
]

# Directory containing the files
directory = '/path/to/your/files/'

# Process each file
for file_name in files:
    file_path = os.path.join(directory, file_name)
    cleaned_file_path = remove_duplicates(file_path)
    print(f'Cleaned file saved as: {cleaned_file_path}')