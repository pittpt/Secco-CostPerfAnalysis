import pandas as pd
import os

# Define the path to the directory containing the Excel files
folder_path = 'Source'

# List all Excel files in the directory
excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

# Initialize an empty DataFrame to hold all the data
combined_df = pd.DataFrame()

# Loop through each file, clean the data, and append it to combined_df
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df_src = pd.read_excel(file_path)
    
    # Perform the specified cleaning operations
    df_clean = df_src[8:]
    df_clean.columns = df_src.iloc[6]  # Rename columns
    df_clean = df_clean.reset_index(drop=True)
    df_clean = df_clean.dropna(subset=['P/O No.'])
    df_clean = df_clean.rename(columns=lambda x: x.strip())  # Remove any trailing spaces from column names
    selected_features = ['P/O No.', 'P/O Date', 'Vendor Name', 'Job Code', 'Job Name', 'Ref.Code', 'Deli. Date', 'Cost Code', 'Material Code', 'Unit', 'Qty', 'Price/Unit', 'Amount', 'Discount', 'VAT', 'Net Amount']
    df_clean = df_clean[selected_features]
    print(df_clean.columns)
    
    # Append to the combined DataFrame
    combined_df = pd.concat([combined_df, df_clean], ignore_index=True)


# Define the path and filename for the output CSV
output_path = 'costs_data2.csv'

# Write the DataFrame to a CSV file
combined_df.to_csv(output_path, index=False)