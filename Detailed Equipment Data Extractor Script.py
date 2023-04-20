#Data extraction and manipulation to get all Equipment Data per Store Number
import os
import pandas as pd

# Load the Excel file
file_path = r'C:\Users\Admin\Documents\Self learn\Data Analytics\Data Analytics Client 1\Audit Report - Master Copy.xlsx'
xl = pd.ExcelFile(file_path)

# Loop through the first 3 sheets
#sheet_names = xl.sheet_names[:3]

#Loop through each sheet
dfs = []
for sheet_name in xl.sheet_names:
    # Read in the specified range of rows
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=7, nrows=32)
    # Get the store number from the sheet name and create a column
    store_number = sheet_name.split(' ')[0]
    df['Store Number'] = store_number
    # Append the resulting dataframe to the list of dataframes
    dfs.append(df)
    # Delete the 23rd and 24th rows of the sheet

# Concatenate all dataframes in the list into a single dataframe
result_df = pd.concat(dfs)

# Join the "Store Number" column with the column values for each row
result_df['Values'] = result_df.apply(lambda row: [row['Store Number'], *row.drop('Store Number').tolist()], axis=1)
result_df = result_df[['Values']]

drop_indices = [22, 23, 31]  # List of index values to drop

#Iterate through rows and delete unwanted rows until EOF
########Working code here deleted
########Code available upon request


# Save the modified CSV file to a new file
new_file_path = r'C:\Users\Admin\Documents\Self learn\Data Analytics\Data Analytics Client 1\Modified Equipment Master Data.csv'
result_df.to_csv(new_file_path, index=False)
