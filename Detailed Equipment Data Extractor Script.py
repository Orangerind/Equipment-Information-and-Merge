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
or index, row in result_df.iterrows():
    mask = result_df.index.isin(drop_indices)  # Boolean mask to identify rows to drop
    result_df.drop(result_df[mask].index, inplace=True)  # Drop the rows corresponding to the mask

# Split the lists in into separate columns
result_df = result_df['Values'].apply(pd.Series)
result_df = result_df.rename(columns={0:'Store Number', 1:'Equipment',2:'Number',3:'Quantity', 4:'Repair', 5:'Replace', 6:'Notes'})
result_df.drop(columns='Number', inplace=True)

#Some script to clean the dataframes further 

# Loop through the 'Quantity' column and convert values to integers
for i in range(len(dfs['Quantity'])):
    try:
        dfs.at[i, 'Quantity'] = int(dfs.at[i, 'Quantity'])
    except ValueError:
        dfs.at[i, 'Quantity'] = 0

# Replace non-positive values with 0
dfs['Quantity'] = dfs['Quantity'].apply(lambda x: max(0, x))

# Define a function to create the switch-case
def create_switch_case(unique_values, ones_list, zeros_list):
    switch_case = {}
    for value in unique_values:
        if value in ones_list:
            switch_case[value] = 1
        elif value in zeros_list:
            switch_case[value] = 0
        else:
            switch_case[value] = None
    return switch_case

# Specify which values should be encoded as 1 or 0 in the switch-case
repair_ones_list = ['y', ' y','Y ','Y','y  ','y ','yes','y/N', '1','2','Repair','Y/N','3','B','b',' Y',' y ', '10','y7']
repair_zeros_list = ['n',None,'N','`n', ' ', 'n ', '0',' n', 'no', 'no ', '?']

replace_ones_list = [ ' y',  'y' ,'Y ', 'Y',  'y ', ' Y', '2', 'Y/N' ,' y ',  'y-2', 'y-1',
 'yes', 'y/N', 'U', '1' ,'y  ', 'Replace', 'B' , 'b', 'Y  ', '3' ,'Y/',  '4', 'poss', 'Y(1)', '2 washer, 1 coolant']
replace_zeros_list = ['n', None, 'N', 'n ', ' ', 'N ', ' n', 'no', 'no ', '?', '0']

# Create the switch-case for 'Repair' column
repair_unique_values = dfs['Repair'].unique()
repair_switch_case = create_switch_case(repair_unique_values, repair_ones_list, repair_zeros_list)

# Create the switch-case for 'Replace' column
replace_unique_values = dfs['Replace'].unique()
replace_switch_case = create_switch_case(replace_unique_values, replace_ones_list, replace_zeros_list)

# Apply switch-case to 'Repair' column
dfs['Repair'] = dfs['Repair'].map(repair_switch_case)

# Apply switch-case to 'Replace' column
dfs['Replace'] = dfs['Replace'].map(replace_switch_case)

dfs.dropna()


# Save the modified CSV file to a new file
new_file_path = r'C:\Users\Admin\Documents\Self learn\Data Analytics\Data Analytics Client 1\Modified Equipment Master Data.csv'
result_df.to_csv(new_file_path, index=False)
