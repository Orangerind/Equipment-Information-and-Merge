# Read the data into a pandas DataFrame
file_path = r'C:\Users\Admin\Documents\Self learn\Data Analytics\Data Analytics Client 1\Modified Equipment Master Data.xlsx'
df = pd.read_excel(file_path)

# Loop through the 'Quantity' column and convert values to integers
for i in range(len(df['Quantity'])):
    try:
        df.at[i, 'Quantity'] = int(df.at[i, 'Quantity'])
    except ValueError:
        df.at[i, 'Quantity'] = 0

# Replace non-positive values with 0
df['Quantity'] = df['Quantity'].apply(lambda x: max(0, x))

# Drop the 'Equipment' column
df.drop('Equipment', axis=1, inplace=True)

# Group the data by 'Store Number' and sum the values across the rows for 'Quantity', 'Repair', and 'Replace' columns
df = df.groupby('Store Number')['Quantity', 'Repair', 'Replace'].sum().reset_index()

# Read the second dataframe from the Excel file
file_path2 = r'C:\Users\Admin\Documents\Self learn\Data Analytics\Data Analytics Client 1\Store Information Final Dataset.xlsx'
df2 = pd.read_excel(file_path2)

# Merge the two dataframes on the 'Store Number' column
merged_df = pd.merge(df2, df, on='Store Number')

# Save the modified CSV file to a new file
new_file_path = r'C:\Users\Admin\Documents\Self learn\Data Analytics\Data Analytics Client 1\Final Store and Equipment Totals Dataset.csv'
merged_df.to_csv(new_file_path, index=False)