import pandas as pd
import json

# --- Configuration: CHANGE THESE VALUES ---
excel_file = 'your_data.xlsx'      # Your Excel file name
output_json_file = 'output.json'   # The name of the JSON file to create

# The 1-based row numbers from your Excel sheet
start_row = 5                      # The first row you want to include (e.g., row 5)
end_row = 20                       # The last row you want to include (e.g., row 20)

# The names of the columns you want to use
col_key = 'a'       # Column for the JSON key
col_val1 = 'b'      # First part of the JSON value
col_val2 = 'c'      # Second part of the JSON value
col_val3 = 'd'      # Third part of the JSON value
# ------------------------------------------

# Initialize an empty dictionary to store the final JSON data
json_data = {}

try:
    # Read the specified columns from the Excel file
    # We use `usecols` to only load the data we need, which is more efficient
    df = pd.read_excel(excel_file, usecols=[col_key, col_val1, col_val2, col_val3])

    # Select the specified range of rows.
    # We use `iloc` which is 0-indexed, so we subtract 1 from the start_row.
    # The end_row in the slice is exclusive, so it works perfectly.
    df_slice = df.iloc[start_row - 1 : end_row]

    # Iterate over each row in our selected data slice
    for index, row in df_slice.iterrows():
        # Get the key from the first column
        key = str(row[col_key])
        
        # Create the concatenated value string, ensuring all parts are strings
        value = f"{str(row[col_val1])}.{str(row[col_val2])}.{str(row[col_val3])}"
        
        # Add the new key-value pair to our dictionary
        json_data[key] = value

    # Write the dictionary to a JSON file
    # `indent=4` makes the JSON file human-readable (pretty-printed)
    with open(output_json_file, 'w') as f:
        json.dump(json_data, f, indent=4)

    print(f"✅ Successfully created '{output_json_file}' with data from rows {start_row} to {end_row}.")

except FileNotFoundError:
    print(f"❌ Error: The file '{excel_file}' was not found.")
except KeyError as e:
    print(f"❌ Error: A column was not found in the Excel file. Please check your column names. Details: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
