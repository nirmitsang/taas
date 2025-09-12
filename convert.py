import pandas as pd
import json

# --- Configuration: CHANGE THESE VALUES ---
excel_file = 'your_data.xlsx'
output_json_file = 'output.json'
start_row = 1
end_row = 20 # Example end row

# Define all the columns you'll be using
col_key = 'a'
col_val1 = 'b'
col_val2 = 'c'
col_val3 = 'd'
col_val4 = 'e' # Added the 4th value column
# ------------------------------------------

json_data = {}

# List of all columns to read from the Excel file
all_cols_to_read = [col_key, col_val1, col_val2, col_val3, col_val4]

try:
    df = pd.read_excel(excel_file, usecols=all_cols_to_read)
    df_slice = df.iloc[start_row - 1 : end_row]

    for index, row in df_slice.iterrows():
        # --- THIS IS THE MODIFIED KEY LOGIC ---
        # Read the raw value from the key column
        key_raw = str(row[col_key])
        
        # Split the string on the first dot and take the part after it
        # If no dot, it uses the original key_raw value
        key = key_raw.split('.', 1)[1] if '.' in key_raw else key_raw
        # ---------------------------------------

        # Define the columns that should be joined for the value
        value_columns = [col_val1, col_val2, col_val3, col_val4] 

        # Build a list of values, but only if they are not empty/blank
        parts = [str(row[col]) for col in value_columns if pd.notna(row[col]) and str(row[col]).strip() != '']
        
        # Join the valid parts with a dot
        value = ".".join(parts)

        # Add the new key-value pair, but only if the value isn't empty
        if value:
            json_data[key] = value

    with open(output_json_file, 'w') as f:
        json.dump(json_data, f, indent=4)

    print(f"✅ Successfully created '{output_json_file}'. Keys were trimmed and blank entries were skipped.")

except FileNotFoundError:
    print(f"❌ Error: The file '{excel_file}' was not found.")
except KeyError as e:
    print(f"❌ Error: A column was not found. Please check your column names. Details: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
