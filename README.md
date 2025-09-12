# taas


# Create a list of all potential value columns from your configuration
value_columns = [col_val1, col_val2, col_val3] # Add col_val4, etc., here if you use them

# Create a temporary list of parts, adding a value only if the cell is NOT empty
parts = [str(row[col]) for col in value_columns if pd.notna(row[col]) and str(row[col]).strip() != '']

# Join the non-blank parts with a dot
value = ".".join(parts)
