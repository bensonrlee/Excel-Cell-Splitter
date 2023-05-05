import os
import pandas as pd
import time
import string

# Ask the user for the filename of the Excel file and the letter of the column with comma-separated or newline-separated values
file_name = input("Enter the filename of the Excel file (or full path if not in current directory): ")
column_letter = input("Enter the letter of the column with comma-separated and newline-separated values (e.g. A, B, C, ...): ")

# Convert the column letter to a zero-indexed column number
column_number = string.ascii_uppercase.index(column_letter.upper())

# Parse the filename to remove any surrounding double quotes
file_name = file_name.strip('"')

# Check if the file exists
if not os.path.isfile(file_name):
    print(f"Error: File '{file_name}' does not exist.")
    exit()

# Ask the user for the output filename
output_file_name = input("Enter the output filename (or full path if not in current directory): ")
output_file_name = output_file_name.strip('"')

# Start a timer
start_time = time.time()

try:
    # Load the Excel file into a pandas dataframe
    df = pd.read_excel(file_name)

    # Split the specified column by comma and newline using a regular expression, and explode the resulting dataframe
    df1 = df.assign(**{df.columns[column_number]: df.iloc[:, column_number].str.split("[,\n]")}).explode(df.columns[column_number])

    # Write the modified dataframe to the output file
    with pd.ExcelWriter(output_file_name) as writer:
        df1.to_excel(writer, index=False)

    # Calculate the time taken and display a message on bright green text
    elapsed_time = time.time() - start_time
    print(f"\033[92mTransformed file {output_file_name} in {elapsed_time:.2f} seconds.\033[0m")
except Exception as e:
    print(f"Error: {e}")
