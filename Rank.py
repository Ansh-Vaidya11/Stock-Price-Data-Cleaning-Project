import pandas as pd

# Input and output file paths
input_file = '/Users/Ansh/Desktop/test.xlsx'  # Replace with your input file path
output_file = 'output1.xlsx'  # Replace with your output file path

# Load the Excel file into a DataFrame
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"File not found: {input_file}")
    exit(1)
except Exception as e:
    print(f"An error occurred while reading the file: {str(e)}")
    exit(1)

# Specify the column name you want to rank
column_name = 'Actual Price'  # Replace with the actual column name

if column_name not in df.columns:
    print(f"Column '{column_name}' not found in the input dataset.")
    exit(1)

# Create a list to store individual group DataFrames
grouped_dataframes = []

# Iterate through the rows in increments of 3 and calculate ranks in descending order
for i in range(0, len(df), 3):
    group = df.iloc[i:i + 3]
    group['Rank'] = group[column_name].rank(ascending=False)
    grouped_dataframes.append(group)

# Concatenate the individual group DataFrames into a single DataFrame
ranked_df = pd.concat(grouped_dataframes, ignore_index=True)

# Save the DataFrame with ranks to a new Excel file
try:
    ranked_df.to_excel(output_file, index=False)
    print(f"Rankings saved to {output_file}")
except Exception as e:
    print(f"An error occurred while saving the rankings: {str(e)}")
