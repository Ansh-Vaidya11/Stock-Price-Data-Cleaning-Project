import pandas as pd

def partition_rows(input_file, output_file, column_name):
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"File not found: {input_file}")
        return
    except Exception as e:
        print(f"An error occurred while reading the file: {str(e)}")
        return

    if column_name not in df.columns:
        print(f"Column '{column_name}' not found in the input dataset.")
        return

    # Create a list to store individual group DataFrames
    grouped_dataframes = []

    # Create a list to store ranked dataframes
    ranked_dataframes = []

    # Iterate through the rows in groups of 3
    for i in range(len(df) - 2):
        group = df.iloc[i:i + 3]

        # Calculate ranks for the current group
        group['Rank'] = group[column_name].rank(ascending=False)

        # Append the group to both lists
        grouped_dataframes.append(group)
        ranked_dataframes.append(group)

    # Concatenate the individual group DataFrames into a single DataFrame
    grouped_df = pd.concat(grouped_dataframes, ignore_index=True)

    # Concatenate the individual ranked DataFrames into a single DataFrame
    ranked_df = pd.concat(ranked_dataframes, ignore_index=True)

    try:
        grouped_df.to_excel(output_file, index=False)
        ranked_df.to_excel(output_file, index=False)
        print(f"Grouped rows and Rankings saved to {output_file}")
    except Exception as e:
        print(f"An error occurred while saving the data: {str(e)}")

# Usage example:
input_file = '/Users/Ansh/Desktop/test.xlsx'  # Replace with your input file path
output_file = 'output12.xlsx'  # Replace with your output file path
column_name = 'Actual Price'  # Replace with the actual column name

partition_rows(input_file, output_file, column_name)
