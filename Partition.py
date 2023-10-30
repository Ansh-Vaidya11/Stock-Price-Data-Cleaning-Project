import pandas as pd

def partition(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"File not found: {input_file}")
        return
    except Exception as e:
        print(f"An error occurred while reading the file: {str(e)}")
        return

    # Exclude specified columns and columns with all null values
    columns_to_rank = [col for col in df.columns if col not in ['Name', 'datetime'] and not df[col].isnull().all()]

    # Create a list to store the data for writing to Excel
    data_to_write = []

    # Iterate through the rows in groups of 3
    for i in range(len(df) - 2):
        group = df.iloc[i:i + 3].copy()

        for col in columns_to_rank:
            group[col + '_rank'] = group[col].rank(ascending=False)

        # Append the group to the list
        data_to_write.append(group)

    # Concatenate the individual DataFrames and reset the index
    result_df = pd.concat(data_to_write, ignore_index=True)

    try:
        # Save the grouped and ranked DataFrame to Excel
        result_df.to_excel(output_file, index=False)
        print(f"Grouped rows and Rankings saved to {output_file}")
    except Exception as e:
        print(f"An error occurred while saving the data: {str(e)}")

# Usage example:
input_file = '/Users/Ansh/Desktop/ACC.xlsx'  # Replace with your input file path
output_file = 'ACC_Output.xlsx'  # Replace with your output file path

partition(input_file, output_file)