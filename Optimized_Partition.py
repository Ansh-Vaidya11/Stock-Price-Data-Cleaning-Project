import os
import pandas as pd

def calculate_ranks(df, columns_to_rank):
    # Calculate the rank for specified columns
    df_ranks = df[columns_to_rank].rank(ascending=False)
    df_ranks.columns = [col + '_rank' for col in columns_to_rank]
    return pd.concat([df, df_ranks], axis=1)

def partition_rows(input_folder, output_folder):
    # Get a list of all Excel files in the input folder
    excel_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

    for excel_file in excel_files:
        input_file = os.path.join(input_folder, excel_file)
        output_file = os.path.join(output_folder, f"{os.path.splitext(excel_file)[0]}_output.xlsx")

        try:
            df = pd.read_excel(input_file)

            # Exclude specified columns and columns with all null values
            columns_to_rank = [col for col in df.columns if col not in ['Name', 'datetime'] and not df[col].isnull().all()]

            # Iterate through the rows in groups of 3
            grouped_data = []
            for i in range(len(df) - 2):
                group = df.iloc[i:i + 3].copy()
                grouped_data.append(group)

            # Calculate ranks for specified columns
            for i in range(len(grouped_data)):
                grouped_data[i] = calculate_ranks(grouped_data[i], columns_to_rank)

            # Concatenate the individual DataFrames and reset the index
            result_df = pd.concat(grouped_data, ignore_index=True)

            try:
                # Save the grouped and ranked DataFrame to Excel
                result_df.to_excel(output_file, index=False)
                print(f"Grouped rows and Rankings saved to {output_file}")
            except Exception as e:
                print(f"An error occurred while saving the data: {str(e)}")

        except FileNotFoundError:
            print(f"File not found: {input_file}")
        except Exception as e:
            print(f"An error occurred while reading the file {excel_file}: {str(e)}")

# Usage example:
input_folder = '/Users/Ansh/Desktop/Project'  # Replace with your input folder path
output_folder = '/Users/Ansh/Desktop/Project'  # Replace with your output folder path

partition_rows(input_folder, output_folder)