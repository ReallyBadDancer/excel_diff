import pandas as pd


def compare_spreadsheets(original_file, updated_file):
    """
    Takes in two spreadsheets, each with three fields: ID, Description, and Clarifying Information.
    Output is a dictionary that is a merged version of the spreadsheets indexed to the updated file that
    has the contents of the description and clarifying information side by side for comparison in the redline tool.
    :param original_file: Contains old description and clarifying information data
    :param updated_file: Contains new description and clarifying information data
    :return: Dictionary containing the merged data with old/new fields side by side with _old and _new suffixes as keys.
    """

    # Read the original and updated spreadsheets
    original_df = pd.read_excel(original_file, engine='openpyxl')
    updated_df = pd.read_excel(updated_file, engine='openpyxl')
    updated_df.reset_index(inplace=True)

    # Merge the dataframes on the ID column
    merged_df = pd.merge(original_df, updated_df, on='ID', how='outer', suffixes=('_old', '_new'))
    merged_df['index'] = merged_df['index'].fillna(len(merged_df))
    # Sort the merged dataframe by the original order

    # merged_df['Order'] = merged_df['Order_new'].fillna(len(merged_df))

    merged_df.sort_values(by='index', inplace=True)
    # Create a list of dictionaries with ID, old description, and new description
    result = []
    for _, row in merged_df.iterrows():
        entry = {
            'ID': row['ID'],
            'Old Description': row['Description_old'] if pd.notna(row['Description_old']) else '',
            'New Description': row['Description_new'] if pd.notna(row['Description_new']) else '',
            'Old Clarifying Information': row['Clarifying Information_old'] if pd.notna(
                row['Clarifying Information_old']) else '',
            'New Clarifying Information': row['Clarifying Information_new'] if pd.notna(
                row['Clarifying Information_new']) else ''
        }
        result.append(entry)

    return result


# Example usage
original_file = 'original.xlsx'
updated_file = 'updated.xlsx'
comparison_result = compare_spreadsheets(original_file, updated_file)

# Output comparison result to excel
df = pd.DataFrame(comparison_result)
df.to_excel("compare_file.xlsx", index=False)
