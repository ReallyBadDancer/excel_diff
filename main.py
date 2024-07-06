from openpyxl import load_workbook
from redlines import Redlines
from diff_writer import output_to_excel_worksheet
from pprint import pp

wb = load_workbook("test.xlsx")
ws = wb.active

row_list = []
for row in ws.values:
    value_list = []
    for cell in row:
        # value_list.append(cell.replace("\n", "<br>"))
        value_list.append(cell)
    row_list.append(value_list)

pp(row_list)

output_to_excel_worksheet(row_list)


# Step 1: Import the workbook in OpenPyxl with the test text and save the contents as a dictionary.
    # Dictionary format: {'Old Text': [A1, A2 ... An], 'New Text': [B1, B2 ... Bn]}

# Step 2: Create a third list in the dictionary that contains the deltas between A and B.
    # Add empty list to dictionary.
    # Iterate through the A, B lists and compare each A cell to each B cell using redlines,
        # append result to C in Markdown format.

# Step 3: Create an xlsx workbook and fonts to use.
    # Create the target workbook and worksheet in xlsxwriter.
    # Create the required fonts (underline, strikeout) as variables.

# Step 4: Walk through list C, divide the list items into a list of strings and formats
    # For each element in C
        #
