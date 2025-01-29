from openpyxl import load_workbook
from diff_writer import output_to_excel_worksheet
from pprint import pp

ifile = input("Enter name of excel file to diff: ")

if not ifile:
    ifile = 'test.xlsx'
elif ifile[-5:] != ".xlsx":
    ifile = f"{ifile}.xlsx"

print(f"Running diff on {ifile}") if ifile else print(f"Running diff on test.xlsx since no filename specified...\n")
wb = load_workbook(ifile)
ws = wb.active

# Walk through a spreadsheet with exactly two columns of data, and create a list of lists of the form:
# [cell A string, cell B string] for comparison in the diff_writer.output_to_excel_worksheet funciton.
row_list = []
for row in ws.values:
    value_list = []
    for cell in row:
        if cell and "\n\n" in cell:  # Removes annoying extra carriage returns in Jama exports.
            value_list.append(cell.replace("\n\n", "\n"))
        else:
            value_list.append(cell)
    if len(value_list) != 2:
        raise ValueError("Can only diff a two-column spreadsheet. "
                         "Please update spreadsheet and be sure to delete extra whitespace.")
    else:
        row_list.append(value_list)

# print("List of strings to diff:")
# pp(row_list)

print("Writing diff output to diff.xlsx")
output_to_excel_worksheet(row_list, "diff.xlsx")
