from openpyxl import load_workbook
from redlines import Redlines

wb = load_workbook("test.xlsx")
ws = wb.active

with open('output.txt', mode='w') as ofile:
    ofile.write("|Orig Value|New Value|Diff|\n|-|-|-|\n")

    for row in ws.values:
        value_list = []
        for cell in row:
            value_list.append(cell.replace("\n", "<br>"))

        test = Redlines(value_list[0], value_list[1], markdown_style="none")
        ofile.write(f"|{value_list[0]}|{value_list[1]}|{test.output_markdown}|\n")
        print(f"|{value_list[0]}|{value_list[1]}|{test.output_markdown}|\n")
