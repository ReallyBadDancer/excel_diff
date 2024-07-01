import xlsxwriter
from difflib import SequenceMatcher


workbook = xlsxwriter.Workbook('rich_strings.xlsx')
worksheet = workbook.add_worksheet()

# Set up some formats to use.
normal = workbook.add_format()
bold = workbook.add_format({'bold': True})
italic = workbook.add_format({'italic': True})
underline = workbook.add_format({'underline': True, 'font_color': 'green'})
strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red'})
wrap = workbook.add_format()
wrap.set_text_wrap()


worksheet.set_column('A:A', 30)


worksheet.write_rich_string(0, 0,
                            normal, 'This is ',
                            bold, 'bold',
                            normal, ' and this is \n',
                            italic, 'italic',
                            normal, ' and this is ',
                            underline, 'underline',
                            normal, ' and this is \n',
                            strikeout, 'strikeout.\n',
                            normal, 'Finally, changed back to normal',
                            wrap)



s1 = "abcdeflmnop"
s2 = "abcdefg\nhijk"

matcher = SequenceMatcher(None, s1, s2)
opcodes = matcher.get_opcodes()
print(opcodes)


def output_excel(orig:str, new:str)->list:
    result = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            result.append(normal)
            result.append("".join(new[i1:i2]))
        elif tag == 'insert':
            result.append(underline)
            result.append("".join(new[j1:j2]))
        elif tag == 'delete':
            result.append(strikeout)
            result.append("".join(new[j1:j2]))
        elif tag == 'replace':
            result.append(strikeout)
            result.append("".join(orig[i1:i2]))
            result.append(underline)
            result.append("".join(new[j1:j2]))

    result.append(wrap)
    return result


excel_cell = output_excel(s1, s2)
worksheet.write_rich_string(1, 0, *excel_cell)
workbook.close()



