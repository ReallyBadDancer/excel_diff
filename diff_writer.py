import xlsxwriter
from difflib import SequenceMatcher


def output_excel_formatting(orig: str, new: str, f_uline, f_sout, f_norm, wr) -> list:
    # Matcher gets a sequence of "opcodes" which say to insert, delete, etc. blocks of text.
    matcher = SequenceMatcher(None, orig, new)
    opcodes = matcher.get_opcodes()

    result = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            result.append(f_norm)
            result.append("".join(orig[i1:i2]))
        elif tag == 'insert':
            result.append(f_uline)
            result.append("".join(new[j1:j2]))
        elif tag == 'delete':
            result.append(f_sout)
            result.append("".join(new[j1:j2]))
        elif tag == 'replace':
            result.append(f_sout)
            result.append("".join(orig[i1:i2]))
            result.append(f_uline)
            result.append("".join(new[j1:j2]))

    result.append(wr)
    return result


def output_to_excel_worksheet(s1, s2, col, row):
    # Open a workbook and choose the column to format
    workbook = xlsxwriter.Workbook('rich_strings.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(first_col=col, last_col=col, width=100)
    # Set up some formats to use.
    normal = workbook.add_format()
    underline = workbook.add_format({'underline': True, 'font_color': 'green'})
    strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red'})
    wrap = workbook.add_format()
    wrap.set_text_wrap()
    # Generate the diff in xlsxwriter output format.
    excel_data = output_excel_formatting(s1, s2, underline, strikeout, normal, wrap)
    # Write the xlsxwriter data to the target cell and close workbook
    worksheet.write_rich_string(row, col, *excel_data)
    workbook.close()


if __name__ == '__main__':
    string_a = "The quick brown fox jumped over the lazy dog"
    string_b = "The slow brown fox walked over a lazy cat"
    output_to_excel_worksheet(string_a, string_b, 0, 0)
