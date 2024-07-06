import xlsxwriter
from xlsxwriter.worksheet import Format
from difflib import SequenceMatcher


def output_excel_formatting(orig: str, new: str, f_uline: Format, f_sout: Format, f_norm: Format, wr: Format) -> list:
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
            result.append("".join(orig[i1:i2]))
        elif tag == 'replace':
            result.append(f_sout)
            result.append("".join(orig[i1:i2]))
            result.append(f_uline)
            result.append("".join(new[j1:j2]))

    result.append(wr)
    return result


def output_to_excel_worksheet(diff_list: list) -> None:
    # Open a workbook and choose the column to format
    workbook = xlsxwriter.Workbook('rich_strings.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(first_col=0, last_col=2, width=40)
    # Set up some formats to use.
    normal = workbook.add_format()
    underline = workbook.add_format({'underline': True, 'font_color': 'green'})
    strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red'})
    wrap = workbook.add_format()
    wrap.set_text_wrap()

    for (s1, s2), row in zip(diff_list, range(len(diff_list))):
        worksheet.write(row, 0, s1)
        worksheet.write(row, 1, s2)
        diff = output_excel_formatting(s1, s2, underline, strikeout, normal, wrap)
        worksheet.write_rich_string(row, 2, *diff)

    # worksheet.write_rich_string(0, 2, *excel_data)
    workbook.close()


if __name__ == '__main__':
    string_a = "The quick brown fox jumped over the lazy dog"
    string_b = "The slow brown fox walked over a lazy cat"
    string_c = "How now brown cow"
    string_d = "How then brown hen"
    output_to_excel_worksheet([[string_c, string_d], [string_a, string_b]])

    # Step by step test
    # workbook = xlsxwriter.Workbook('xlsxwriter_test.xlsx')
    # worksheet = workbook.add_worksheet()
    # worksheet.set_column(first_col=0, last_col=2, width=30)
    # # Set up some formats to use.
    # normal = workbook.add_format()
    # underline = workbook.add_format({'underline': True, 'font_color': 'green'})
    # strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red'})
    # wrap = workbook.add_format()
    # wrap.set_text_wrap()
    # format_list = [normal, "Test", underline, " Underline"]
    # worksheet.write_rich_string(0, 0, *format_list)
    # workbook.close()

