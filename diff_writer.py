import xlsxwriter
from xlsxwriter.worksheet import Format
from difflib import SequenceMatcher


def output_excel_formatting(orig: str, new: str, f_uline: Format, f_stkout: Format, f_norm: Format) -> list:

    if orig is None:  # New item that didn't exist in orig. Use underline style.
        return [f_uline, new]
    elif orig == new:  # No change since orig, use normal style.
        return [f_norm, orig]
    elif new is None:  # Completely deleted orig, use strikeout style.
        return [f_stkout, orig]

    # Matcher gets a sequence of "opcodes" which say to insert, delete, etc. blocks of text.
    matcher = SequenceMatcher(None, orig, new)
    opcodes = matcher.get_opcodes()

    result = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            result.append(f_norm)
            result.append(orig[i1:i2])
        elif tag == 'insert':
            result.append(f_uline)
            result.append(new[j1:j2])
        elif tag == 'delete':
            if orig[i1:i2] != "\n":
                result.append(f_stkout)
                result.append(orig[i1:i2].replace("\n", ""))
        elif tag == 'replace':
            if orig[i1:i2] != "\n":
                result.append(f_stkout)
                result.append(orig[i1:i2].replace("\n", ""))
            result.append(f_uline)
            result.append(new[j1:j2])

    return result


def output_to_excel_worksheet(diff_list: list, ofilename: str) -> None:
    # Open a workbook and choose the column to format
    workbook = xlsxwriter.Workbook(ofilename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column(first_col=0, last_col=2, width=40)
    # Set up some formats to use.
    normal = workbook.add_format({'text_wrap': True})
    underline = workbook.add_format({'underline': True, 'font_color': 'green', 'text_wrap': True})
    strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red', 'text_wrap': True})

    for (s1, s2), row in zip(diff_list, range(len(diff_list))):
        worksheet.write(row, 0, s1, normal)
        worksheet.write(row, 1, s2, normal)
        diff = output_excel_formatting(s1, s2, underline, strikeout, normal)

        if len(diff) < 3:  # Either no change, or didn't exist in original version.
            worksheet.write(row, 2, diff[1], diff[0])
        else:
            worksheet.write_rich_string(row, 2, *diff)

    workbook.close()


if __name__ == '__main__':
    string_a = "The quick brown fox jumped over the lazy dog"
    string_b = "The slow brown fox walked over a lazy cat"
    string_c = "How now brown cow"
    string_d = "How then brown hen"

    output_to_excel_worksheet([[string_c, string_d], [string_a, string_b]])
