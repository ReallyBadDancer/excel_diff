import xlsxwriter
from xlsxwriter.worksheet import Format
from difflib import SequenceMatcher


def output_excel_formatting(orig: str, new: str, f_uline: Format, f_stkout: Format, f_norm: Format) -> list:
    """
    Compares two strings, orig and new. Returns a result as a list of formats and string slices that will tell
    the calling function how to format the resulting text in Excel.
    :param orig:
    :param new:
    :param f_uline:
    :param f_stkout:
    :param f_norm:
    :return result: A variable length list of alternating Excel formats and string slices describing the redline
    text, which can be either normal (no change), underline (all new), strikeout (all deleted), or a longer list
    of alternating formats and string slices multiple insertions and deletions.
    """
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
            result.append(orig[i1:i2])  # Insert the old text w/ plain formatting
        elif tag == 'insert':
            result.append(f_uline)
            result.append(new[j1:j2])  # Insert the new text w/ underlined green formatting
        elif tag == 'delete':
            if orig[i1:i2] != "\n":
                result.append(f_stkout)
                result.append(orig[i1:i2].replace("\n", ""))  # Insert the old text with strikeout red.
        elif tag == 'replace':
            if orig[i1:i2] != "\n":
                result.append(f_stkout)
                result.append(orig[i1:i2].replace("\n", ""))  # Insert old text with strikeout red...
            result.append(f_uline)
            result.append(new[j1:j2])  # ...then insert new text with underline green.

    return result


def output_to_excel_worksheet(diff_list: list, ofilename: str) -> None:
    """
    Creates an Excel workbook and adds a worksheet to the workbook, adding some basic formatting to the columns.
    Creates styles for new text and deleted text, then passes those along with the diff_list to
    output_excel_formatting, which will output string A, B, and a diff that looks like a redline from A to B.
    Finally, it writes all of the [A, B, Redline] items to the workbook.

    :param diff_list: A list of [A, B] strings to compare to each other.
    :param ofilename: Output file name.
    :return: None
    """
    # Open a workbook and choose the column to format
    workbook = xlsxwriter.Workbook(ofilename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column(first_col=0, last_col=2, width=40)
    # Set up some formats to use.
    normal = workbook.add_format({'text_wrap': True})
    underline = workbook.add_format({'underline': True, 'font_color': 'green', 'text_wrap': True})
    strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red', 'text_wrap': True})

    for (s1, s2), row in zip(diff_list, range(len(diff_list))):
        worksheet.write(row, 0, s1, normal)  # Write the original text to column A
        worksheet.write(row, 1, s2, normal)  # write the updated text to column B
        diff = output_excel_formatting(s1, s2, underline, strikeout, normal)  # Generate the formatted redline list.

        if len(diff) < 3:  # Either no change, or didn't exist in original version. Just reinsert old text for both.
            worksheet.write(row, 2, diff[1], diff[0])  # diff[0] is probably always 'normal'
        else:  # Write the redline text to col C with formats as enumerated in list diff.
            worksheet.write_rich_string(row, 2, *diff)

    workbook.close()


if __name__ == '__main__':
    string_a = "The quick brown fox jumped over the lazy dog"
    string_b = "The slow brown fox walked over a lazy cat"
    string_c = "How now brown cow"
    string_d = "How then brown hen"

    output_to_excel_worksheet([[string_c, string_d], [string_a, string_b]])
