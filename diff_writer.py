import xlsxwriter
from xlsxwriter.worksheet import Format
from difflib import SequenceMatcher
import pandas as pd


class DiffWriter:
    """
    Takes two Excel files with an ID field and any number of data fields, and creates a new Excel file with the old
    and new data merged on the ID field. Input Excel files must have the following columns:
    ID: Jama ID
    Data Fields: Columns of data with identical field names between orig and updated spreadsheets.
    """
    def __init__(self, orig_fname: str, updated_fname: str, redline_cols: tuple, dest_fname: str):
        self.orig_fname = orig_fname  # Excel file with original requirement data
        self.updated_fname = updated_fname  # Excel file with modified requirement data
        self.dest_fname = dest_fname  # Output Excel file
        (self.workbook,
         self.worksheet,
         self.normal,
         self.underline,
         self.strikeout) = self.create_excel_worksheet_and_formats()  # Open workbook and create formats
        self.df = self.compare_spreadsheets()  # Create a dataframe with orig, updated, orig -> updated redline data.
        self.df.fillna("", inplace=True)  # Replace NaN with empty strings
        print("Created dataframe from orig/dest requirement data.")
        self.redline_cols = redline_cols
        for col in redline_cols:
            self.add_redline_column_to_df(col)
        print("Done creating redlines for requirement data.")
        self.output_to_excel()
        print("Done writing output to Excel.")

    def create_excel_worksheet_and_formats(self):
        """
        Create a workbook, add a worksheet to the workbook, and add basic formats representing normal, inserted, and
        deleted text.
        :return: Workbook, Worksheet, Formats
        """
        # Open a workbook and choose the column to format
        workbook = xlsxwriter.Workbook(self.dest_fname)
        worksheet = workbook.add_worksheet()
        worksheet.set_column(first_col=0, last_col=25, width=50)
        # Set up some formats to use.
        normal = workbook.add_format({'text_wrap': True})
        underline = workbook.add_format({'underline': True, 'font_color': 'green', 'text_wrap': True})
        strikeout = workbook.add_format({'font_strikeout': True, 'font_color': 'red', 'text_wrap': True})

        return workbook, worksheet, normal, underline, strikeout

    def compare_spreadsheets(self) -> pd.DataFrame:
        """
        Takes in two spreadsheets, each with identical fields: ID, and an arbitrary number of other data fields.
        Output is a dictionary that is a merged version of the spreadsheets indexed to the updated file that
        has the contents of the data fields side by side for comparison in the redline tool.

        :return: Dataframe containing the merged spreadsheet data.
        """

        # Read the original and updated spreadsheets
        original_df = pd.read_excel(self.orig_fname, engine='openpyxl')
        updated_df = pd.read_excel(self.updated_fname, engine='openpyxl')
        updated_df.reset_index(inplace=True)

        # Merge the dataframes on the ID column
        merged_df = pd.merge(original_df, updated_df, on='ID', how='outer', suffixes=('_old', '_new'))
        merged_df['index'] = merged_df['index'].fillna(len(merged_df))

        # Sort the merged dataframe by the original order
        merged_df.sort_values(by='index', inplace=True)
        return merged_df

    def add_redline_column_to_df(self, column_name: str) -> None:
        """
        Apply the create_redline to each row in self.df and create a new column with the results for the given column.
        :param column_name: The name of the column to be compared. Must have _old and _new suffixes.
        :return: None
        """
        self.df[f"{column_name}_redline"] = self.df.apply(lambda row: self.create_redline(row[f"{column_name}_old"], row[f"{column_name}_new"]), axis=1)

    def create_redline(self, orig: str, new: str) -> list:
        """
        Compares two strings, orig and new. Returns a result as a list of formats and string slices that will tell
        the calling function how to format the resulting text in Excel.
        :param orig: Name of column with old data.
        :param new: Name of column with new data.
        :return result: A variable length list of alternating Excel formats and string slices describing the redline
        text, which can be either normal (no change), underline (all new), strikeout (all deleted), or a longer list
        of alternating formats and string slices representing multiple insertions and deletions.
        """
        if orig is None:  # New item that didn't exist in orig. Use underline style.
            return [self.underline, new]
        elif orig == new:  # No change since orig, use normal style.
            return [self.normal, orig]
        elif new is None:  # Completely deleted orig, use strikeout style.
            return [self.strikeout, orig]

        # Matcher gets a sequence of "opcodes" which say to insert, delete, etc. blocks of text.
        matcher = SequenceMatcher(None, orig, new)
        opcodes = matcher.get_opcodes()

        result = []
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == 'equal':
                result.append(self.normal)
                result.append(orig[i1:i2])  # Insert the old text w/ plain formatting
            elif tag == 'insert':
                result.append(self.underline)
                result.append(new[j1:j2])  # Insert the new text w/ underlined green formatting
            elif tag == 'delete':
                if orig[i1:i2] != "\n":
                    result.append(self.strikeout)
                    result.append(orig[i1:i2].replace("\n", ""))  # Insert the old text with strikeout red.
            elif tag == 'replace':
                if orig[i1:i2] != "\n":
                    result.append(self.strikeout)
                    result.append(orig[i1:i2].replace("\n", ""))  # Insert old text with strikeout red...
                result.append(self.underline)
                result.append(new[j1:j2])  # ...then insert new text with underline green.

        return result

    def output_to_excel(self):
        """
            Creates an Excel workbook and adds a worksheet to the workbook, adding some basic formatting to the columns.
            Creates styles for new text and deleted text, then passes those along with the diff_list to
            output_excel_formatting, which will output string A, B, and a diff that looks like a redline from A to B.
            Finally, it writes all the [A, B, Redline] items to the workbook.

            :param diff_list: A list of [A, B] strings to compare to each other.
            :param ofilename: Output file name.
            :return: None
            """

        headings = ["ID"]
        for heading in self.redline_cols:
            headings.extend([f"{heading}_old", f"{heading}_new", f"{heading}_redline"])
        print("Found headings to diff: ", headings)

        for heading in list(self.df):
            if heading not in headings:
                headings.append(heading)

        for inx, heading in enumerate(headings):
            self.worksheet.write(0, inx, heading, self.normal)

        for inx, row in enumerate(self.df.to_dict(orient='records'), start=1):
            for jnx, heading in enumerate(headings):
                if "redline" in heading:
                    if row[heading] and len(row[heading]) > 2:
                        self.worksheet.write_rich_string(inx, jnx, *row[heading])
                    elif row[heading] and len(row[heading]) == 2:
                        if not row[headings[jnx-2]]:  # Old description/CI doesn't exist.
                            self.worksheet.write(inx, jnx, row[headings[jnx-1]], self.underline)  # Write new text.
                        elif not row[headings[jnx-1]]:
                            self.worksheet.write(inx, jnx, row[headings[jnx-2]], self.strikeout)  # Write deleted txt.
                    else:
                        self.worksheet.write(inx, jnx, "", self.normal)  # Write blank string to cell
                else:
                    self.worksheet.write(inx, jnx, row[heading], self.normal)

        self.workbook.close()


if __name__ == '__main__':
    diff = DiffWriter("original.xlsx", "updated.xlsx", "output.xlsx")
