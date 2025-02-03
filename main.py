from diff_writer import DiffWriter

OLD_DATA_EXCEL_FNAME = "original.xlsx"
NEW_DATA_EXCEL_FNAME = "updated.xlsx"
REDLINE_COLS = (
    "Name",
    "Description",
    "Clarifying Information",
    "Release",
    "Verification Method",
    "Status",

)

if __name__ == "__main__":
    diff = DiffWriter(OLD_DATA_EXCEL_FNAME, NEW_DATA_EXCEL_FNAME, REDLINE_COLS, 'output.xlsx')
