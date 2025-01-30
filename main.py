from diff_writer import DiffWriter

OLD_DATA_EXCEL_FNAME = "original.xlsx"
NEW_DATA_EXCEL_FNAME = "updated.xlsx"


if __name__ == "__main__":
    diff = DiffWriter(OLD_DATA_EXCEL_FNAME, NEW_DATA_EXCEL_FNAME, 'output.xlsx')
