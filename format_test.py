import xlsxwriter



workbook = xlsxwriter.Workbook('rich_strings.xlsx')
worksheet = workbook.add_worksheet()

# Set up some formats to use.
bold = workbook.add_format({'bold': True})
italic = workbook.add_format({'italic': True})
underline = workbook.add_format({'underline': True})
strikeout = workbook.add_format({'font_strikeout': True})
wrap = workbook.add_format()
wrap.set_text_wrap()


worksheet.set_column('A:A', 30)


worksheet.write_rich_string('A1',
                            'This is ',
                            bold, 'bold',
                            ' and this is \n',
                            italic, 'italic',
                            ' and this is ',
                            underline, 'underline',
                            ' and finally, this is \n',
                            strikeout, 'strikeout.', wrap)
workbook.close()