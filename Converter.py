def XLSXConv(ExcelPath,TextPath):

    import xlrd
    wb = xlrd.open_workbook(loc) # Access work book
    Sheet = wb.sheet_by_index(0) # Specify worksheet

    No_rows = Sheet.nrows # Number of rows in sheet
    No_columns = Sheet.ncols # Number of columns in sheet

    f = open(TextPath,"w+") # Creates txt file and overwrites content, if file already exists

    for Position_row in range(No_rows):
        Row = []

        for Position_col in range(No_columns):
            Row.append(Sheet.cell_value(Position_row,Position_col)) # Get each row

        f.write('\t \t'.join(map(str, Row)))

    f.close()


