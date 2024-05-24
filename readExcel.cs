public static DataSet ReadExcelFile(string filePath)
{
    var dataSet = new DataSet();

    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
        {
            DataTable dt = new DataTable();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();

            if (rows.Count() == 0)
                continue;

            // Add columns to DataTable
            foreach (Cell cell in rows.ElementAt(0))
            {
                dt.Columns.Add(GetCellValue(spreadsheetDocument, cell));
            }

            // Add rows to DataTable
            foreach (Row row in rows.Skip(1))
            {
                DataRow tempRow = dt.NewRow();
                int columnIndex = 0;
                foreach (Cell cell in row.Descendants<Cell>())
                {
                    int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                    if (columnIndex < cellColumnIndex)
                    {
                        do
                        {
                            tempRow[columnIndex] = string.Empty;
                            columnIndex++;
                        }
                        while (columnIndex < cellColumnIndex);
                    }
                    tempRow[columnIndex] = GetCellValue(spreadsheetDocument, cell);
                    columnIndex++;
                }

                dt.Rows.Add(tempRow);
            }

            dt.TableName = workbookPart.Workbook.Descendants<Sheet>().First(s => workbookPart.GetIdOfPart(worksheetPart) == s.Id).Name;
            dataSet.Tables.Add(dt);
        }
    }

    return dataSet;
}

private static string GetCellValue(SpreadsheetDocument document, Cell cell)
{
    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
    string value = cell.CellValue.InnerXml;

    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    {
        return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
    }
    else
    {
        return value;
    }
}

private static string GetColumnName(string cellReference)
{
    // Match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellReference);
    return match.Value;
}

private static int GetColumnIndexFromName(string columnName)
{
    int columnIndex = 0;
    int factor = 1;

    // Calculate column index based on column name
    for (int position = columnName.Length - 1; position >= 0; position--)
    {
        columnIndex += (columnName[position] - 'A' + 1) * factor;
        factor *= 26;
    }

    return columnIndex - 1;
}
