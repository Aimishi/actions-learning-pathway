//System.NullReferenceException
/*HResult = 0x80004003
    Сообщение = Ссылка на объект не указывает на экземпляр объекта.
    Источник = DFI_0010_4.3.3.1
    Трассировка стека:
    at DFI0010_documentManipulation.Program.GetCellValue(SpreadsheetDocument document, Cell cell) in \\0400RPADB02\Blueprism_Temp_v2\syyevdo1\source\repos\DFI_0010_4.3.3.1\Program.cs:line 249
*/

string? excelZakazPath = @"\\0000UIPATHSV07\Share\RPA\Справочники\0035 процесс\СПРАВОЧНИК.xlsx";

DataSet excelDataZakaz = GetDataFromExcel(excelZakazPath);

static DataSet GetDataFromExcel(string filePathReferenceBook)
{

    var dataSet = new DataSet();

    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePathReferenceBook, false))
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
        {
            DataTable dt = new DataTable();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();

            if (rows.Count() == 0)
                continue;

            foreach (Cell cell in rows.ElementAt(0))
            {
                string columnName = GetCellValue(spreadsheetDocument, cell);

                columnName = columnName.Replace(" ", "").ToLower();

                dt.Columns.Add(columnName);
            }

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

     Regex regex = new Regex("[A-Za-z]+");
     Match match = regex.Match(cellReference);
     return match.Value;
 }

 private static int GetColumnIndexFromName(string columnName)
 {
     int columnIndex = 0;
     int factor = 1;

     for (int position = columnName.Length - 1; position >= 0; position--)
     {
         columnIndex += (columnName[position] - 'A' + 1) * factor;
         factor *= 26;
     }

     return columnIndex - 1;
 }