using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DFI0010_documentManipulation
{
    public class WordDocumentEditor
    {
        public void ReplacePlaceholdersInDocument(string sourceFilePath, Dictionary<string, string> replacements)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(sourceFilePath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                var paragraphs = body.Elements<Paragraph>();

                foreach (var para in paragraphs)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            foreach (var key in replacements.Keys)
                            {
                                if (text.Text.Contains(key))
                                {
                                    text.Text = text.Text.Replace(key, replacements[key]);

                                    Console.WriteLine($"Replaced key {key} with {replacements[key]}");
                                }
                            }
                        }
                    }
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        public bool CopyDocument(string sourceFilePath, string destinationFilePath)
        {
            try
            {
                File.Copy(sourceFilePath, destinationFilePath, true);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            string sourcePathToTemplate = @"\\10.35.42.60\Blueprism_Temp\syyevdo1\DFI\0010\4.4.14(Создание ДС в Word)\D01_D63256L26LT-3MOD — template.docx";
            string destinationPath = @"\\10.35.42.60\Blueprism_Temp\syyevdo1\DFI\0010\4.4.14(Создание ДС в Word)\D01_D63256L26LT-3MOD_copy1.docx";

            try
            {
                var editor = new WordDocumentEditor();

                bool copyFile = editor.CopyDocument(sourcePathToTemplate, destinationPath);

                if (copyFile)
                {
                    DataTable oeBSData = GetDataFromOeBS();

                    DataSet excelData = GetDataFromExcel(destinationPath);

                    Dictionary<string, string> placeholders = ConsolidateData(oeBSData, excelData);

                    editor.ReplacePlaceholdersInDocument(destinationPath, placeholders);

                    Console.WriteLine("Документ удачно модифицирован и обработан");
                }
                else
                {
                    throw new Exception("Ошибка копирования документа");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static DataTable GetDataFromOeBS()
        {
            //string zakaz = "D2305166L18-4OUT1";
            string zakaz = "D63256L26LT-3MOD";

            zakaz = zakaz.Replace(" ", "");

            string connectionString = $"User Id=APPS;Password=pOQLgtsgZVDs;Data Source=(DESCRIPTION=(ENABLE=BROKEN)(load_balance=yes)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=oebs-test-db1.nnov.inside.mts.ru)(PORT=1521))(ADDRESS=(PROTOCOL=TCP)(HOST=oebs-test-db2.nnov.inside.mts.ru)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=OEBS20)(failover_mode=(type=select)(method=basic)(retries=180)(delay=5))))";

            string request = "select * from apps.XX_RPA_HEADERS_ENTITY_CARDS_V \r\nwhere segment1 = :zakaz";

            string customerName = "ТЕЛЕКОМПЛЮС ООО";

            OracleConnection conn = new OracleConnection(connectionString);

            conn.Open();

            OracleCommand cmd = new OracleCommand(request, conn);

            //cmd.Parameters.Add(new OracleParameter("customerName", customerName));

            cmd.Parameters.Add(new OracleParameter("zakaz", zakaz));

            OracleDataAdapter adapter = new OracleDataAdapter();

            adapter.SelectCommand = cmd;

            DataTable dt = new DataTable();

            adapter.Fill(dt);

            return dt;
        }

        static DataSet GetDataFromExcel(string filePath)
        {
            string? filePathReferenceBook = @"\\10.35.42.60\Blueprism_Temp\syyevdo1\DFI\0010\4.4.3 Поиск заказа\СПРАВОЧНИКsample.xlsx";

            DataSet dataSet = new DataSet();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePathReferenceBook, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheets sheets = workbookPart.Workbook.Sheets;



                foreach (Sheet sheet in sheets)
                {
                    string sheetName = sheet.Name;
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    DataTable dataTable = new DataTable(sheetName);

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (row.RowIndex == 1)
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string columnName = string.Empty;

                                string value = cell.InnerText;

                                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                                {
                                    SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;
                                    if (sharedStringPart != null)
                                    {
                                        value = sharedStringPart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
                                    }

                                    columnName = value;
                                }

                                dataTable.Columns.Add(columnName);
                            }
                        }
                        else
                        {
                            DataRow dataRow = dataTable.NewRow();

                            int columnIndex = 0;

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = string.Empty;

                                string value = cell.InnerText;

                                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                                {
                                    SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;
                                    if (sharedStringPart != null)
                                    {
                                        value = sharedStringPart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
                                    }

                                    cellValue = value;
                                }

                                dataRow[columnIndex] = cellValue;

                                columnIndex++;
                            }

                            dataTable.Rows.Add(dataRow);
                        }
                    }

                    dataSet.Tables.Add(dataTable);
                }


                /*DataTable firstSheet = dataSet.Tables[0];

                foreach (DataTable dataTable in dataSet.Tables)
                {

                }*/
            }


            return dataSet;
        }

        static Dictionary<string, string> ConsolidateData(DataTable oeBSData, DataSet excelData)
        {
            var consolidatedData = new Dictionary<string, string>();


            DataTable Подписанты = excelData.Tables["Подписанты"];
            DataTable Поставщики = excelData.Tables["Поставщики"];

            if (oeBSData.Rows.Count > 0)
            {
                DataRow headerRow = oeBSData.Rows[0];


                string oeName = headerRow["Name"].ToString();
                string vendorName = headerRow["Vendor_name"].ToString();
                string vendorCode = headerRow["Vendor_site_code"].ToString();

                //Блок формирования номера ДС

                string orderNumber = headerRow["Segment1"].ToString();

                string newOrderNumber = "D01/" + orderNumber;

                if (newOrderNumber.Length > 20)
                {
                    newOrderNumber = newOrderNumber.Substring(0, 20);
                }

                consolidatedData["DS"] = newOrderNumber;

                consolidatedData["zakazNum"] = headerRow["Segment1"].ToString();

                DateTime result_date_fulfilled = DateTime.Now;
           
                bool resultTryParse = DateTime.TryParse(headerRow["date_fulfilled"].ToString(), out result_date_fulfilled);

                if (resultTryParse) 
                {
                    consolidatedData["zakazDate"] = result_date_fulfilled.ToString("d");
                }
                else 
                {
                    consolidatedData["zakazDate"] = headerRow["date_fulfilled"].ToString();
                }

                consolidatedData["contractNum"] = headerRow["contract_num"].ToString();
                consolidatedData["contractDate"] = "01.02.2024"; //заменить получением из вью


                var подписантыRow = Подписанты.Select($"ОЕ = '{oeName}'").FirstOrDefault();
                if (подписантыRow != null)
                {
                    consolidatedData["ozdCity"] = "г." + подписантыRow["Город"].ToString();

                    //Блок формирования полного имени 
                    string podpisant = подписантыRow["Подписант"].ToString();

                    string io = подписантыRow["ИО"].ToString();

                    string surname = podpisant.Split(' ')[0];

                    string fullName = surname + " " + io;

                    //Конец блока формирования полного имени

                    consolidatedData["ozdOwnerFull"] = подписантыRow["Должность"].ToString() + " " + fullName;
                    consolidatedData["ozdOwnerShort"] = подписантыRow["Подписант"].ToString(); // заменить получением из вью или эксель
                    consolidatedData["notaryNum"] = подписантыRow["Доверенность"].ToString();

                    DateTime result_notaryDate = DateTime.Now;

                    bool resultTryParse2 = DateTime.TryParse(подписантыRow["Дата"].ToString(), out result_notaryDate);

                    if (resultTryParse2) 
                    {
                        consolidatedData["notaryDate"] = result_notaryDate.ToString("d");
                    }
                    else 
                    {
                        consolidatedData["notaryDate"] = подписантыRow["Дата"].ToString();
                    }

                }

                var поставщикиRow = Поставщики.Select($"ОЕ = '{oeName}' AND Поставщик = '{vendorName}'").FirstOrDefault();
                if (поставщикиRow != null)
                {
                    consolidatedData["KaFullName"] = поставщикиRow["Полное название"].ToString() + " " + поставщикиRow["Название"].ToString();

                    consolidatedData["KaShortName"] = поставщикиRow["Название"].ToString();

                    consolidatedData["KaCEOFull"] = поставщикиRow?["Должность подписанта скл"].ToString() + " " + поставщикиRow?["Подписант скл"].ToString();

                    consolidatedData["KaCEOshort"] = поставщикиRow?["Подписант"].ToString();

                    consolidatedData["Foundation"] = поставщикиRow?["Доверенность скл"].ToString();
                    
                }
            }

            return consolidatedData;
        }
    }
}
