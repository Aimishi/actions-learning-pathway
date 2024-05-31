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
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using LTools.Scripting.Model;
using LTools.Network.Model;

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

            foreach (var table in body.Elements<Table>())
            {
                foreach (var row in table.Elements<TableRow>())
                {
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        foreach (var para in cell.Elements<Paragraph>())
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

public class PrimoScript
{
	public static LTools.Scripting.CSharp.ScriptDebugger __debug;
	
	public void main(LTools.Common.Model.WorkflowData wf)
    {
		//Получить значение входных аргументов
		
		//Настраиваем строку подключения
		//string connectionString = wf.GetArgument("in_strConnectionString").ToString();
		string connectionString = $"User Id=APPS;Password=pOQLgtsgZVDs;Data Source=(DESCRIPTION=(ENABLE=BROKEN)(load_balance=yes)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=oebs-test-db1.nnov.inside.mts.ru)(PORT=1521))(ADDRESS=(PROTOCOL=TCP)(HOST=oebs-test-db2.nnov.inside.mts.ru)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=OEBS20)(failover_mode=(type=select)(method=basic)(retries=180)(delay=5))))";
	
		
		//string user = wf.GetArgument("in_strUser");

		//string pass = wf.GetArgument("in_strPass");
		
		//connectionString = connectionString.Replace("%user%", user).Replace("%pass%", pass);
		
		//Текст запроса
		string request = "select * from apps.XX_RPA_HEADERS_ENTITY_CARDS_V where segment1 = :zakaz";
		//string? request = wf.GetArgument("in_requestSQL").ToString();
		
		//string? zakaz = wf.GetArgument("in_strZakazNum").ToString();
		string zakaz = "D2305166L18-4OUT1";
		
		//string downloadPath = wf.GetArgument("in_transactionWorkFolder").ToString();
		string downloadPath = @"C:\Distr\DFI\DFI.RPA.0010\Data\Temp\D06-D230161206_29-05-2024_copy3";
		
		string sourcePathToTemplate = @"\\10.35.42.60\Blueprism_Temp\syyevdo1\DFI\0010\4.4.14(Создание ДС в Word)\DS_template.docx";
		//string sourcePathToTemplate = wf.GetArgument("in_pathTemplateDS").ToString();
		
		//string destinationPath = @"\\10.35.42.60\Blueprism_Temp\syyevdo1\DFI\0010\4.4.14(Создание ДС в Word)\D01_copy1.docx";
		string destinationPath = Path.Combine(downloadPath, $"D01_{zakaz}.docx");
		
		//string? filePathReferenceBook = wf.GetArgument("in_pathToReferenceBook").ToString();
		string? filePathReferenceBook = @"\\0000UIPATHSV07\Share\RPA\Справочники\0035 процесс\СПРАВОЧНИК.xlsx";
		
		try
		{
			//СТАРТ блока формирования ДС
		    var editor = new WordDocumentEditor();

			DataTable oeBSData = new DataTable();
			
			DataSet excelData = new DataSet();
			
			Dictionary<string, string> placeholders = new Dictionary<string, string>();
			
		    bool copyFile = editor.CopyDocument(sourcePathToTemplate, destinationPath);
		
		    if (copyFile)
		    {
		        oeBSData = GetDataFromOeBS(connectionString, request, zakaz);
				
				wf.SetArgument("out_dtZakazLines", oeBSData);
		
		        excelData = GetDataFromExcel(filePathReferenceBook);
				
				wf.SetArgument("out_setReferenceBook", excelData);
		
		        placeholders = ConsolidateData(oeBSData, excelData);
		
		        editor.ReplacePlaceholdersInDocument(destinationPath, placeholders);
				
				wf.SetArgument("out_DSnum", placeholders["newOrderNumber"]);
				
				wf.SetArgument("out_DSpath", destinationPath);
		
		        Console.WriteLine($"Документ ДС удачно модифицирован и создан: {destinationPath}");
		    }
		    else
		    {
		        throw new Exception("Ошибка копирования документа шаблона: {sourcePathToTemplate}");
		    }
			//КОНЕЦ блока формирования ДС
			
			//СТАРТ блока получения значений из таблицы данных заказа в ОеБС
			
			DataRow result = oeBSData.AsEnumerable()
			    .FirstOrDefault(row => row.Field<string>("ACTIVE_FLAG") == "Y");
			
			string? link = result != null ? result.Field<string>("LINK_FILENET").ToString() : null;
			
			string? customerName = result != null ? result.Field<string>("VENDOR_NAME").ToString() : null;
			
			string? operatingUnit = result != null ? result.Field<decimal>("ORG_ID").ToString() : null;
			
			string? contractID = result != null ? result.Field<decimal>("CONTRACT_ID").ToString() : null;
			
			string? startDate = result != null ? result.Field<DateTime>("DATE_FULFILLED").ToString("yyyy/MM/dd") : null;
			
			wf.SetArgument("out_strLinkUri", link);
			
			wf.SetArgument("out_VENDOR_NAME", customerName);
			
			wf.SetArgument("out_operatingUnit", operatingUnit);
			
			wf.SetArgument("out_contractID", contractID);		
			
			wf.SetArgument("out_startDate", startDate);
			
			//КОНЕЦ блока получения значений из таблицы данных заказа в ОеБС
			
			//СТАРТ блока получения файла контракта из UCM
			
			//string downloadPath = System.IO.Path.GetFullPath(".\\Data\\Temp");
			
			
			
			var options = new ChromeOptions();
			
			options.AddUserProfilePreference("download.default_directory", downloadPath);
			options.AddUserProfilePreference("download.prompt_for_download", false);
			options.AddUserProfilePreference("download.directory_upgrade", true);
			options.AddUserProfilePreference("safebrowsing.enabled", true);
			
			using (var driver = new ChromeDriver(@"C:\Distr\processesConfig\DFI.0008\Chromedriver\chromedriver_98.0.4758.102.exe", options))
			{
				driver.Navigate().GoToUrl(link);
				
				System.Threading.Thread.Sleep(10000);

			}
			
			string[] filesPdf = Directory.GetFiles(downloadPath, "*.pdf");
			
			if(filesPdf.Length == 0)
			{
				throw new Exception("Не удалось скачать файл заказа. Ошибка сервера UCM.");
			}
			
			wf.SetArgument("out_pathPdfFile", filesPdf[0].ToString());
			
			//КОНЕЦ блока получения файла контракта из UCM
			
		}
		catch (Exception ex)
		{
		    Console.WriteLine(ex.Message);
			
			wf.SetArgument("out_str_exception", ex.Message);
		}
        
    }
	
	static DataTable GetDataFromOeBS(string connectionString, string sqlRequest, string zakaz)
	{
	    //string zakaz = "D2305166L18-4OUT1";
	    //string zakaz = "D63256L26LT-3MOD";
	
	    zakaz = zakaz.Replace(" ", "");
	
	    //string connectionString = $"User Id=APPS;Password=;Data Source=(DESCRIPTION=(ENABLE=BROKEN)(load_balance=yes)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=oebs-test-db1.nnov.inside.mts.ru)(PORT=1521))(ADDRESS=(PROTOCOL=TCP)(HOST=oebs-test-db2.nnov.inside.mts.ru)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=OEBS20)(failover_mode=(type=select)(method=basic)(retries=180)(delay=5))))";
	
	    //string request = "select * from apps.XX_RPA_HEADERS_ENTITY_CARDS_V \r\nwhere segment1 = :zakaz";
	
	    //string customerName = "ТЕЛЕКОМПЛЮС ООО";
	
	    OracleConnection conn = new OracleConnection(connectionString);
	
	    conn.Open();
	
	    OracleCommand cmd = new OracleCommand(sqlRequest, conn);
	
	    //cmd.Parameters.Add(new OracleParameter("customerName", customerName));
	
	    cmd.Parameters.Add(new OracleParameter("zakaz", zakaz));
	
	    OracleDataAdapter adapter = new OracleDataAdapter();
	
	    adapter.SelectCommand = cmd;
	
	    DataTable dt = new DataTable();
	
	    adapter.Fill(dt);
	
	    return dt;
	}
	
	static DataSet GetDataFromExcel(string filePathReferenceBook)
	{
	    //string? filePathReferenceBook = @"\\10.35.42.60\Blueprism_Temp\syyevdo1\DFI\0010\4.4.3 Поиск заказа\СПРАВОЧНИКsample.xlsx";
	    //string? filePathReferenceBook = @"\\0000UIPATHSV07\Share\RPA\Справочники\0035 процесс\СПРАВОЧНИК.xlsx";
	
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
