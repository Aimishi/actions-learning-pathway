using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

                // Fetch data from OeBS and Excel
                DataTable oeBSData = GetDataFromOeBS();
                DataSet excelData = GetDataFromExcel(@"path_to_excel_file");

                // Prepare dictionary with consolidated data
                Dictionary<string, string> placeholders = ConsolidateData(oeBSData, excelData);

                // Copy and edit the document
                bool copyFile = editor.CopyDocument(sourcePathToTemplate, destinationPath);

                if (copyFile)
                {
                    editor.ReplacePlaceholdersInDocument(destinationPath, placeholders);
                    Console.WriteLine("Document has been successfully modified and saved.");
                }
                else
                {
                    throw new Exception("Failed to copy the document.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static DataTable GetDataFromOeBS()
        {
            // Your implementation to get data from OeBS database
            // Return DataTable
            return new DataTable(); // Placeholder
        }

        static DataSet GetDataFromExcel(string filePath)
        {
            // Your implementation to get data from Excel file
            // Return DataSet
            return new DataSet(); // Placeholder
        }

        static Dictionary<string, string> ConsolidateData(DataTable oeBSData, DataSet excelData)
        {
            var consolidatedData = new Dictionary<string, string>();

            // Extract tables from DataSet
            DataTable Подписанты = excelData.Tables["Подписанты"];
            DataTable Поставщики = excelData.Tables["Поставщики"];

            if (oeBSData.Rows.Count > 0)
            {
                DataRow headerRow = oeBSData.Rows[0];

                // Assuming only one row per 'Заголовки' for simplicity
                string oeName = headerRow["Name"].ToString();
                string vendorName = headerRow["Vendor_name"].ToString();
                string vendorCode = headerRow["Vendor_site_code"].ToString();

                consolidatedData["DS"] = "12345";
                consolidatedData["zakazNum"] = headerRow["Segment1"].ToString();
                consolidatedData["zakazDate"] = headerRow["date_fulfilled"].ToString();
                consolidatedData["contractNum"] = headerRow["contract_num"].ToString();
                consolidatedData["contractDate"] = "01.02.2024"; // Static as per your example

                // Find Подписанты row matching Заголовки.Name
                var подписантыRow = Подписанты.Select($"ОЕ = '{oeName}'").FirstOrDefault();
                if (подписантыRow != null)
                {
                    consolidatedData["ozdCity"] = подписантыRow["Город"].ToString();
                    consolidatedData["ozdOwnerFull"] = подписантыRow["Подписант"].ToString() + " " + подписантыRow["Должность"].ToString();
                    consolidatedData["ozdOwnerShort"] = "Иванов И.И."; // Static as per your example
                    consolidatedData["notaryNum"] = подписантыRow["Доверенность"].ToString();
                    consolidatedData["notaryDate"] = подписантыRow["Дата"].ToString();
                }

                // Find Поставщики row matching Заголовки.Name, Vendor_name and Vendor_site_code
                var поставщикиRow = Поставщики.Select($"ОЕ = '{oeName}' AND Поставщик = '{vendorName}' AND Отделение = '{vendorCode}'").FirstOrDefault();
                if (поставщикиRow != null)
                {
                    consolidatedData["KaFullName"] = поставщикиRow["Полное название"].ToString() + " " + поставщикиRow["Название"].ToString();
                    consolidatedData["KaShortName"] = поставщикиRow["Название"].ToString();
                    consolidatedData["KaCEOFull"] = поставщикиRow["Поставщик"].ToString();
                    consolidatedData["KaCEOshort"] = подписантыRow?["Подписант"].ToString();
                }
            }

            return consolidatedData;
        }
    }
}
