using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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

                                    Console.WriteLine($"Заменен ключ {key} на {replacements[key]}");
                                }
                            }
                        }
                    }
                }

                foreach (var replacement in replacements)
                {
                    foreach (var text in body.Descendants<Text>()) 
                    {
                        if (text.Text.Contains(replacement.Key)) 
                        {
                            text.Text = text.Text.Replace(replacement.Key, replacement.Value); 
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
            catch(Exception ex) 
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
                    var placeholders = new Dictionary<string, string>
                    {
                        {"DS", "12345"},
                        {"zakazNum", "789"},
                        {"zakazDate", "01.01.2024"},
                        {"contractNum", "456"},
                        {"contractDate", "01.02.2024"},
                        {"ozdCity", "г.Москва"},
                        {"ozdOwnerFull", "Иванов Иван Иванович"},
                        {"ozdOwnerShort", "Иванов И.И."},
                        {"notaryNum", "N-001"},
                        {"notaryDate", "12.12.2023"},
                        {"KaFullName", "OOO 'Рога и копыта'"},
                        {"KaShortName", "OOO 'Рога и копыта'"},
                        {"KaCEOFull", "Петров Петр Петрович"},
                        {"KaCEOshort", "Петров П.П."}
                    };

                    editor.ReplacePlaceholdersInDocument(destinationPath, placeholders);

                    Console.WriteLine("Document has been successfully modified and saved.");
                }
                else
                {
                    throw new Exception();
                }

            }
            catch(Exception ex)
            { 
                Console.WriteLine(ex.Message);
            }

            
        }
    }
}
