using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public void ReplacePlaceholdersInDocument(string sourceFilePath, Dictionary<string, string> replacements)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(sourceFilePath, true))
    {
        var body = wordDoc.MainDocumentPart.Document.Body;
        ReplacePlaceholdersInBody(body, replacements);
        wordDoc.MainDocumentPart.Document.Save();
    }
}

private void ReplacePlaceholdersInBody(Body body, Dictionary<string, string> replacements)
{
    foreach (var element in body.Elements())
    {
        if (element is Paragraph)
        {
            ReplacePlaceholdersInParagraph(element as Paragraph, replacements);
        }
        else if (element is Table)
        {
            ReplacePlaceholdersInTable(element as Table, replacements);
        }
    }
}

private void ReplacePlaceholdersInParagraph(Paragraph paragraph, Dictionary<string, string> replacements)
{
    foreach (var run in paragraph.Elements<Run>())
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

private void ReplacePlaceholdersInTable(Table table, Dictionary<string, string> replacements)
{
    foreach (var row in table.Elements<TableRow>())
    {
        foreach (var cell in row.Elements<TableCell>())
        {
            foreach (var element in cell.Elements())
            {
                if (element is Paragraph)
                {
                    ReplacePlaceholdersInParagraph(element as Paragraph, replacements);
                }
                else if (element is Table)
                {
                    ReplacePlaceholdersInTable(element as Table, replacements);
                }
            }
        }
    }
}
