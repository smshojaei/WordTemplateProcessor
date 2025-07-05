using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordTemplateEngine
{
    public class Engine
    {
        public byte[] FillTemplate(byte[] templateData, Dictionary<string, string> textPlaceholders, Dictionary<string, List<Dictionary<string, string>>> tablePlaceholders)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    MainDocumentPart? mainPart = wordDoc.MainDocumentPart;
                    if (mainPart == null) return mem.ToArray(); // Or throw exception

                    Document? document = mainPart.Document;
                    if (document == null) return mem.ToArray(); // Or throw exception

                    Body? body = document.Body;
                    if (body == null) return mem.ToArray(); // Or throw exception

                    // Process text placeholders
                    if (textPlaceholders != null && textPlaceholders.Any())
                    {
                        ReplaceTextPlaceholders(body, textPlaceholders);
                    }

                    // Process table placeholders
                    if (tablePlaceholders != null && tablePlaceholders.Any())
                    {
                        ReplaceTablePlaceholders(wordDoc, tablePlaceholders);
                    }
                }
                return mem.ToArray();
            }
        }

        private void ReplaceTextPlaceholders(Body body, Dictionary<string, string> placeholders)
        {
            // Null check for placeholders already done by caller
            // Null check for body already done by caller

            // Combine all text elements to handle split tags
            string allText = string.Concat(body.Descendants<Text>().Select(t => t.Text));

            foreach (var placeholder in placeholders)
            {
                string tag = placeholder.Key;
                string value = placeholder.Value ?? string.Empty; // Ensure value is not null

                // Regex to find the tag, allowing for it to be split across runs
                // This is a simplified approach. A more robust solution might involve iterating through runs.
                string pattern = Regex.Escape(tag);

                if (allText.Contains(tag))
                {
                    // This is a complex problem. For tags split across runs,
                    // we need to identify the runs involved and replace the content.
                    // A simple string replacement on concatenated text won't work directly on the XML structure.
                    // This needs a more sophisticated approach of iterating runs, accumulating text,
                    // finding the placeholder, and then modifying the involved runs.

                    // For now, let's try a basic replacement that works for non-split tags.
                    // A more advanced implementation is needed for split tags.
                    foreach (var textElement in body.Descendants<Text>())
                    {
                        if (textElement.Text.Contains(tag))
                        {
                            textElement.Text = textElement.Text.Replace(tag, value);
                        }
                    }
                }
            }

            // Second pass to handle tags that might have been split across multiple Text elements
            // DRASTIC SIMPLIFICATION TO ISOLATE CS1513: Removing the entire loop body.
            // If this compiles, the error is within the removed code.
            // foreach (var placeholder in placeholders)
            // {
            // }
        }

        private void ReplaceTablePlaceholders(WordprocessingDocument wordDoc, Dictionary<string, List<Dictionary<string, string>>> tableData)
        {
            if (tableData == null || !tableData.Any())
                return;

            var mainDocPart = wordDoc.MainDocumentPart;
            if (mainDocPart == null) return;
            var document = mainDocPart.Document;
            if (document == null) return;
            var body = document.Body;
            if (body == null) return;

            var tables = body.Elements<Table>().ToList();
            if (!tables.Any()) return;

            foreach (var tableEntry in tableData)
            {
                string tableNameIdentifier = $"@@Table:{tableEntry.Key}@@";
                List<Dictionary<string, string>>? rowsData = tableEntry.Value;
                if (rowsData == null) continue;

                Table? targetTable = null;
                TableRow? headerRowForColumns = null;
                TableRow? templateRow = null;

                foreach (var table in tables)
                {
                    var firstRow = table.Elements<TableRow>().FirstOrDefault();
                    if (firstRow != null)
                    {
                        var firstCellText = string.Concat(firstRow.Descendants<Text>().Select(t => t.Text));
                        if (firstCellText.Contains(tableNameIdentifier))
                        {
                            targetTable = table;
                            headerRowForColumns = table.Elements<TableRow>().ElementAtOrDefault(1);
                            templateRow = table.Elements<TableRow>().ElementAtOrDefault(1);
                            if (headerRowForColumns != null)
                            {
                                var firstCellOfFirstRow = firstRow.Elements<TableCell>().FirstOrDefault();
                                if (firstCellOfFirstRow != null)
                                {
                                    foreach (var textElement in firstCellOfFirstRow.Descendants<Text>().ToList())
                                    {
                                        if (textElement.Text != null)
                                            textElement.Text = textElement.Text.Replace(tableNameIdentifier, "");
                                    }
                                }
                            }
                            break;
                        }
                    }
                }

                if (targetTable != null && templateRow != null && headerRowForColumns != null)
                {
                    var originalTemplateRowXml = templateRow.OuterXml;
                    targetTable.RemoveChild(templateRow);

                    if (rowsData.Any())
                    {
                        List<string> columnPlaceholders = headerRowForColumns.Elements<TableCell>()
                                                            .Select(cell => string.Concat(cell.Descendants<Text>().Select(t => t.Text)).Trim())
                                                            .ToList();

                        foreach (var dataRow in rowsData)
                        {
                            TableRow newRow = new TableRow(originalTemplateRowXml);

                            var cells = newRow.Elements<TableCell>().ToList();
                            for (int i = 0; i < cells.Count; i++)
                            {
                                if (i < columnPlaceholders.Count)
                                {
                                    string cellPlaceholder = columnPlaceholders[i];
                                    dataRow.TryGetValue(cellPlaceholder, out string? value);
                                    if (value == null)
                                    {
                                        dataRow.TryGetValue(cellPlaceholder.Replace("@@", ""), out value);
                                    }

                                    cells[i].RemoveAllChildren<Paragraph>();
                                    cells[i].Append(new Paragraph(new Run(new Text(value ?? ""))));
                                }
                                else
                                {
                                    cells[i].RemoveAllChildren<Paragraph>();
                                    cells[i].Append(new Paragraph(new Run(new Text(""))));
                                }
                            }
                            targetTable.Append(newRow);
                        }
                    }
                }
            }
        }
    }
}
