using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text; // Added for StringBuilder

namespace WordTemplateEngine
{
    public class Engine
    {
        public byte[] FillTemplate(byte[] templateData, Dictionary<string, string>? textPlaceholders, Dictionary<string, List<Dictionary<string, string>>>? tablePlaceholders)
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
            if (placeholders == null || !placeholders.Any()) return;

            List<Paragraph> paragraphs = body.Descendants<Paragraph>().ToList();

            foreach (Paragraph para in paragraphs)
            {
                // Loop continuously over the same paragraph until no more replacements can be made in it.
                bool madeReplacementInParagraphPass;
                do
                {
                    madeReplacementInParagraphPass = false;
                    List<Text> textElements = para.Descendants<Text>().ToList();
                    if (!textElements.Any()) break;

                    // Build current view of paragraph text and map characters to their source Text elements
                    StringBuilder currentParagraphText = new StringBuilder();
                    // Stores: Text Element, Start Index in paragraphString, End Index in paragraphString
                    List<Tuple<Text, int, int>> elementMap = new List<Tuple<Text, int, int>>();
                    int charIndex = 0;
                    foreach (Text txtEl in textElements)
                    {
                        if (txtEl.Text != null) // Consider empty strings as they might be placeholders for future text
                        {
                            currentParagraphText.Append(txtEl.Text);
                            elementMap.Add(Tuple.Create(txtEl, charIndex, charIndex + txtEl.Text.Length -1));
                            charIndex += txtEl.Text.Length;
                        }
                    }
                    string paragraphString = currentParagraphText.ToString();
                    if (string.IsNullOrEmpty(paragraphString)) continue; // Nothing to process

                    foreach (var placeholderEntry in placeholders)
                    {
                        string tagToFind = $"@@{placeholderEntry.Key}@@";
                        string replacementString = placeholderEntry.Value ?? string.Empty;

                        int tagLocationInParagraph = paragraphString.IndexOf(tagToFind, StringComparison.Ordinal);
                        if (tagLocationInParagraph == -1) continue;

                        // Found a tag. Identify all Text elements that are part of this specific tag occurrence.
                        int tagEndLocationInParagraph = tagLocationInParagraph + tagToFind.Length - 1;

                        Text firstElementInTag = null;
                        int tagStartIndexInFirstElement = -1; // Relative index of tag's start within firstElementInTag.Text

                        Text lastElementInTag = null;
                        int tagEndIndexInLastElement = -1; // Relative index of tag's end within lastElementInTag.Text

                        List<Text> intermediateElementsToClear = new List<Text>();

                        for(int mapIdx = 0; mapIdx < elementMap.Count; mapIdx++)
                        {
                            var mapEntry = elementMap[mapIdx];
                            Text currentTextEl = mapEntry.Item1;
                            int elStartInParagraph = mapEntry.Item2;
                            int elEndInParagraph = mapEntry.Item3; // Inclusive end index

                            // Check for overlap: max(start1, start2) <= min(end1, end2)
                            bool overlaps = Math.Max(tagLocationInParagraph, elStartInParagraph) <= Math.Min(tagEndLocationInParagraph, elEndInParagraph);

                            if (overlaps)
                            {
                                if (firstElementInTag == null) // This is the first Text element hit by the tag
                                {
                                    firstElementInTag = currentTextEl;
                                    tagStartIndexInFirstElement = tagLocationInParagraph - elStartInParagraph;
                                }

                                // This element will be the last one if its end covers or goes beyond the tag's end
                                // or if it's simply the last one in a sequence of overlapping elements.
                                lastElementInTag = currentTextEl;
                                tagEndIndexInLastElement = tagEndLocationInParagraph - elStartInParagraph;

                                // If not the first and (potentially) not the last, it's intermediate.
                                // This logic will be refined after identifying first and last.
                            }
                        }

                        if (firstElementInTag != null) // Should always be true if tagLocationInParagraph != -1
                        {
                            // Collect intermediate elements AFTER first and last are definitively known
                            bool withinTagSpan = false;
                            foreach (var mapEntry in elementMap)
                            {
                                Text currentTextEl = mapEntry.Item1;
                                if (currentTextEl == firstElementInTag) withinTagSpan = true;

                                if (withinTagSpan && currentTextEl != firstElementInTag && currentTextEl != lastElementInTag)
                                {
                                    intermediateElementsToClear.Add(currentTextEl);
                                }

                                if (currentTextEl == lastElementInTag) withinTagSpan = false;
                            }


                            if (firstElementInTag == lastElementInTag)
                            {
                                // Tag is contained within a single Text element
                                string currentText = firstElementInTag.Text ?? string.Empty;
                                string prefix = currentText.Substring(0, Math.Min(currentText.Length, tagStartIndexInFirstElement));
                                // Suffix starts after the tag ends in this element.
                                // tagStartIndexInFirstElement + tagToFind.Length gives the point *after* the tag in the element's text.
                                int tagActualEndIndexInElement = tagStartIndexInFirstElement + tagToFind.Length;
                                string suffix = currentText.Substring(Math.Min(tagActualEndIndexInElement, currentText.Length));
                                firstElementInTag.Text = prefix + replacementString + suffix;
                            }
                            else
                            {
                                // Tag spans multiple Text elements
                                // Handle first element (contains beginning of tag)
                                string firstText = firstElementInTag.Text ?? string.Empty;
                                firstElementInTag.Text = firstText.Substring(0, Math.Min(firstText.Length,tagStartIndexInFirstElement)) + replacementString;

                                // Handle last element (contains end of tag)
                                string lastText = lastElementInTag.Text ?? string.Empty;
                                // tagEndIndexInLastElement is inclusive. Substring starts at char *after* this index.
                                int lastElementContentStart = Math.Min(lastText.Length, tagEndIndexInLastElement + 1);
                                lastElementInTag.Text = lastText.Substring(lastElementContentStart);

                                // Clear text in intermediate elements
                                foreach (Text intermediateElement in intermediateElementsToClear)
                                {
                                    intermediateElement.Text = string.Empty;
                                }
                            }
                            madeReplacementInParagraphPass = true;
                            // A replacement was made, so break from iterating placeholders
                            // and restart the do-while loop for this paragraph to rebuild text map and re-scan.
                            break;
                        }
                    } // End foreach placeholderEntry
                } while (madeReplacementInParagraphPass); // Loop while replacements are made in this paragraph
            } // End foreach para
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
                                    string columnHeaderTag = columnPlaceholders[i]; // This is "@@Key@@"
                                    string simpleKey = columnHeaderTag.Replace("@@", "");

                                    dataRow.TryGetValue(simpleKey, out string? value);

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
