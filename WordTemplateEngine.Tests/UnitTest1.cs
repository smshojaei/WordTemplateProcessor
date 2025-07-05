using Xunit;
using WordTemplateEngine;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Text;

namespace WordTemplateEngine.Tests
{
    public class UnitTest1
    {
        private Engine wordEngine = new Engine();

        private byte[] CreateSimpleDocWithText(string textContent)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text(textContent));
                    mainPart.Document.Save();
                }
                return mem.ToArray();
            }
        }

        private byte[] CreateDocWithSplitText(string part1, string part2, string part3)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    para.AppendChild(new Run(new Text(part1)));
                    para.AppendChild(new Run(new Text(part2)));
                    para.AppendChild(new Run(new Text(part3)));
                    mainPart.Document.Save();
                }
                return mem.ToArray();
            }
        }

        private string GetDocumentText(byte[] docData)
        {
            using (MemoryStream mem = new MemoryStream(docData))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, false))
            {
                return string.Concat(wordDoc.MainDocumentPart.Document.Body.Descendants<Text>().Select(t => t.Text));
            }
        }

        [Fact]
        public void FillTemplate_SimpleTextReplacement_ReplacesTag()
        {
            byte[] template = CreateSimpleDocWithText("Hello @@Name@@!");
            var textPlaceholders = new Dictionary<string, string> { { "@@Name@@", "World" } };

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);

            Assert.Contains("Hello World!", GetDocumentText(result));
            Assert.DoesNotContain("@@Name@@", GetDocumentText(result));
        }

        [Fact]
        public void FillTemplate_MultipleTextReplacements_ReplacesAllTags()
        {
            byte[] template = CreateSimpleDocWithText("@@Greeting@@, @@User@@. Today is @@Date@@.");
            var textPlaceholders = new Dictionary<string, string>
            {
                { "@@Greeting@@", "Hi" },
                { "@@User@@", "Jules" },
                { "@@Date@@", "2024-07-05" }
            };

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);

            Assert.Contains("Hi, Jules. Today is 2024-07-05.", GetDocumentText(result));
        }

        [Fact]
        public void FillTemplate_TagNotFound_LeavesTextAsIs()
        {
            byte[] template = CreateSimpleDocWithText("Hello @@Name@@!");
            var textPlaceholders = new Dictionary<string, string> { { "@@NonExistentTag@@", "World" } };

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);

            Assert.Contains("Hello @@Name@@!", GetDocumentText(result));
        }

        [Fact]
        public void FillTemplate_EmptyValue_ReplacesTagWithEmptyString()
        {
            byte[] template = CreateSimpleDocWithText("Hello @@Name@@!");
            var textPlaceholders = new Dictionary<string, string> { { "@@Name@@", "" } };

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);

            Assert.Contains("Hello !", GetDocumentText(result)); // Note the space remains
        }

        [Fact]
        public void FillTemplate_NullValue_ReplacesTagWithEmptyString()
        {
            byte[] template = CreateSimpleDocWithText("Hello @@Name@@!");
            var textPlaceholders = new Dictionary<string, string> { { "@@Name@@", null } };

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);

            Assert.Contains("Hello !", GetDocumentText(result));
        }

        [Fact]
        public void FillTemplate_UnicodeCharacters_ReplacesCorrectly()
        {
            byte[] template = CreateSimpleDocWithText("نام: @@FullName@@"); // Persian: "Name: @@FullName@@"
            var textPlaceholders = new Dictionary<string, string> { { "@@FullName@@", "جولیا" } }; // Persian: "Julia"

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);

            Assert.Contains("نام: جولیا", GetDocumentText(result));
        }

        [Fact]
        public void FillTemplate_TagSplitAcrossRuns_ReplacesTag()
        {
            // Simulating a tag split like "@@Na" + "me@@"
            byte[] template = CreateDocWithSplitText("Hello @@Na", "me@@", "!");
            var textPlaceholders = new Dictionary<string, string> { { "@@Name@@", "SplitWorld" } };

            byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);
            string resultText = GetDocumentText(result);

            Assert.Contains("Hello SplitWorld!", resultText);
            Assert.DoesNotContain("@@Name@@", resultText);
            Assert.DoesNotContain("@@Na", resultText);
            Assert.DoesNotContain("me@@", resultText);
        }

        [Fact]
        public void FillTemplate_TagSplitAcrossMoreRuns_ReplacesTag()
        {
            // Simulating a tag split like "@@N" + "a" + "m" + "e@@"
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    para.AppendChild(new Run(new Text("Data: @@")));
                    para.AppendChild(new Run(new Text("Val")));
                    para.AppendChild(new Run(new Text("ue@@")));
                    mainPart.Document.Save();
                    byte[] template = mem.ToArray();

                    var textPlaceholders = new Dictionary<string, string> { { "@@Value@@", "Correct" } };
                    byte[] result = wordEngine.FillTemplate(template, textPlaceholders, null);
                    string resultText = GetDocumentText(result);

                    Assert.Contains("Data: Correct", resultText);
                    Assert.DoesNotContain("@@Value@@", resultText);
                }
            }
        }

        private byte[] CreateDocWithTableTemplate(string tableIdentifier, List<string> columnPlaceholders)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Add a paragraph before the table
                    Paragraph paraBefore = body.AppendChild(new Paragraph());
                    Run runBefore = paraBefore.AppendChild(new Run());
                    runBefore.AppendChild(new Text("Document with a table:"));

                    Table table = new Table();

                    // Table properties (optional, for borders etc.)
                    TableProperties tblProps = new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new BottomBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new LeftBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new RightBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new InsideHorizontalBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new InsideVerticalBorder { Val = new DocumentFormat.OpenXml.EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                        )
                    );
                    table.AppendChild(tblProps);

                    // First row: Table Identifier
                    TableRow row1 = new TableRow();
                    TableCell cell1_1 = new TableCell(new Paragraph(new Run(new Text(tableIdentifier))));
                    row1.Append(cell1_1);
                    // Add empty cells if columnPlaceholders define more than one column, to make it look like a merged cell
                    for(int i = 1; i < columnPlaceholders.Count; i++)
                    {
                        row1.Append(new TableCell(new Paragraph(new Run(new Text(""))))); // Empty cell
                    }
                    table.Append(row1);

                    // Second row: Column Placeholders (Template Row)
                    TableRow row2 = new TableRow();
                    foreach (var placeholder in columnPlaceholders)
                    {
                        TableCell cell = new TableCell(new Paragraph(new Run(new Text(placeholder))));
                        row2.Append(cell);
                    }
                    table.Append(row2);

                    body.Append(table);

                    // Add a paragraph after the table
                    Paragraph paraAfter = body.AppendChild(new Paragraph());
                    Run runAfter = paraAfter.AppendChild(new Run());
                    runAfter.AppendChild(new Text("End of table section."));

                    mainPart.Document.Save();
                }
                return mem.ToArray();
            }
        }

        private List<List<string>> GetTableData(byte[] docData, int expectedTableIndex = 0)
        {
            var tableData = new List<List<string>>();
            using (MemoryStream mem = new MemoryStream(docData))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, false))
            {
                Table table = wordDoc.MainDocumentPart.Document.Body.Elements<Table>().ElementAtOrDefault(expectedTableIndex);
                if (table != null)
                {
                    foreach (var row in table.Elements<TableRow>())
                    {
                        var rowData = new List<string>();
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            rowData.Add(string.Concat(cell.Descendants<Text>().Select(t => t.Text)));
                        }
                        tableData.Add(rowData);
                    }
                }
            }
            return tableData;
        }

        [Fact]
        public void FillTemplate_SimpleTablePopulation_PopulatesRows()
        {
            string tableId = "Employees";
            byte[] template = CreateDocWithTableTemplate($"@@Table:{tableId}@@", new List<string> { "@@Name@@", "@@Age@@" });

            var tableData = new Dictionary<string, List<Dictionary<string, string>>>
            {
                {
                    tableId,
                    new List<Dictionary<string, string>>
                    {
                        new Dictionary<string, string> { { "@@Name@@", "Alice" }, { "@@Age@@", "30" } },
                        new Dictionary<string, string> { { "@@Name@@", "Bob" }, { "@@Age@@", "25" } }
                    }
                }
            };

            byte[] result = wordEngine.FillTemplate(template, null, tableData);
            List<List<string>> resultTable = GetTableData(result);

            // Expected: Row 1 (Identifier cleared), Row 2 (Alice, 30), Row 3 (Bob, 25)
            // The original template row with @@Name@@, @@Age@@ is removed.
            // The first row which had @@Table:Employees@@ should now be empty or just have residual formatting.
            // The engine implementation clears the @@Table:ID@@ tag.

            Assert.Equal(3, resultTable.Count); // Identifier row + 2 data rows
            Assert.Equal("", resultTable[0][0].Trim()); // Identifier cell cleared

            Assert.Equal("Alice", resultTable[1][0]);
            Assert.Equal("30", resultTable[1][1]);
            Assert.Equal("Bob", resultTable[2][0]);
            Assert.Equal("25", resultTable[2][1]);
        }

        [Fact]
        public void FillTemplate_TableDataWithMissingColumns_PopulatesAvailableAndClearsMissing()
        {
            string tableId = "Products";
            byte[] template = CreateDocWithTableTemplate($"@@Table:{tableId}@@", new List<string> { "@@ID@@", "@@Name@@", "@@Price@@" });

            var tableData = new Dictionary<string, List<Dictionary<string, string>>>
            {
                {
                    tableId,
                    new List<Dictionary<string, string>>
                    {
                        new Dictionary<string, string> { { "@@ID@@", "1" }, { "@@Name@@", "Laptop" } }, // Price missing
                        new Dictionary<string, string> { { "@@Name@@", "Mouse" }, { "@@Price@@", "20" } } // ID missing
                    }
                }
            };

            byte[] result = wordEngine.FillTemplate(template, null, tableData);
            List<List<string>> resultTable = GetTableData(result);

            Assert.Equal(3, resultTable.Count); // Identifier row + 2 data rows
            // Row 1 (Data)
            Assert.Equal("1", resultTable[1][0]);
            Assert.Equal("Laptop", resultTable[1][1]);
            Assert.Equal("", resultTable[1][2]); // Missing Price should be empty
            // Row 2 (Data)
            Assert.Equal("", resultTable[2][0]); // Missing ID should be empty
            Assert.Equal("Mouse", resultTable[2][1]);
            Assert.Equal("20", resultTable[2][2]);
        }

        [Fact]
        public void FillTemplate_NoDataForTable_KeepsTableStructureEmpty()
        {
            string tableId = "Tasks";
            byte[] template = CreateDocWithTableTemplate($"@@Table:{tableId}@@", new List<string> { "@@Description@@", "@@Status@@" });

            var tableData = new Dictionary<string, List<Dictionary<string, string>>>
            {
                { tableId, new List<Dictionary<string, string>>() } // Empty list of rows
            };

            byte[] result = wordEngine.FillTemplate(template, null, tableData);
            List<List<string>> resultTable = GetTableData(result);

            // Identifier row remains, template row is removed, no data rows added.
            Assert.Equal(1, resultTable.Count);
            Assert.Equal("", resultTable[0][0].Trim()); // Identifier cell cleared
        }

        [Fact]
        public void FillTemplate_TableTagNotFound_LeavesTableAsIs()
        {
            string tableId = "Orders";
            byte[] template = CreateDocWithTableTemplate($"@@Table:{tableId}@@", new List<string> { "@@OrderID@@", "@@Amount@@" });

            var tableData = new Dictionary<string, List<Dictionary<string, string>>>
            {
                {
                    "NonExistentTable", // Data for a table not in the template
                    new List<Dictionary<string, string>>
                    {
                        new Dictionary<string, string> { { "@@OrderID@@", "100" }, { "@@Amount@@", "500" } }
                    }
                }
            };

            byte[] result = wordEngine.FillTemplate(template, null, tableData);
            List<List<string>> resultTable = GetTableData(result);

            // Table should be unchanged (2 rows: identifier, template)
            Assert.Equal(2, resultTable.Count);
            Assert.Contains($"@@Table:{tableId}@@", resultTable[0][0]); // Identifier still there
            Assert.Equal("@@OrderID@@", resultTable[1][0]);
            Assert.Equal("@@Amount@@", resultTable[1][1]);
        }

        [Fact]
        public void FillTemplate_MultipleTables_PopulatesCorrectly()
        {
            // Create a document with two tables
            string table1Id = "Employees";
            List<string> table1Cols = new List<string> { "@@Name@@", "@@Role@@" };
            string table2Id = "Departments";
            List<string> table2Cols = new List<string> { "@@DeptName@@", "@@Manager@@" };

            byte[] template;
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Table 1
                    Table table1 = new Table();
                    table1.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text($"@@Table:{table1Id}@@")))), new TableCell(new Paragraph(new Run(new Text(""))))));
                    TableRow table1TemplateRow = new TableRow();
                    foreach(var col in table1Cols) table1TemplateRow.Append(new TableCell(new Paragraph(new Run(new Text(col)))));
                    table1.Append(table1TemplateRow);
                    body.Append(table1);

                    body.AppendChild(new Paragraph(new Run(new Text("Some text between tables."))));

                    // Table 2
                    Table table2 = new Table();
                    table2.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text($"@@Table:{table2Id}@@")))), new TableCell(new Paragraph(new Run(new Text(""))))));
                    TableRow table2TemplateRow = new TableRow();
                    foreach(var col in table2Cols) table2TemplateRow.Append(new TableCell(new Paragraph(new Run(new Text(col)))));
                    table2.Append(table2TemplateRow);
                    body.Append(table2);

                    mainPart.Document.Save();
                }
                template = mem.ToArray();
            }

            var tableData = new Dictionary<string, List<Dictionary<string, string>>>
            {
                {
                    table1Id,
                    new List<Dictionary<string, string>>
                    {
                        new Dictionary<string, string> { { "@@Name@@", "Eve" }, { "@@Role@@", "Engineer" } }
                    }
                },
                {
                    table2Id,
                    new List<Dictionary<string, string>>
                    {
                        new Dictionary<string, string> { { "@@DeptName@@", "IT" }, { "@@Manager@@", "Adam" } },
                        new Dictionary<string, string> { { "@@DeptName@@", "HR" }, { "@@Manager@@", "Olivia" } }
                    }
                }
            };

            byte[] result = wordEngine.FillTemplate(template, null, tableData);

            List<List<string>> resultTable1 = GetTableData(result, 0); // First table
            List<List<string>> resultTable2 = GetTableData(result, 1); // Second table

            // Check Table 1
            Assert.Equal(2, resultTable1.Count); // Identifier row + 1 data row
            Assert.Equal("", resultTable1[0][0].Trim());
            Assert.Equal("Eve", resultTable1[1][0]);
            Assert.Equal("Engineer", resultTable1[1][1]);

            // Check Table 2
            Assert.Equal(3, resultTable2.Count); // Identifier row + 2 data rows
            Assert.Equal("", resultTable2[0][0].Trim());
            Assert.Equal("IT", resultTable2[1][0]);
            Assert.Equal("Adam", resultTable2[1][1]);
            Assert.Equal("HR", resultTable2[2][0]);
            Assert.Equal("Olivia", resultTable2[2][1]);
        }
    }
}
