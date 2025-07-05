// See https://aka.ms/new-console-template for more information
using WordTemplateEngine;

Console.WriteLine("Hello, World!");

var processor = new WordTemplateProcessor();

var templateBytes = File.ReadAllBytes("d:\\Temp\\doc01.docx");

var tags = new Dictionary<string, string>
{
    { "Name", "محسن" },
    { "Date", "1403/04/14" }
};

var tableData = new Dictionary<string, List<Dictionary<string, string>>>
{
    ["Employees"] = new List<Dictionary<string, string>>
    {
        new() { ["EName"] = "علی", ["EAge"] = "40", ["ERemark"] = "این برای تست 123 است." },
      new() { ["EName"] = "سارا", ["EAge"] = "30" },
        new() { ["EName"] = "علی", ["EAge"] = "40", ["ERemark"] = "این برای تست 123 است." },
        new() { ["EName"] = "علی", ["EAge"] = "40", ["ERemark"] = "این برای تست 123 است." },
        new() { ["EName"] = "علی", ["EAge"] = "40", ["ERemark"] = "این برای تست 123 است." }
    }
};

var filledBytes = processor.FillTemplate(templateBytes, tags, tableData);
File.WriteAllBytes("d:\\Temp\\doc01_fill3.docx", filledBytes);