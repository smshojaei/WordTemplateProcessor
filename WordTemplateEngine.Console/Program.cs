// See https://aka.ms/new-console-template for more information
using WordTemplateEngine;
using static System.Runtime.InteropServices.JavaScript.JSType;

Console.WriteLine("Hello, World!");

Test1();

//Test2();


static void Test1()
{
    var tags = new Dictionary<string, string>
    {
        { "Port", "شهید رجایی" },
        { "Date", "1403/03/03" },
        { "No", "43-GB45-456" },
        { "VoyageNoticeNo", "582961-r26" },
        { "VoyageNoticeDate", "1403/03/03" },
        { "ArrivalVoyageNo", "434456-4564" },
        { "Terminal", "خزر قشم" },
        { "ExpirationDate", "1403/03/03" },
        { "OwnerName", "رضا محسنی" },
        { "OwnerIdNumber", "13211321231" },
        { "OwnerAddress", "تهران خ یسبتستمینت، خ سیمبتنک، خ نیبملمتیبلت خ یلبلن" },
    };

    var tableData = new Dictionary<string, List<Dictionary<string, string>>>
    {
        ["Items"] = new List<Dictionary<string, string>>
        {
            new() { ["FirstDischargeDate"] = "1403/03/03",
                    ["GoodWidth"         ] = "25356",
                    ["Quantity"          ] = "1253",
                    ["GoodPackageType"   ] = "عدد",
                    ["GoodHSCode"        ] = "4534534",
                    ["GoodBrandName"     ] = "مداد"
                 },
            new() { ["FirstDischargeDate"] = "1403/03/02",
                    ["GoodWidth"         ] = "345345",
                    ["Quantity"          ] = "3434",
                    ["GoodPackageType"   ] = "عدد",
                    ["GoodHSCode"        ] = "534534",
                    ["GoodBrandName"     ] = "خودکار"
                 },
            new() { ["FirstDischargeDate"] = "1403/03/03",
                    ["GoodWidth"         ] = "3435",
                    ["Quantity"          ] = "3453445",
                    ["GoodPackageType"   ] = "متر",
                    ["GoodHSCode"        ] = "534535354",
                    ["GoodBrandName"     ] = "طناب "
                 },
        }
    };

    var processor = new WordTemplateProcessor();
    var templateBytes = File.ReadAllBytes("d:\\Temp\\Q01.docx");
    var filledBytes = processor.FillTemplate(templateBytes, tags, tableData);
    var word2Pdf = new Word2Pdf();
    word2Pdf.ConvertDocx2Pdf(filledBytes, "d:\\Temp\\Q01_o.pdf");

}

static void Test2()
{
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

    var templateBytes = File.ReadAllBytes("d:\\Temp\\doc01.docx");
    var processor = new WordTemplateProcessor();
    var filledBytes = processor.FillTemplate(templateBytes, tags, tableData);
    var word2Pdf = new Word2Pdf();
    word2Pdf.ConvertDocx2Pdf(filledBytes, "d:\\Temp\\doc01_fill1.pdf");
    //File.WriteAllBytes("d:\\Temp\\doc01_fill3.docx", filledBytes);
}