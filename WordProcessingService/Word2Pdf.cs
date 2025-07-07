using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Licensing;
using Syncfusion.Pdf;

namespace WordProcessingService
{
    public class Word2Pdf
    {
        public Word2Pdf()
        {
            SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NHaF5cWWdCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdmWXpfeXVRR2hYV0R+W0NWYUo=");
        }

        public void ConvertDocx2Pdf(byte[] inputDocument, string destinationFilPath)
        {

            MemoryStream memStream = new MemoryStream(inputDocument);
            using WordDocument document = new(memStream, FormatType.Docx);

            using DocIORenderer renderer = new();
            using PdfDocument pdfDocument = renderer.ConvertToPDF(document);

            using FileStream outputStream = new(destinationFilPath, FileMode.Create, FileAccess.Write);
            pdfDocument.Save(outputStream);

        }
    }
}
