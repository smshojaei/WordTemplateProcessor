using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Licensing;
using System.Runtime.Serialization.Formatters.Binary;

namespace WordTemplateEngine
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
