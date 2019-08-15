using System;
using Aspose.Words;

namespace WordOpenXmlClassConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Word2PDF("C:/Users/shenh/Desktop/1565852520.docx", "C:/Users/shenh/Desktop/1565852520.pdf");
        }
        public static void Word2PDF(string srcDocPath, string dstPdfPath)
        {
            // Load the document from disk.
            Document srcDoc = new Document(srcDocPath);
            // Save the document in PDF format.
            srcDoc.Save(dstPdfPath, SaveFormat.Pdf);

        }
    }
}
