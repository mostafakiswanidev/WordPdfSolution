using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using WordToPDF;

namespace WordPdfSolution
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = "D:\\WOrdToPDF\\word.docx";
            string name = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string outputFile = "D:\\WOrdToPDF\\" + name + ".docx";

            // Copy Word document.
            File.Create(outputFile).Dispose();

            File.Copy(inputFile, outputFile, true);
            // Open copied document.
            using (var flatDocument = new FlatDocument(outputFile))
            {
                // Search and replace document's text content.
                flatDocument.FindAndReplace("انثى", "ذكر");
                flatDocument.FindAndReplace("Qudamah", "Mostafa");
            }

            Word2Pdf objWorPdf = new Word2Pdf();
            #region ToReplace
            string backfolder1 = "D:\\WOrdToPDF\\";
            //string strFileName = "word.docx";
            string strFileName = name + ".docx";
            #endregion

            object FromLocation = backfolder1 + "\\" + strFileName;
            string FileExtension = Path.GetExtension(strFileName);
            string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");
            if (FileExtension == ".doc" || FileExtension == ".docx")
            {
                object ToLocation = backfolder1 + "\\" + ChangeExtension;
                objWorPdf.InputLocation = FromLocation;
                objWorPdf.OutputLocation = ToLocation;
                objWorPdf.Word2PdfCOnversion();
            }



            Console.WriteLine("Hello World!");
        }
    }
}
