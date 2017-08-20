using System;
using System.IO;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Power = Microsoft.Office.Interop.PowerPoint;
using iPDF = iTextSharp.text.pdf;


namespace DCon
{
    public static class DocConverter
    {
        public static void Word(string input)
        {

            string doc = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), input));

            Word.Application app = new Word.Application();

            try
            {
                Word.Document document = app.Documents.Open(doc);
                Console.WriteLine(document.Content.Text.ToString());
            }
            finally
            {
                app.Quit();
            }
        }

        public static void Excel(string input)
        {
            string doc = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), input));

            Excel.Application app = new Excel.Application();

            try
            {
                app.DisplayAlerts = false;
                app.Visible = false;

                Excel.Workbook book = app.Workbooks.Open(doc);

                Excel.Worksheet worksheet = book.ActiveSheet;
                Excel.Range xlRange = worksheet.UsedRange;

                foreach (Excel.Range cell in xlRange.Cells)
                {
                    Console.WriteLine(cell.Value2.ToString());
                }
            }
            finally
            {
                app.Quit();
            }
        }

        public static void PowerPoint(string input)
        {
            //    string doc = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), input));

            //    Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();

            //    try
            //    {
            //        Microsoft.Office.Interop.PowerPoint.Presentation pres = app.Presentation.Open(doc);
            //    }
            //    finally
            //    {
            //        app.Quit();
            //    }
        }

        public static void PDF(string input)
        {
            string doc = string.Empty;
            try
            {
                iPDF.PdfReader reader = new iPDF.PdfReader(input);

                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    iPDF.parser.ITextExtractionStrategy its = new iPDF.parser.LocationTextExtractionStrategy();
                    String parseString = iPDF.parser.PdfTextExtractor.GetTextFromPage(reader, page, its);

                    parseString = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(parseString)));
                    doc = doc + parseString;
                    Console.WriteLine(doc);
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
