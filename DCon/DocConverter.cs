using System;
using System.IO;
using System.Text;

using iWord = Microsoft.Office.Interop.Word;
using iExcel = Microsoft.Office.Interop.Excel;
using iPower = Microsoft.Office.Interop.PowerPoint;
using iPDF = iTextSharp.text.pdf;


namespace DCon
{
    public static class DocConverter
    {
        public static void Word(string input)
        {

            string doc = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), input));

            iWord.Application app = new iWord.Application();

            try
            {
                iWord.Document document = app.Documents.Open(doc);
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

            iExcel.Application app = new iExcel.Application();

            try
            {
                app.DisplayAlerts = false;
                app.Visible = false;

                iExcel.Workbook book = app.Workbooks.Open(doc);

                iExcel.Worksheet worksheet = book.ActiveSheet;
                iExcel.Range xlRange = worksheet.UsedRange;

                object[,] valueArray = (object[,])xlRange.get_Value(
                        iExcel.XlRangeValueDataType.xlRangeValueDefault);
                
                // iterate through each cell and display the contents.

                string colLet = string.Empty;

                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    { 
                        try
                        {
                            Console.WriteLine(ColumnLetters.GetLetter(col) + row + ":" + valueArray[row, col].ToString());
                        }
                        catch
                        {
                            Console.WriteLine(ColumnLetters.GetLetter(col) + row + ": ");
                        }
                    }
                }
                book.Close();
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
