using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Power = Microsoft.Office.Interop.PowerPoint;
using iTextSharp.text.pdf;

namespace DCon
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args[0].ToUpper() == "-HELP" || args[0].ToUpper() == "-H")
            {
                HelpCommand.GetHelp();
            }
            else if (args[0].ToUpper().Contains(".DOCX") || args[0].ToUpper().Contains(".DOC") || args[0].ToUpper().Contains(".RTF") || args[0].ToUpper().Contains(".DOT") || args[0].ToUpper().Contains(".ODT"))
            {

                string input = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), args[0]));
                //string output = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), args[1]));

                Word.Application app = new Word.Application();

                try
                {
                    Word.Document document = app.Documents.Open(input);
                    //app.ActiveDocument.SaveAs(output, WdSaveFormat.wdFormatText);
                    Console.WriteLine(document.Content.Text.ToString());
                }
                finally
                {
                    app.Quit();
                }
            }
            else if (args[0].ToUpper().Contains(".XLSX") || args[0].ToUpper().Contains(".XLS"))
            {

                string input = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), args[0]));
                //string output = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), args[1]));

                Excel.Application app = new Excel.Application();

                try
                {
                    app.DisplayAlerts = false;
                    app.Visible = false;

                    Excel.Workbook book = app.Workbooks.Open(input);
                    //book.SaveAs(output, XlFileFormat.xlTextWindows);

                    Excel.Worksheet worksheet = book.ActiveSheet;
                    Excel.Range xlRange = worksheet.UsedRange;

                    //Console.WriteLine(xlRange.Cells.Value2.ToString());

                    foreach (Excel.Range s in xlRange.Cells)
                    {
                        Console.WriteLine(s.Value2.ToString());
                    }

                }
                finally
                {
                    app.Quit();
                }
            }

            //else if (args[0].Contains(".pptx") || args[0].Contains(".ppt"))
            //{

            //    string input = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), args[0]));
            //    string output = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), args[1]));

            //    Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();

            //    try
            //    {
            //        Microsoft.Office.Interop.PowerPoint.Presentation pres = app.Presentation.Open(input);
            //    }
            //    finally
            //    {
            //        app.Quit();
            //    }

            //}

            else if (args[0].ToUpper().Contains(".PDF"))
            {

                string input = args[0];
                //string output = args[1];
                //string strText = string.Empty;
                //try
                //{
                //    PdfReader reader = new PdfReader(input);

                //    for (int page = 1; page <= reader.NumberOfPages; page++)
                //    {
                //        iTextSharp.text.pdf.parser.ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
                //        String s = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, page, its);

                //        s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
                //        strText = strText + s;
                //        //File.WriteAllText(output, strText);
                //        Console.WriteLine(strText);
                //    }
                //    reader.Close();

                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.ToString());
                //}
                DocConverter.PDF(input);
            }
        }
    }
}
