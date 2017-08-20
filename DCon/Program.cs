using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

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
                DocConverter.Word(args[0]);
            }
            else if (args[0].ToUpper().Contains(".XLSX") || args[0].ToUpper().Contains(".XLS"))
            {
                DocConverter.Excel(args[0]);
            }
            else if (args[0].Contains(".pptx") || args[0].Contains(".ppt"))
            {


            }
            else if (args[0].ToUpper().Contains(".PDF"))
            {
                DocConverter.PDF(args[0]);
            }
        }
    }
}
