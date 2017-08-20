using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCon
{
    public static class HelpCommand
    {
        public static void GetHelp()
        {
            Console.WriteLine("usage: dcon [-help, -h] <inputPath>\n");
            Console.WriteLine("example: dcon input.docx \n dcon input.xlsx \n input.pdf");
        }
    }
}
