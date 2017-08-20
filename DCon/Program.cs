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
            try
            {
                if (HelpFlag.IsHelp(args[0]))
                {
                    HelpFlag.GetHelp();
                }
                else if (FileExt.IsWord(args[0]))
                {
                    DocConverter.Word(args[0]);
                }
                else if (FileExt.IsExcel(args[0]))
                {
                    DocConverter.Excel(args[0]);
                }
                else if (FileExt.IsPowerPoint(args[0]))
                {
                
                }
                else if (FileExt.IsPDF(args[0]))
                {
                    DocConverter.PDF(args[0]);
                }
            }
            catch
            {
                HelpFlag.GetHelp();
            }
            
        }
    }
}
