using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCon
{
    public static class HelpCommand
    {
        public static bool IsHelp(string input)
        {
            if (input.ToUpper() == "-HELP" || input.ToUpper() == "-H")
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }

        public static void GetHelp()
        {
            //Console.WriteLine("usage: dcon [-help, -h] <inputPath>\n");
            //Console.WriteLine("example: dcon input.docx \n dcon input.xlsx \n input.pdf");

            // Start with XmlReader object
            //here, we try to setup Stream between the XML file nad xmlReader
            using (XmlReader reader = XmlReader.Create("helpDoc.xml"))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        //return only when you have START tag
 
                        switch (reader.Name.ToString())
                        {
                            case "Usage":
                                Console.WriteLine("Usage : " + reader.ReadString());
                                break;
 
                            case "Example":
                                Console.WriteLine("Example : " + reader.ReadString());
                                break;
                        }
                    }
                }
            }
        }
    }
}
