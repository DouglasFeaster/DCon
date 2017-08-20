using System;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCon
{
    public static class HelpFlag
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
            // Start with XmlReader object
            // Here, we try to setup Stream between the XML file and xmlReader
            using (XmlReader reader = XmlReader.Create("HelpDocs.xml"))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name.ToString() == "Examples")
                        {
                            if (reader.Name.ToString() == "Example")
                            {
                                Console.WriteLine(reader.ReadString());
                            }
                        }
                    }
                }
            }
        }
    }
}
