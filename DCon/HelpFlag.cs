using System;
using System.IO;
using System.Xml;

namespace DCon
{
    public static class HelpFlag
    {
        //TODO: Fix Absolute Path and change to install Path
        private static string _helpFile = "HelpDocs.xml"; //@"C:\Users\Douglas\Documents\Visual Studio 2017\Projects\DCon\DCon\HelpDocs.xml"

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
            try
            {

            
                // Start with XmlReader object
                // Here, we try to setup Stream between the XML file and xmlReader
                using (XmlReader reader = XmlReader.Create(_helpFile))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            if (reader.Name.ToString() == "Example")
                            {
                                Console.WriteLine(reader.ReadString());
                            }
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error Occurred when accessing " + _helpFile);
            }
        }
    }
}
