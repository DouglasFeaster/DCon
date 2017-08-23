using System;
using System.Xml;

namespace DCon
{
    public static class HelpFlag
    {
#if DEBUG
        private static string _helpFile = "HelpDocs.xml";
#else
        private static string _helpFile = @"C:\Program Files (x86)\Johnson University\Document Converter\HelpDocs.xml";
#endif

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
