using System;
using System.Xml;

namespace DCon
{

    /// <summary>
    /// Help Flag Class
    /// </summary>
    public static class HelpFlag
    {
#if DEBUG
        private static string _helpFile = "HelpDocs.xml";
#else
        private static string _helpFile = @"C:\Program Files (x86)\Johnson University\Document Converter\HelpDocs.xml";
#endif

        /// <summary>
        /// Evaluates Input Argument to see if Input Argument is Help Command
        /// </summary>
        /// <param name="inputArg">Command Line Input Argument</param>
        /// <returns>True or False</returns>
        public static bool IsHelp(string inputArg)
        {
            if (inputArg.ToUpper() == "-HELP" || inputArg.ToUpper() == "-H")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Get Help Class for writing help from HelpDocs XML
        /// </summary>
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
