using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCon
{
    public static class ColumnLetters
    {
#if DEBUG
        private static string _ABCsFile = "ABCs.xml";
#else
        private static string _ABCsFile = @"C:\Program Files (x86)\Johnson University\Document Converter\ABCs.xml";
#endif
        public static string GetLetter(string input)
        {
            string letter = String.Empty;

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(_ABCsFile);
                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Alphabet");

                foreach (XmlNode node in nodeList)
                {
                    foreach (XmlNode item in node.SelectNodes("Letter"))
                    {
                        letter = item.InnerText;

                       
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error Occurred when accessing " + _ABCsFile);
            }

            return letter;
        }
    }
}
