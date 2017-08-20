using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCon
{
    public static class FileExt
    {
        public static bool IsWord(string input)
        {
            // Start with XmlReader object
            //here, we try to setup Stream between the XML file nad xmlReader
            using (XmlReader reader = XmlReader.Create("OfficeFileExt.xml"))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name.ToString() == "Extension")
                        {
                            Console.WriteLine(reader.ReadString());
                        }
                    }
                }
            }

            if (input.ToUpper().Contains(".DOCX") || input.ToUpper().Contains(".DOC") 
                || input.ToUpper().Contains(".RTF") || input.ToUpper().Contains(".DOT") 
                || input.ToUpper().Contains(".ODT"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsExcel(string input)
        {
            if (input.ToUpper().Contains(".XLSX") || input.ToUpper().Contains(".XLS") || input.ToUpper().Contains(".ODS") || input.ToUpper().Contains(".CSV"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsPowerPoint(string input)
        {
            if (input.Contains(".PPTX") || input.Contains(".PPT"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsPDF(string input)
        {
            if (input.ToUpper().Contains(".PDF"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
