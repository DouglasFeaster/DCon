using System;
using System.Xml;

namespace DCon
{
    public static class FileExt
    {
        private static string _fileExtFile = "OfficeFileExt.xml";

        public static bool IsWord(string input)
        {
            string ext = String.Empty;
            // Start with XmlReader object
            // Here, we try to setup Stream between the XML file and xmlReader
            using (XmlReader reader = XmlReader.Create(_fileExtFile))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name.ToString() == "Word")
                        {
                            if (reader.Name.ToString() == "Extension")
                            {
                                ext = reader.ReadString();
                            }
                        }
                    }
                }
            }

            if (input.ToUpper().Contains(ext.ToUpper()))
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
            string ext = String.Empty;
            // Start with XmlReader object
            // Here, we try to setup Stream between the XML file and xmlReader
            using (XmlReader reader = XmlReader.Create(_fileExtFile))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name.ToString() == "Excel")
                        {
                            if (reader.Name.ToString() == "Extension")
                            {
                                ext = reader.ReadString();
                            }
                        }
                    }
                }
            }

            if (input.ToUpper().Contains(ext.ToUpper()))
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
            string ext = String.Empty;
            // Start with XmlReader object
            // Here, we try to setup Stream between the XML file and xmlReader
            using (XmlReader reader = XmlReader.Create(_fileExtFile))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name.ToString() == "Power")
                        {
                            if (reader.Name.ToString() == "Extension")
                            {
                                ext = reader.ReadString();
                            }
                        }
                    }
                }
            }

            if (input.ToUpper().Contains(ext.ToUpper()))
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
            string ext = String.Empty;
            // Start with XmlReader object
            // Here, we try to setup Stream between the XML file and xmlReader
            using (XmlReader reader = XmlReader.Create(_fileExtFile))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name.ToString() == "PDF")
                        {
                            if (reader.Name.ToString() == "Extension")
                            {
                                ext = reader.ReadString();
                            }
                        }
                    }
                }
            }

            if (input.ToUpper().Contains(ext.ToUpper()))
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
