using System;
using System.Xml;

namespace DCon
{

    /// <summary>
    /// File Extension Differentiation Class 
    /// </summary>
    public static class FileExt
    {

#if DEBUG
        private static string _fileExtFile = "OfficeFileExt.xml";
#else
        private static string _fileExtFile = @"C:\Program Files (x86)\Johnson University\Document Converter\OfficeFileExt.xml";
#endif

        /// <summary>
        /// Evaluates Input File Path to see if target is MS Word Document 
        /// </summary>
        /// <param name="input">Input File Path</param>
        /// <returns>True or False</returns>
        public static bool IsWord(string input)
        {
            string ext = String.Empty;

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(_fileExtFile);
                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Office/Word");

                foreach (XmlNode node in nodeList)
                {
                    foreach (XmlNode item in node.SelectNodes("Extension"))
                    {
                        ext = item.InnerText;

                        if (input.ToUpper().Contains(ext.ToUpper()))
                        {
                            break;
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error Occurred when accessing " + _fileExtFile);
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

        /// <summary>
        /// Evaluates Input File Path to see if target is MS Excel Spreadsheet 
        /// </summary>
        /// <param name="input">Input File Path</param>
        /// <returns>True or False</returns>
        public static bool IsExcel(string input)
        {
            string ext = String.Empty;

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(_fileExtFile);
                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Office/Excel");

                foreach (XmlNode node in nodeList)
                {
                    foreach (XmlNode item in node.SelectNodes("Extension"))
                    {
                        ext = item.InnerText;

                        if (input.ToUpper().Contains(ext.ToUpper()))
                        {
                            break;
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error Occurred when accessing " + _fileExtFile);
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

        /// <summary>
        /// Evaluates Input File Path to see if target is MS PowerPoint Presentation 
        /// </summary>
        /// <param name="input">Input File Path</param>
        /// <returns>True or False</returns>
        public static bool IsPowerPoint(string input)
        {
            string ext = String.Empty;

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(_fileExtFile);
                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Office/Power");

                foreach (XmlNode node in nodeList)
                {
                    foreach (XmlNode item in node.SelectNodes("Extension"))
                    {
                        ext = item.InnerText;

                        if (input.ToUpper().Contains(ext.ToUpper()))
                        {
                            break;
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error Occurred when accessing " + _fileExtFile);
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

        /// <summary>
        /// Evaluates Input File Path to see if target is PDF Document
        /// </summary>
        /// <param name="input">Input File Path</param>
        /// <returns>True or False</returns>
        public static bool IsPDF(string input)
        {
            string ext = String.Empty;

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(_fileExtFile);
                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Office/PDF");

                foreach (XmlNode node in nodeList)
                {
                    foreach (XmlNode item in node.SelectNodes("Extension"))
                    {
                        ext = item.InnerText;

                        if (input.ToUpper().Contains(ext.ToUpper()))
                        {
                            break;
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error Occurred when accessing " + _fileExtFile);
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
