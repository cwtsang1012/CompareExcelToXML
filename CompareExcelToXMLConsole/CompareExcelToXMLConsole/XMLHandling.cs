using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Xml.Linq;

namespace CompareExcelToXMLConsole
{
    public class XMLHandling
    {
        public List<Dictionary<string, string>> ReadXMLData(string xmlFilePath, List<string> keyList)
        {
            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine("Unable to find the XML File!");
                return null;
            }
            else
            {
                List<Dictionary<string, string>> xmlDataList = new List<Dictionary<string, string>>();
                XDocument doc = XDocument.Load(xmlFilePath);
                foreach (XElement jobElement in doc.Element("DEFTABLE").Elements("FOLDER").Elements())
                {
                    if (jobElement.HasAttributes)
                    {
                        Dictionary<string, string> jobDict = new Dictionary<string, string>();
                        List<XAttribute> attributes = jobElement.Attributes().ToList();
                        foreach (string key in keyList)
                        {
                            var attr = attributes.FindAll(a => a.Name == key);
                            if (attr.Count > 0)
                            {
                                jobDict.Add(key, attr[0].Value.ToString());
                            }

                            
                        }

                        if (jobDict.Count > 1)
                            xmlDataList.Add(jobDict);
                    }
                    
                }

                return xmlDataList;
            }
        }
    }
}
