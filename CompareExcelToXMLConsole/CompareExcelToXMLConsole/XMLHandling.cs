using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Xml.Linq;
using System.Collections.Specialized;
using System.Data;

namespace CompareExcelToXML
{
    public class XMLHandling
    {
        public DataTable ReadXMLData(string xmlFilePath, string sectionName, DataTable dtExcel)
        {
            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine("Unable to find the XML File!");
                return null;
            }
            else
            {
                //Duplicate the column definition from Excel
                DataTable dtXML = new DataTable();
                foreach (DataColumn column in dtExcel.Columns)
                {
                    var dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
                    dc.ColumnName = column.ColumnName;
                    dtXML.Columns.Add(dc);
                }

                NameValueCollection sectionSettings = ConfigurationManager.GetSection(sectionName) as NameValueCollection;
                Dictionary<string, string> xmlDataList = new Dictionary<string, string>();
                XDocument doc = XDocument.Load(xmlFilePath);
                foreach (XElement jobElement in doc.Element("DEFTABLE").Elements("FOLDER").Elements())
                {
                    if (jobElement.HasAttributes)
                    {
                        List<XAttribute> attributes = jobElement.Attributes().ToList();
                        var dr = dtXML.NewRow();
                        foreach (string key in sectionSettings.AllKeys.Where(k => k.IndexOf('-') == -1))
                        {
                            //Retrieve Attribute value
                            var attr = attributes.FindAll(a => a.Name == key);
                            if (attr.Count > 0)
                            {
                                var columnIndex = Convert.ToInt32(sectionSettings[key]);
                                dr[columnIndex] = attr[0].Value.ToString();
                                //jobDict.Add(key, attr[0].Value.ToString());
                            }
                        }

                        Dictionary<string, string> jobDict = new Dictionary<string, string>();
                        foreach (string key in sectionSettings.AllKeys.Where(k => k.IndexOf('-') > -1))
                        {
                            switch (ConfigurationManager.AppSettings["FileType"])
                            {
                                case "ControlM":
                                    var nodeName = key.Split('-')[0];
                                    var nodeAttributes = key.Split('-')[1];
                                    var nodeAttributeKey = nodeAttributes.Split('|')[0];
                                    var nodeAttributeValue = nodeAttributes.Split('|')[1];
                                    var columnIndex = Convert.ToInt32(sectionSettings[key].Split(';')[0]);
                                    var requiredKeys = sectionSettings[key].Split(';')[1].Split('|');
                                    foreach (XElement childElememnt in jobElement.Elements(nodeName))
                                    {
                                        var attr = childElememnt.Attributes().Single(a => a.Name == nodeAttributeKey);
                                        if (attr != null)
                                        {
                                            if (nodeAttributeKey != "CODE")
                                            {
                                                if (requiredKeys.Any(k => attr.Value.IndexOf(k) > -1))
                                                {
                                                    jobDict.Add(attr.Value, childElememnt.Attributes().Single(a => a.Name == nodeAttributeValue).Value);
                                                }
                                            }
                                            else
                                            {
                                                if (childElememnt.Attributes().Single(a => a.Name == nodeAttributeKey).Value == "NOTOK")
                                                {
                                                    XElement subChildElement = childElememnt.Elements("DOMAIL").ToList().FirstOrDefault();
                                                    if (subChildElement != null)
                                                    {
                                                        //Conidtion 1: remove those words added by control-m for hostname
                                                        var unformattedMsg = subChildElement.Attributes().Single(a => a.Name == nodeAttributeValue).Value;
                                                        var unusedString = unformattedMsg.IndexOf("00000018Hostname:");
                                                        var outputMsg = (unusedString == -1) ? unformattedMsg : unformattedMsg.Substring(0, unusedString);
                                                        //Condition 2: Remove all pattern 00XX or 00000XXX
                                                        var msg1Array = outputMsg.Split(' ');
                                                        var patternArr = msg1Array.Where(s => s.IndexOf("00000") > -1 || s.IndexOf("00") > -1);
                                                        foreach (var word in patternArr)
                                                        {
                                                            if (word.Where(w => w.Equals('0')).Count() >= 5)
                                                            {
                                                                outputMsg = outputMsg.Replace(word, word.Replace(word.Substring(word.IndexOf("00"), 8), ""));
                                                            }
                                                            else if (word.Where(w => w.Equals('0')).Count() > 1 && word.Where(w => w.Equals('0')).Count() < 5)
                                                            {
                                                                outputMsg = outputMsg.Replace(word, word.Replace(word.Substring(word.IndexOf("00"), 4), ""));
                                                            }
                                                        }

                                                        //jobDict.Add(nodeAttributeValue, outputMsg);
                                                        dr[columnIndex] = outputMsg;
                                                    }
                                                }
                                            }

                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }

                        if (jobDict.Count() > 0) 
                        {
                            //Handling Parameters
                            string param = "";
                            foreach (var element in jobDict) 
                            {
                                string key = element.Key;
                                string value = element.Value;

                                if (key.Contains("%%FileWatch-"))
                                {
                                    key = key.Replace("%%FileWatch-", "");
                                    switch (key) 
                                    {
                                        case "FILE_PATH":
                                            var seperateIndex = value.LastIndexOf('\\');
                                            string scriptPath = value.Substring(0, seperateIndex + 1);
                                            string scriptName = value.Substring(seperateIndex + 1);
                                            dr[7] = scriptPath;
                                            dr[6] = scriptName;
                                            break;
                                        case "TIME_LIMIT":
                                            dr[11] = dr[11] + "\n(Time Limit: " + value + "minutes )\n";
                                            break;
                                        case "MIN_AGE":
                                            dr[11] = dr[11] + "File age: Minimal age = " + value + "Min";
                                            break;
                                    }
                                }
                                else 
                                {
                                    if (param == "")
                                    {
                                        param = element.Key.ToString() + " = " + element.Value.ToString();
                                    }
                                    else
                                    {
                                        param += "\n" + element.Key.ToString() + " = " + element.Value.ToString();
                                    }
                                    dr[8] = param;
                                }
                            }
                        }
                        dtXML.Rows.Add(dr);
                    }

                }
                return dtXML;
            }
        }
    }
}
