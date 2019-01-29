using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading.Tasks;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Specialized;
using System.Text.RegularExpressions;

namespace CompareExcelToXML
{
    class Program
    {
        public static void Main(string[] args)
        {
            //C#讀取Excel 幾種方法的體會
            Console.WriteLine("Program Start.");

            ExcelHandling excel = new ExcelHandling();
            DataTable dtExcel = excel.GetExcelData(ConfigurationManager.AppSettings["ExcelFilePath"]);
            Console.WriteLine("Retrieve Excel data successfully.");

            XMLHandling xml = new XMLHandling();
            DataTable dtXML = xml.ReadXMLData(ConfigurationManager.AppSettings["XMLFilePath"], "ControlMDataColumn", dtExcel);
            Console.WriteLine("Retrieve XML data successfully.");

            DataRow[] dtFilterExcelRows = null;
            if (ConfigurationManager.AppSettings["SectionName"] != "")
            {
                dtFilterExcelRows = dtExcel.Select("Jobname LIKE '%" + ConfigurationManager.AppSettings["SectionName"] + "%'");
            }
            else
            {
                dtFilterExcelRows = dtExcel.AsEnumerable().ToArray();
            }

            DataSet dsOutput = CompareExcelToXML(dtFilterExcelRows, dtXML);
            excel.ToExcelSheet(dsOutput, ConfigurationManager.AppSettings["OutputFilePath"]);

            Console.WriteLine("The result file is generated successfully.");
            Console.WriteLine("Program End.");
            Console.ReadKey();
        }

        public static DataSet CompareExcelToXML(DataRow[] dtExcelRows, DataTable dtXML)
        {
            DataSet ds = new DataSet();

            DataTable dtSummary = new DataTable();
            dtSummary.TableName = "Summary";

            DataColumn dcs1 = new DataColumn();
            dcs1.DataType = System.Type.GetType("System.String");
            dcs1.ColumnName = "JobName";
            dtSummary.Columns.Add(dcs1);
            DataColumn dcs2 = new DataColumn();
            dcs2.DataType = System.Type.GetType("System.String");
            dcs2.ColumnName = "Invalid Column";
            dtSummary.Columns.Add(dcs2);

            DataTable dtDetail = new DataTable();
            dtDetail.TableName = "Details";

            DataColumn dcd1 = new DataColumn();
            dcd1.DataType = System.Type.GetType("System.String");
            dcd1.ColumnName = "JobName";
            dtDetail.Columns.Add(dcd1);
            DataColumn dcd2 = new DataColumn();
            dcd2.DataType = System.Type.GetType("System.String");
            dcd2.ColumnName = "Excel Column Name";
            dtDetail.Columns.Add(dcd2);
            DataColumn dcd3 = new DataColumn();
            dcd3.DataType = System.Type.GetType("System.String");
            dcd3.ColumnName = "Excel Value";
            dtDetail.Columns.Add(dcd3);
            DataColumn dcd4 = new DataColumn();
            dcd4.DataType = System.Type.GetType("System.String");
            dcd4.ColumnName = "XML Value";
            dtDetail.Columns.Add(dcd4);

            var exludeColumns = ConfigurationManager.AppSettings["ExcludeColumns"].Split('|');
            foreach (var row in dtExcelRows)
            {
                var xmlRow = dtXML.Select("Jobname = '" + row[3] + "'");
                if (xmlRow.Count() > 0)
                {
                    foreach (DataColumn column in dtXML.Columns)
                    {
                        if (!exludeColumns.Contains(column.Ordinal.ToString()))
                        {
                            bool equal = true;
                            var transformedExcelValue = "";
                            var transformedXMLValue = "";
                            switch (column.Ordinal)
                            {
                                case 0:
                                    //Only consider "Disable" case
                                    if (row[column.Ordinal].ToString().ToLower() == "disable")
                                    {
                                        if (xmlRow[0][column.Ordinal].ToString().Trim() == "")
                                            equal = false;
                                    }
                                    break;
                                case 4:
                                    //Ignore space, new line(\n) & \r
                                    transformedExcelValue = row[column.Ordinal].ToString().Trim().Replace(" ", "").Replace("\n", "").Replace("\r", "");
                                    transformedXMLValue = xmlRow[0][column.Ordinal].ToString().Trim().Replace(" ", "").Replace("\n", "").Replace("\r", "");
                                    if (!transformedXMLValue.Equals(transformedExcelValue))
                                        equal = false;
                                    break;
                                case 7:
                                    //Ignore uppercase or lowercase
                                    /*
                                     * Adjust the compared value ended with "\"
                                     * Advoid the problem like C:\abc vs C:\abc\
                                     */
                                    transformedExcelValue = row[column.Ordinal].ToString().ToLower();
                                    transformedXMLValue = xmlRow[0][column.Ordinal].ToString().ToLower();
                                    if (transformedExcelValue[transformedExcelValue.Length - 1] != '\\')
                                        transformedExcelValue = transformedExcelValue + "\\";
                                    if (transformedXMLValue[transformedXMLValue.Length - 1] != '\\')
                                        transformedXMLValue = transformedXMLValue + "\\";
                                    if (!transformedXMLValue.Equals(transformedExcelValue))
                                        equal = false;
                                    break;
                                case 8:
                                    // if excel value == "N/A" => no value can be found in XML
                                    if (row[column.Ordinal].ToString().Trim() == "N/A")
                                    {
                                        if (xmlRow[0][column.Ordinal].ToString().Trim() != "")
                                            equal = false;
                                    }
                                    else
                                    {
                                        //Ignore space
                                        transformedExcelValue = row[column.Ordinal].ToString().Trim().Replace(" ", "");
                                        if (!xmlRow[0][column.Ordinal].ToString().Replace(" ", "").Equals(transformedExcelValue))
                                            equal = false;
                                    }
                                    break;
                                case 11:
                                    // if excel value == "N/A" => no value can be found in XML
                                    if (row[column.Ordinal].ToString().Trim() == "N/A")
                                    {
                                        if (xmlRow[0][column.Ordinal].ToString().Trim() != "")
                                            equal = false;
                                    }
                                    else
                                    {
                                        //Ignore : and all letters
                                        transformedExcelValue = row[column.Ordinal].ToString().Replace(":", "");
                                        transformedExcelValue = Regex.Replace(transformedExcelValue, @"[A-Za-z]+", "").Trim();
                                        //In xml, time alwasys in format with four length => add 0 as start for those length below 4
                                        if (transformedExcelValue.Length == 3)
                                            transformedExcelValue = "0" + transformedExcelValue;
                                        transformedXMLValue = Regex.Replace(xmlRow[0][column.Ordinal].ToString().Replace(":", ""), @"[A-Za-z]+", "").Trim();
                                        if (!transformedXMLValue.Equals(transformedExcelValue))
                                            equal = false;
                                    }
                                    break;
                                case 12:
                                    //ignore new line(\n), uppercase or lowercase
                                    transformedExcelValue = row[column.Ordinal].ToString().ToLower().Replace("\n", "").Trim();
                                    transformedXMLValue = xmlRow[0][column.Ordinal].ToString().ToLower().Replace("\n", "").Trim();
                                    if (!transformedXMLValue.Equals(transformedExcelValue))
                                    {
                                        /*sometimes dummy text "error handling : " will be added 
                                         * (position of space is varied => replace them seperately and use Trim() 
                                         */
                                        transformedXMLValue = transformedXMLValue.Replace("error handling", "").Replace(":", "").Trim();
                                        if (!transformedXMLValue.Equals(transformedExcelValue))
                                            equal = false;
                                    }
                                    break;
                                default:
                                    //Ignore uppercase or lowercase
                                    if (!xmlRow[0][column.Ordinal].ToString().ToLower().Equals(row[column.Ordinal].ToString().ToLower()))
                                        equal = false;
                                    break;
                            }

                            if (!equal)
                            {
                                DataRow dr = dtDetail.NewRow();
                                dr[0] = row[3].ToString();
                                dr[1] = column.ColumnName;
                                dr[2] = row[column.Ordinal].ToString();
                                dr[3] = xmlRow[0][column.Ordinal].ToString();
                                dtDetail.Rows.Add(dr);
                            }
                        }
                    }
                }
            }


            //Handle Summary Content
            var details = dtDetail.AsEnumerable();
            var jobLists = (from r in details
                            select r["JobName"]).Distinct().ToList();
            foreach (var jobname in jobLists)
            {
                var invalidColumns = (from r in details
                                      where r[0] == jobname
                                      let outDesc = (r[1].ToString().Replace("\n", "").Length > 25 ? r[1].ToString().Replace("\n", "").Substring(0, 13) : r[1].ToString().Replace("\n", ""))
                                      select outDesc).ToList();
                DataRow dr = dtSummary.NewRow();
                dr[0] = jobname;
                dr[1] = String.Join(", ", invalidColumns);
                dtSummary.Rows.Add(dr);
            }

            ds.Tables.Add(dtDetail);
            ds.Tables.Add(dtSummary);
            return ds;
        }
    }
}
