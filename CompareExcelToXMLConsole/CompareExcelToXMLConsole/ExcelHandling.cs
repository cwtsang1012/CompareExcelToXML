using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CompareExcelToXML
{
    public class ExcelHandling
    {
        /// <summary>
        /// Use COM to read Excel
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns>A DataTable with excel data</returns>
        public DataTable GetExcelData(string excelFilePath)
        {
            Excel.Application app = new Excel.Application();
            Excel.Sheets sheets;
            Excel.Workbook workbook = null;
            //object oMissiong = System.Reflection.Missing.Value;
            DataTable dt = new DataTable();

            try
            {
                if (app == null)
                {
                    return null;
                }

                workbook = app.Workbooks.Open(excelFilePath);

                //Start to Read data into DataTable
                sheets = workbook.Worksheets;
                //Read the worksheet
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(3); //1st worksheet with index = 2
                if (worksheet == null)
                    return null;

                string cellContent;
                int iRowCount = worksheet.UsedRange.Rows.Count;
                int iColCount = worksheet.UsedRange.Columns.Count - 2; //default read two more column
                Excel.Range range;

                //Read the Column definition - start
                DataColumn dc;
                int ColumnID = 1;
                range = (Excel.Range)worksheet.Cells[1, 1];
                while (range.Text.ToString().Trim() != "")
                {
                    dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
                    dc.ColumnName = range.Text.ToString().Trim();
                    //Console.WriteLine(ColumnID + ": " + dc.ColumnName);
                    dt.Columns.Add(dc);

                    range = (Excel.Range)worksheet.Cells[1, ++ColumnID];
                }
                //Read the Column definition - end

                //Read the data - start
                for (int iRow = 2; iRow <= iRowCount; iRow++)
                {
                    DataRow dr = dt.NewRow();

                    for (int iCol = 1; iCol <= iColCount; iCol++)
                    {
                        range = (Excel.Range)worksheet.Cells[iRow, iCol];

                        cellContent = (range.Value2 == null) ? "" : range.Text.ToString();

                        //if (iRow == 1)
                        //{
                        //dt.Columns.Add(cellContent);
                        //}
                        //else
                        //{
                        //dr[iCol - 1] = cellContent;
                        //}

                        if (iCol == 4)
                            dr[iCol - 1] = cellContent.Replace("\n", "");
                        else
                            dr[iCol - 1] = cellContent;
                    }

                    if (dr.ItemArray[0].ToString() != "" || dr.ItemArray[3].ToString() != "")
                    {
                        dt.Rows.Add(dr);
                    }
                }
                //Read the data - end

                return dt;
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                workbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void ToExcelSheet(DataSet ds, string fileName)
        {
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbook workbookData = null;
            Excel.Worksheet worksheetData;
            Excel.Range range;
            try
            {
                workbookData = appExcel.Workbooks.Add(System.Reflection.Missing.Value);
                appExcel.DisplayAlerts = false;//不顯示警告
                //xlApp.Visible = true;//excel是否可見
                //
                //for (int i = workbookData.Worksheets.Count; i > 0; i--)
                //{
                //    Microsoft.Office.Interop.Excel.Worksheet oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbookData.Worksheets.get_Item(i);
                //    oWorksheet.Select();
                //    oWorksheet.Delete();
                //}

                for (int k = 0; k < ds.Tables.Count; k++)
                {
                    worksheetData = (Excel.Worksheet)workbookData.Worksheets.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    // testnum--;
                    if (ds.Tables[k] != null)
                    {
                        worksheetData.Name = ds.Tables[k].TableName;
                        //寫入標題
                        for (int i = 0; i < ds.Tables[k].Columns.Count; i++)
                        {
                            worksheetData.Cells[1, i + 1] = ds.Tables[k].Columns[i].ColumnName;
                            range = (Excel.Range)worksheetData.Cells[1, i + 1];
                            //range.Interior.ColorIndex = 15;
                            range.Font.Bold = true;
                            range.Interior.Color = Excel.XlRgbColor.rgbCadetBlue;
                            range.EntireColumn.NumberFormatLocal = "@";//文本格式 
                            range.ColumnWidth = 15;
                            //range.EntireColumn.AutoFit();//自動調整列寬 
                            //range.WrapText = true; //文本自動換行   
                            switch (i) 
                            {
                                case 0:
                                case 1:
                                    range.EntireColumn.AutoFit(); //自動調整列寬 
                                    break;
                                case 2:
                                case 3:
                                    range.EntireColumn.ColumnWidth = 67.14;
                                    range.EntireColumn.WrapText = true;
                                    break;
                            }
                        }


                        //寫入數值
                        for (int r = 0; r < ds.Tables[k].Rows.Count; r++)
                        {
                            for (int i = 0; i < ds.Tables[k].Columns.Count; i++)
                            {
                                worksheetData.Cells[r + 2, i + 1] = ds.Tables[k].Rows[r][i];
                            }
                        }
                    }
                    worksheetData.Columns.EntireColumn.AutoFit();
                    workbookData.Saved = true;
                }
            }
            catch (Exception ex) { Console.WriteLine("Error: " + ex.ToString()); }
            finally
            {
                if (ConfigurationManager.AppSettings["SectionName"] != "")
                    fileName = fileName.Replace(".xlsx", "_" + ConfigurationManager.AppSettings["SectionName"] + ".xlsx");
                workbookData.SaveCopyAs(fileName);
                workbookData.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                appExcel.Quit();
                GC.Collect();
            }
        }

        /*
            /// <summary>
        /// 使用COM，多線程讀取Excel（1 主線程、4 副線程）
        /// </summary>
        /// <param name="excelFilePath">路徑</param>
        /// <returns>DataTabel</returns>
        public System.Data.DataTable ThreadReadExcel(string excelFilePath)
        {
            Excel.Application app = new Excel.Application();
            Excel.Sheets sheets = null;
            Excel.Workbook workbook = null;
            object oMissiong = System.Reflection.Missing.Value;
            System.Data.DataTable dt = new System.Data.DataTable();
 
            wath.Start();
 
            try
            {
                if (app == null)
                {
                    return null;
                }
 
                workbook = app.Workbooks.Open(excelFilePath, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
 
                //將數據讀入到DataTable中——Start   
                sheets = workbook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);//讀取第一張表
                if (worksheet == null)
                    return null;
 
                string cellContent;
                int iRowCount = worksheet.UsedRange.Rows.Count;
                int iColCount = worksheet.UsedRange.Columns.Count;
                Excel.Range range;
 
                //負責列頭Start
                DataColumn dc;
                int ColumnID = 1;
                range = (Excel.Range)worksheet.Cells[1, 1];
                //while (range.Text.ToString().Trim() != "")
                while (iColCount >= ColumnID)
                {
                    dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
 
                    string strNewColumnName = range.Text.ToString().Trim();
                    if (strNewColumnName.Length == 0) strNewColumnName = "_1";
                    //判斷列名是否重複
                    for (int i = 1; i < ColumnID; i++)
                    {
                        if (dt.Columns[i - 1].ColumnName == strNewColumnName)
                            strNewColumnName = strNewColumnName + "_1";
                    }
 
                    dc.ColumnName = strNewColumnName;
                    dt.Columns.Add(dc);
 
                    range = (Excel.Range)worksheet.Cells[1, ++ColumnID];
                }
                //End
 
                //數據大於500條，使用多進程進行讀取數據
                if (iRowCount - 1 > 500)
                {
                    //開始多線程讀取數據
                    //新建線程
                    int b2 = (iRowCount - 1) / 10;
                    DataTable dt1 = new DataTable("dt1");
                    dt1 = dt.Clone();
                    SheetOptions sheet1thread = new SheetOptions(worksheet, iColCount, 2, b2 + 1, dt1);
                    Thread othread1 = new Thread(new ThreadStart(sheet1thread.SheetToDataTable));
                    othread1.Start();
 
                    //阻塞 1 毫秒，保證第一個讀取 dt1
                    Thread.Sleep(1);
 
                    DataTable dt2 = new DataTable("dt2");
                    dt2 = dt.Clone();
                    SheetOptions sheet2thread = new SheetOptions(worksheet, iColCount, b2 + 2, b2 * 2 + 1, dt2);
                    Thread othread2 = new Thread(new ThreadStart(sheet2thread.SheetToDataTable));
                    othread2.Start();
 
                    DataTable dt3 = new DataTable("dt3");
                    dt3 = dt.Clone();
                    SheetOptions sheet3thread = new SheetOptions(worksheet, iColCount, b2 * 2 + 2, b2 * 3 + 1, dt3);
                    Thread othread3 = new Thread(new ThreadStart(sheet3thread.SheetToDataTable));
                    othread3.Start();
 
                    DataTable dt4 = new DataTable("dt4");
                    dt4 = dt.Clone();
                    SheetOptions sheet4thread = new SheetOptions(worksheet, iColCount, b2 * 3 + 2, b2 * 4 + 1, dt4);
                    Thread othread4 = new Thread(new ThreadStart(sheet4thread.SheetToDataTable));
                    othread4.Start();
 
                    //主線程讀取剩餘數據
                    for (int iRow = b2 * 4 + 2; iRow <= iRowCount; iRow++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int iCol = 1; iCol <= iColCount; iCol++)
                        {
                            range = (Excel.Range)worksheet.Cells[iRow, iCol];
                            cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                            dr[iCol - 1] = cellContent;
                        }
                        dt.Rows.Add(dr);
                    }
 
                    othread1.Join();
                    othread2.Join();
                    othread3.Join();
                    othread4.Join();
 
                    //將多個線程讀取出來的數據追加至 dt1 後面
                    foreach (DataRow dr in dt.Rows)
                        dt1.Rows.Add(dr.ItemArray);
                    dt.Clear();
                    dt.Dispose();
 
                    foreach (DataRow dr in dt2.Rows)
                        dt1.Rows.Add(dr.ItemArray);
                    dt2.Clear();
                    dt2.Dispose();
 
                    foreach (DataRow dr in dt3.Rows)
                        dt1.Rows.Add(dr.ItemArray);
                    dt3.Clear();
                    dt3.Dispose();
 
                    foreach (DataRow dr in dt4.Rows)
                        dt1.Rows.Add(dr.ItemArray);
                    dt4.Clear();
                    dt4.Dispose();
 
                    return dt1;
                }
                else
                {
                    for (int iRow = 2; iRow <= iRowCount; iRow++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int iCol = 1; iCol <= iColCount; iCol++)
                        {
                            range = (Excel.Range)worksheet.Cells[iRow, iCol];
                            cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                            dr[iCol - 1] = cellContent;
                        }
                        dt.Rows.Add(dr);
                    }
                }
 
                wath.Stop();
                TimeSpan ts = wath.Elapsed;
                //將數據讀入到DataTable中——End
                return dt;
            }
            catch
            {
 
                return null;
            }
            finally
            {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                 
                
                //object objmissing = System.Reflection.Missing.Value;
 
                //Excel.ApplicationClass application = new ApplicationClass();
                //Excel.Workbook book = application.Workbooks.Add(objmissing);
                //Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.Add（objmissing,objmissing,objmissing,objmissing);
 
                //操作過程 ^&%&×&……&%&&……
 
                //釋放
                //sheet.SaveAs(path,objmissing,objmissing,objmissing,objmissing,objmissing,objmissing,objmissing,objmissing);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)sheet);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)book);
                //application.Quit();
                //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)application);
                //System.GC.Collect();
            }
        }
 
 
        /// <summary>
        /// 刪除Excel行
        /// </summary>
        /// <param name="excelFilePath">Excel路徑</param>
        /// <param name="rowStart">開始行</param>
        /// <param name="rowEnd">結束行</param>
        /// <param name="designationRow">指定行</param>
        /// <returns></returns>
        public string DeleteRows(string excelFilePath, int rowStart, int rowEnd, int designationRow)
        {
            string result = "";
            Excel.Application app = new Excel.Application();
            Excel.Sheets sheets;
            Excel.Workbook workbook = null;
            object oMissiong = System.Reflection.Missing.Value;
            try
            {
                if (app == null)
                {
                    return "分段讀取Excel失敗";
                }
 
                workbook = app.Workbooks.Open(excelFilePath, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                sheets = workbook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);//讀取第一張表
                if (worksheet == null)
                    return result;
                Excel.Range range;
 
                //先刪除指定行，一般為列描述
                if (designationRow != -1)
                {
                    range = (Excel.Range)worksheet.Rows[designationRow, oMissiong];
                    range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }
                Stopwatch sw = new Stopwatch();
                sw.Start();
 
                int i = rowStart;
                for (int iRow = rowStart; iRow <= rowEnd; iRow++, i++)
                {
                    range = (Excel.Range)worksheet.Rows[rowStart, oMissiong];
                    range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }
 
                sw.Stop();
                TimeSpan ts = sw.Elapsed;
                workbook.Save();
 
                //將數據讀入到DataTable中——End
                return result;
            }
            catch
            {
 
                return "分段讀取Excel失敗";
            }
            finally
            {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
 
        public void ToExcelSheet(DataSet ds, string fileName)
        {
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbook workbookData = null;
            Excel.Worksheet worksheetData;
            Excel.Range range;
            try
            {
                workbookData = appExcel.Workbooks.Add(System.Reflection.Missing.Value);
                appExcel.DisplayAlerts = false;//不顯示警告
                //xlApp.Visible = true;//excel是否可見
                //
                //for (int i = workbookData.Worksheets.Count; i > 0; i--)
                //{
                //    Microsoft.Office.Interop.Excel.Worksheet oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbookData.Worksheets.get_Item(i);
                //    oWorksheet.Select();
                //    oWorksheet.Delete();
                //}
 
                for (int k = 0; k < ds.Tables.Count; k++)
                {
                    worksheetData = (Excel.Worksheet)workbookData.Worksheets.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    // testnum--;
                    if (ds.Tables[k] != null)
                    {
                        worksheetData.Name = ds.Tables[k].TableName;
                        //寫入標題
                        for (int i = 0; i < ds.Tables[k].Columns.Count; i++)
                        {
                            worksheetData.Cells[1, i + 1] = ds.Tables[k].Columns[i].ColumnName;
                            range = (Excel.Range)worksheetData.Cells[1, i + 1];
                            //range.Interior.ColorIndex = 15;
                            range.Font.Bold = true;
                            range.NumberFormatLocal = "@";//文本格式 
                            range.EntireColumn.AutoFit();//自動調整列寬 
                            // range.WrapText = true; //文本自動換行   
                            range.ColumnWidth = 15;
                        }
                        //寫入數值
                        for (int r = 0; r < ds.Tables[k].Rows.Count; r++)
                        {
                            for (int i = 0; i < ds.Tables[k].Columns.Count; i++)
                            {
                                worksheetData.Cells[r + 2, i + 1] = ds.Tables[k].Rows[r][i];
                                //Range myrange = worksheetData.get_Range(worksheetData.Cells[r + 2, i + 1], worksheetData.Cells[r + 3, i + 2]);
                                //myrange.NumberFormatLocal = "@";//文本格式 
                                //// myrange.EntireColumn.AutoFit();//自動調整列寬 
                                ////   myrange.WrapText = true; //文本自動換行   
                                //myrange.ColumnWidth = 15;
                            }
                            //  rowRead++;
                            //System.Windows.Forms.Application.DoEvents();
                        }
                    }
                    worksheetData.Columns.EntireColumn.AutoFit();
                    workbookData.Saved = true;
                }
            }
            catch (Exception ex) { }
            finally
            {
                workbookData.SaveCopyAs(fileName);
                workbookData.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                appExcel.Quit();
                GC.Collect();
            }
          
        */
    }
}
