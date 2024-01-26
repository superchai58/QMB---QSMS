using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Excel=Microsoft.Office.Interop.Excel;
using System.Media;
using QuantaSDK.Excel;

namespace QSMS.BrLibrary
{
    class PublicFunction
    {
        #region  保证窗体唯一方法
        public void HaveOpened(Form frm, string FormName)
        {
            bool hasform = false;
            //遍历所有窗体对象,判断窗体是否已经弹出
            foreach (string f in Parameter.Openforms)
            {
                //判断弹出的窗体是否重复
                if (f == FormName)
                {
                    hasform = true;
                }
            }
            if (hasform)
            {
                frm.Close();
            }
            else
            {
                //添加到所有窗体中
                Parameter.Openforms.Add(FormName);
                //并打开该窗体
                frm.Show();
            }
        }
        public void RemoveForm(string FormName)
        {
            if (Parameter.Openforms.Contains(FormName))
            {
                Parameter.Openforms.Remove(FormName);
            }
        }
        #endregion

        public string[] GetExcelSheetName(string Path)
        {
            //连接串  
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Path + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";

            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            //返回Excel的架构，包括各个sheet表的名称,类型，创建时间和修改时间等   
            DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            //包含excel中表名的字符串数组  
            string[] strTableNames = new string[dtSheetName.Rows.Count];

            for (int k = 0; k < dtSheetName.Rows.Count; k++)
            {
                string strSheetTableName = dtSheetName.Rows[k]["TABLE_NAME"].ToString();
                if (strSheetTableName.Contains("$") && strSheetTableName.Replace("'", "").EndsWith("$"))
                {
                    strTableNames[k] = strSheetTableName.Substring(0, strSheetTableName.Length - 1);
                }
                else
                {
                    strTableNames[k] = dtSheetName.Rows[k]["TABLE_NAME"].ToString();
                }
                
            }
            conn.Close();
            return strTableNames;
        }

        public void BindComboBox(ComboBox ddl, string[] sheet)
        {
            ddl.Items.Clear();
            for (int i = 0; i < sheet.Length; i++)
            {
                ddl.Items.Add(sheet[i]);
            }
        }

        public DataTable GetDataFromExcel(string Path, string Sheet)
        {
            //连接串  
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            string strConn = "";

            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Path + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";


            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();

            OleDbDataAdapter myCommand = null;
            DataTable dt = new DataTable();
            //从指定的表明查询数据,可先把所有表明列出来供用户选择  
            string strExcel = "select * from [" + Sheet + "$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            dt = new DataTable();
            myCommand.Fill(dt);
            conn.Close();
            return dt;
        }

        public DataTable ReadFromExcel(string sXlsxPath, string sSheetName)
        {
            string sExt = System.IO.Path.GetExtension(sXlsxPath);
            string sConn = "";
            if (sExt == ".xlsx") //Excel2007
            {
                sConn =
                     "Provider=Microsoft.ACE.OLEDB.12.0;" +
                     "Data Source=" + sXlsxPath + ";" +
                     "Extended Properties='Excel 12.0;HDR=YES'";
            }
            else if (sExt == ".xls") //Excel2003
            {
                sConn =
                    "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + sXlsxPath + ";" +
                    "Extended Properties='Excel 8.0;HDR=YES'";
            }
            else
            {
                throw new Exception("未知的文件类型");
            }
            OleDbConnection oledbConn = new OleDbConnection(sConn);
            oledbConn.Open();
            OleDbDataAdapter command = new OleDbDataAdapter(
                "SELECT * FROM [" + sSheetName + "]", oledbConn);
            DataSet ds = new DataSet();
            command.Fill(ds, sSheetName);
            oledbConn.Close();
            return ds.Tables[sSheetName];
        }

        public bool CopyToExcel(DataGridView dgv, string SheetName, bool isShowExcle)
        {
            if (dgv.Rows.Count == 0)
                return false;
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range xlRang;
            oXL = new Excel.Application();
            oXL.Visible = isShowExcle;
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            for (int i = 0; i < dgv.ColumnCount; i++)
            {
                oSheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
            }
            for (int i = 0; i < dgv.RowCount - 1; i++)
            {
                for (int j = 0; j < dgv.ColumnCount; j++)
                {
                    if (dgv[j, i].ValueType == typeof(string))
                    {
                        oSheet.Cells[i + 2, j + 1] = "'" + dgv[j, i].Value.ToString();
                    }
                    else
                    {
                        oSheet.Cells[i + 2, j + 1] = dgv[j, i].Value.ToString();
                    }
                }
            }

            oSheet.Name = SheetName;
            xlRang = oXL.Columns;
            xlRang.EntireColumn.AutoFit();
            xlRang.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            oXL.Visible = true;
            return true;

            //建立Excel对象    
            /*  Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

           
              excel.Application.Workbooks.Add(true);
              excel.Visible = isShowExcle;
              //生成字段名称    
              for (int i = 0; i < dgv.ColumnCount; i++)
              {
                  excel.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
              }
              //填充数据    
              for (int i = 0; i < dgv.RowCount  ; i++)
              {
                  for (int j = 0; j < dgv.ColumnCount; j++)
                  {
                      if (dgv[j, i].ValueType == typeof(string))
                      {
                         // excel.Cells[i + 2, j + 1] = "'" + dgv[j, i].Value.ToString();
                          excel.Cells[i + 2, j + 1] = dgv[j, i].Value.ToString();
                      }
                      else
                      {
                          excel.Cells[i + 2, j + 1] = dgv[j, i].Value.ToString();
                      }
                  }
              }
            
              excel.Visible = true;
              return true;*/
        }

        public void ExportDataSetToExcel(DataSet ds, string[] names)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            int m = 0;
            foreach (System.Data.DataTable table in ds.Tables)
            {
                Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;  //sheet名为datatable表名
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }

                //这里可以添加Sheet的格式代码
                excelWorkSheet.Name = names[m];
                Excel.Range xlRang = excelWorkSheet.Columns;
                xlRang.EntireColumn.AutoFit();
                xlRang.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                m++;
            }

            excelApp.Visible = true;

        }
        //Luck Add 添加DataSet类型
        public void doExport(DataSet ds)    
        {
            if (ds.Tables[0].Rows.Count == 0)
            {
                return;
            }
            ExcelHelper eh = new ExcelHelper();
            eh.CreateDocument();
            eh.Export(ds, 0, 0, true);
            eh.Quit(true, false);
        }

        public void doExport(DataTable dt)
        {
            if (dt == null)
            {
                return;
            }
            ExcelHelper eh = new ExcelHelper();
            eh.CreateDocument();            
            eh.Export(dt, 0, 0, true);
            eh.Quit();
        }
        //Viter Add 添加保存路径
        public void doExportSave(DataTable dt)
        {
            string path = "";
            System.Windows.Forms.SaveFileDialog fbd = new System.Windows.Forms.SaveFileDialog();
            fbd.Filter = "Excel|*.xlsx;*.xls;";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.FileName;
                if (path == "")
                {
                    return;
                }
            }
            if (dt == null)
            {
                return;
            }
            ExcelHelper eh = new ExcelHelper();
            eh.CreateDocument();
            eh.Export(dt, 0, 0, true);
            eh.Save(path);
            eh.Dispose();
        }



        #region  获取Excel中指定位置的值；
        //public string getExcelOneCell(string fileName, int row, int column)
        //{
        //    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(fileName, Type.Missing, Type.Missing,
        //         Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //         Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //         Type.Missing, Type.Missing);

        //    Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[1];

        //    string temp = ((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[row, column]).Text.ToString();

        //    wbook.Close(false, fileName, false);
        //    app.Quit();
        //    NAR(app);
        //    NAR(wbook);
        //    NAR(workSheet);
        //    return temp;
        //}

        private void NAR(Object o)
        {
            try
            {
                //使用此方法，来释放引用某些资源的基础 COM 对象。 这里的o就是要释放的对象
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null; GC.Collect();
            }
        }
        #endregion  
        
        public string ConfigListGetValue(string item)
        {
            if (Parameter.ConfigList.ContainsKey(item.ToUpper()) == false)
            {
                return "";
            }
            return Parameter.ConfigList[item.ToUpper()].ToString().ToUpper();
        }
        //Paul Add
        public void Sound(string strStyle)
        {
            SoundPlayer player = new SoundPlayer();
            if (strStyle.ToUpper().Trim() == "OK")
            {
                player.SoundLocation = Application.StartupPath.Trim() + @"\Sound\OK.wav";
            }
            else if (strStyle.ToUpper().Trim() == "ERROR")
            {
                player.SoundLocation = Application.StartupPath.Trim() + @"\Sound\OO.wav";
            }
            player.Play();
        }
        //Paul Add
        public bool IsNumeric(string message, string Type)
        {
            try
            {
                if (Type == "INT")
                {
                    int result = Convert.ToInt32(message);//整数
                }
                else if (Type == "DOUBLE")
                {
                    double result = double.Parse(message);//数字
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public string DIDGetRefIDByResult(string strResult)
        {
            int iPos, Jpos;
            iPos = strResult.IndexOf(":");
            Jpos = strResult.IndexOf(",");
            if (iPos < 0 || Jpos < 9)
            {
                return "";
            }
            //return strResult.Substring(iPos, Jpos - iPos - 2);
            return strResult.Substring(iPos + 1, Jpos - iPos - 1);  ///202201051341   Rain
        }

        public string ReadIniFile(string DB, string item, string fileName)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + fileName;
            return QMSSDK.Br.FileSystem.Ini.ReadIniValue(DB, item, fileName);
        }
    }
}










