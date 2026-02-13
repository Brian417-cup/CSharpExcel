using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;

namespace CSharpForExcel
{
    class SimpleExcelTool
    {
        //执行的SQL语句
        private string excelSQL = "";
        //返回结果数据集
        private DataSet queryDataset = null;
        //连接驱动
        string strConn = "";
        //连接
        OleDbConnection oledbConn = null;
        //适配
        OleDbDataAdapter oleDbDataAdapter = null;

        public DataSet QueryDataset
        {
            get { return queryDataset; }
            set { queryDataset = value; }
        }

        public string ExcelSQL
        {
            get { return excelSQL; }
            set { excelSQL = value; }
        }

        //打开连接
        public bool openConnection(string filePath)
        {
            try
            {
                //这种版本的连接需要装Office客户端
                // strConn = string.Format("Provider={0};Data Source={1}; Extended Properties={2};",
                //new object[] { "Microsoft.ACE.OLEDB.12.0", filePath, "'Excel 12.0'" });
                //这种版本的不需要装Office客户端，可直接使用，但是只支持excel2003
                strConn = string.Format("Provider={0};Data Source={1}; Extended Properties={2};",
               new object[] { "Microsoft.Jet.OLEDB.4.0", filePath, "'Excel 8.0'" });

                oledbConn = new OleDbConnection(strConn);
                oledbConn.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        //清空上次数据
        public bool clearLastQueryResult()
        {
            try
            {
                if (queryDataset != null)
                {
                    queryDataset.Clear();
                    queryDataset = new DataSet();
                }

                if (oleDbDataAdapter != null)
                {
                    oleDbDataAdapter.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

            return true;

        }

        //关闭连接
        public bool closeConnection()
        {
            try
            {
                if (oledbConn != null)
                {
                    oledbConn.Dispose();
                    oledbConn.Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        //执行类SQL查询语句
        //执行的前提:连接已经建立且被打开了，上一次的查询结果已经被清空了
        public DataSet executeExcelQuerySQL(string sqlStr)
        {
            if (oledbConn == null)
            {
                return null;
            }

            try
            {
                queryDataset = new DataSet();
                //数据适配器查询
                oleDbDataAdapter = new OleDbDataAdapter(sqlStr, strConn);
                //查询结果填充到结果集中
                oleDbDataAdapter.Fill(queryDataset);

                return queryDataset;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        //执行类SQL查询语句
        //执行的前提:连接已经建立且被打开了，上一次的查询结果已经被清空了
        public int executeExcelInsertAndDeleteSQL(string sqlStr)
        {
            if (oledbConn == null)
            {
                return -1;
            }

            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = oledbConn;
                cmd.CommandText = sqlStr;
                return cmd.ExecuteNonQuery();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -1;
            }
        }

        //导出到Excel中
        public bool ExportExcels(string saveFileName, DataGridView myDGV)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("无法创建Excel对象，可能您的机子未安装Excel");
                    return false;
                }
                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
                //写入标题
                for (int i = 0; i < myDGV.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = myDGV.Columns[i].HeaderText;
                }
                //写入数值
                for (int r = 0; r < myDGV.Rows.Count; r++)
                {
                    for (int i = 0; i < myDGV.ColumnCount; i++)
                    {
                        worksheet.Cells[r + 2, i + 1] = myDGV.Rows[r].Cells[i].Value;
                    }
                    System.Windows.Forms.Application.DoEvents();
                }
                worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
                if (saveFileName != "")
                {
                    try
                    {
                        workbook.Saved = true;
                        workbook.SaveCopyAs(saveFileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                    }
                }
                xlApp.Quit();
                
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex);
                return false;
            }

            return true;
        }

        //导出到csv中
        public bool ExportCSV(Stream targetStream,string saveFileName, DataGridView myDGV)
        {
            try
            {
                Stream stream = targetStream;
                StreamWriter writer = new StreamWriter(stream, System.Text.Encoding.GetEncoding(-0));
                string strLine = "";

                //表标题
                for (int i = 0; i < myDGV.ColumnCount;i++ )
                {
                    if (i>0)
                    {
                        strLine += ",";
                    }
                    strLine += myDGV.Columns[i].HeaderText;
                }

                strLine.Remove(strLine.Length - 1);
                writer.WriteLine(strLine);

                //表内容
                for (int i = 0; i < myDGV.Rows.Count;i++ )
                {
                    strLine = "";
                    int colCnt = myDGV.Columns.Count;
                    for (int j = 0; j < colCnt;j++ )
                    {
                        if (j>0&&j<colCnt)
                        {
                            strLine += ",";
                        }
                        if (myDGV.Rows[i].Cells[j].Value==null)
                        {
                            strLine += "";
                        }
                        else
                        {
                            string cellData = myDGV.Rows[i].Cells[j].Value.ToString();
                            //防止特殊符号
                            cellData = cellData.Replace("\"", "\"\"");
                            strLine += cellData;
                        }
                    }
                    writer.WriteLine(strLine);
                }

                writer.Close();
                stream.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                Console.WriteLine(ex);
                return false;
            }

            return true;
        }
    }

}
