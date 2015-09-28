using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Data;
using System.Windows.Forms;

namespace ReactiveControlAnalysis
{
    class CReadWordFile
    {
        private string fileName;
        private Microsoft.Office.Interop.Word.Application cls = null;
        private Microsoft.Office.Interop.Word.Document doc = null;
        private Microsoft.Office.Interop.Word.Table table = null;
        private object missing = Missing.Value;
        //Word是否处于打开状态
        private bool openState;

        public CReadWordFile(string fileName)
        {
            this.fileName = fileName;
        }

        
       /// <summary>
       /// 打开Word文档
       /// </summary>
        public void Open()
        {
            object path = this.fileName;
            cls = new Microsoft.Office.Interop.Word.Application();
            cls.Visible = true;
            //object missing = System.Reflection.Missing.Value;
            //object strNewPath = "E:\\堆芯\\附件4-CoreEasy输出参数表.docx";
            try
            {
                doc = cls.Documents.Open(ref path, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                openState = true;
            }
            catch
            {
                openState = false;
                MessageBox.Show("打开< "+ path +" >失败！","提示");
            }
        }
        /// <summary>
        /// 返回指定单元格中的数据
        /// </summary>
        /// <param name="表的序号"></param>
        /// <param name="行号"></param>
        /// <param name="第几列"></param>
        /// <returns></returns>
        public string ReadWord(int tableIndex, int rowIndex, int colIndex)
        {
            //Give the value to the tow Int32 params.
            try
            {
                if (openState == true)
                {
                    table = doc.Tables[tableIndex];
                    string text = table.Cell(rowIndex, colIndex).Range.Text.ToString();
                    text = text.Substring(0, text.Length - 2);    //去除尾部的mark
                    return text;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ee)
            {
                return ee.ToString();
            }
        }
        /// <summary>
        /// 向指定单元格中写值
        /// </summary>
        /// <param name="表的序号"></param>
        /// <param name="行号"></param>
        /// <param name="第几列"></param>
        /// <param name="要写入的值"></param>
        /// <returns></returns>
        public bool WriteToWord(int tableIndex, int rowIndex, int colIndex, string text)
        {
            //Give the value to the tow Int32 params.
            try
            {
                if (openState == true)
                {
                    table = doc.Tables[tableIndex];
                    table.Cell(rowIndex, colIndex).Range.Text = text;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        public DataTable WordTable(int tableIndex)
        {
            DataTable dt = new DataTable();
            if (openState == true)
            {
                table = doc.Tables[tableIndex];
                for (int ii = 0; ii < table.Columns.Count; ii++)
                {
                    dt.Columns.Add("cl" + ii.ToString(), typeof(string));
                }
                for (int ii = 0; ii < table.Rows.Count; ii++)
                {
                    DataRow rw = dt.NewRow();
                    for (int jj = 0; jj < table.Columns.Count; jj++)
                    {
                        string text = table.Cell(ii + 1, jj + 1).Range.Text.ToString();
                        //string text = table.Rows[ii].Cells[jj].ToString();
                        text = text.Substring(0, text.Length - 2);
                        rw["cl" + (jj).ToString()] = text;
                    }
                    dt.Rows.Add(rw);
                }
            }
            return dt;
        }
        /// <summary>
        /// 关闭Word文档
        /// </summary>
        public void Close()
        {
            if (openState == true)
            {
                if (doc != null)
                    doc.Close(ref missing, ref missing, ref missing);
                cls.Quit(ref missing, ref missing, ref missing);
            }
        }
        /// <summary>
        /// 保存文档
        /// </summary>
        public void Save()
        {
            if (openState == true)
            {
                if (doc != null)
                    doc.Save();
            }
        }


      
    }
}
