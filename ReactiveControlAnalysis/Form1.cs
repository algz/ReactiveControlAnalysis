using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraEditors.Repository;



namespace ReactiveControlAnalysis
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        private string strPath3;   //附件3的路径
        private string strPath4;    //附件4的路径
        private string strPath5;    //附件5的路径
        private int tableNumber3 = 2;      //附件3中的第几个表
        private int tableNumber41 = 4;      //附件4中的第几个表
        private int tableNumber42 = 2;
        private int tableNumber5 = 1;        //附件5中的第几个表


        MyGridPainter painter;
        BandedGridView _view;
        List<MyMergedCell> _MergedCells = new List<MyMergedCell>();

        public Form1()
        {
            string workPath = GetWorkPath();
            strPath3 = workPath + "\\附件3-反应性控制需求表.docx";
            strPath4 = workPath + "\\附件4-CoreEasy输出参数表.docx";
            strPath5 = workPath + "\\附件5-反应性控制分析计算.docx";
            InitializeComponent();
            LoadView();


            _view = this.bandedGridView1;

            //AddMergedCell(7, this.bandedGridView1.Columns["Value1"], bandedGridView1.Columns["Error1"]);
            //AddMergedCell(7, this.bandedGridView1.Columns["Value2"], bandedGridView1.Columns["Error2"]);
            //AddMergedCell(7, this.bandedGridView1.Columns["Value3"], bandedGridView1.Columns["Error3"]);
            //AddMergedCell(9, this.bandedGridView1.Columns["Value1"], bandedGridView1.Columns["Error1"]);
            //AddMergedCell(9, this.bandedGridView1.Columns["Value2"], bandedGridView1.Columns["Error2"]);
            //AddMergedCell(9, this.bandedGridView1.Columns["Value3"], bandedGridView1.Columns["Error3"]);
            //AddMergedCell(10, this.bandedGridView1.Columns["Value1"], bandedGridView1.Columns["Error1"]);
            //AddMergedCell(10, this.bandedGridView1.Columns["Value2"], bandedGridView1.Columns["Error2"]);
            //AddMergedCell(10, this.bandedGridView1.Columns["Value3"], bandedGridView1.Columns["Error3"]);
            //AddMergedCell(11, this.bandedGridView1.Columns["Value1"], bandedGridView1.Columns["Error1"]);
            //AddMergedCell(11, this.bandedGridView1.Columns["Value2"], bandedGridView1.Columns["Error2"]);
            //AddMergedCell(11, this.bandedGridView1.Columns["Value3"], bandedGridView1.Columns["Error3"]);

            //AddMergedCell(7, this.bandedGridView1.Columns["Error2"], bandedGridView1.Columns["Value2"]);
            //AddMergedCell(7, this.bandedGridView1.Columns["Error3"], bandedGridView1.Columns["Value1"]);

            //object va = 12345678;
            //AddMergedCell(7, 1, 2, va);
            painter = new MyGridPainter(this.bandedGridView1);
            Jisuan();
        }


        private void LoadView()
        {
            //BandedGridView view = bandedGridView1  as BandedGridView;
            List<string> ValueFrom4 = GetValueFrom4();
            List<string> ValueFrom3 = GetValueFrom3();
            if (ValueFrom4 != null)
            {
                bandedGridView1.BeginUpdate(); //开始视图的编辑，防止触发其他事件
                bandedGridView1.BeginDataUpdate(); //开始数据的编辑

                bandedGridView1.Bands.Clear();

                #region 修改附加选项
                bandedGridView1.OptionsView.ShowColumnHeaders = false;                         //因为有Band列了，所以把ColumnHeader隐藏
                bandedGridView1.OptionsView.ShowGroupPanel = false;                            //如果没必要分组，就把它去掉
                bandedGridView1.OptionsView.EnableAppearanceEvenRow = false;                   //是否启用偶数行外观
                bandedGridView1.OptionsView.EnableAppearanceOddRow = true;                     //是否启用奇数行外观
                bandedGridView1.OptionsView.ShowFilterPanelMode = ShowFilterPanelMode.Never;   //是否显示过滤面板
                bandedGridView1.OptionsCustomization.AllowColumnMoving = false;                //是否允许移动列
                bandedGridView1.OptionsCustomization.AllowColumnResizing = false;              //是否允许调整列宽
                bandedGridView1.OptionsCustomization.AllowGroup = false;                       //是否允许分组
                bandedGridView1.OptionsCustomization.AllowFilter = false;                      //是否允许过滤
                bandedGridView1.OptionsCustomization.AllowSort = true;                         //是否允许排序
                bandedGridView1.OptionsSelection.EnableAppearanceFocusedCell = true;           //???
                bandedGridView1.OptionsBehavior.Editable = true;                               //是否允许用户编辑单元格
                //bandedGridView1.OptionsView.AllowCellMerge = true;
                #endregion

                #region  添加列标题
                GridBand bandID = bandedGridView1.Bands.AddBand("ID");
                bandID.Visible = false; //隐藏ID列
                GridBand bandName = bandedGridView1.Bands.AddBand("需求项");
                GridBand bandSystem1 = bandedGridView1.Bands.AddBand("第一系统/pcm");
                GridBand bandSystem2 = bandedGridView1.Bands.AddBand("第二系统/pcm");
                GridBand bandSystem3 = bandedGridView1.Bands.AddBand("非能动棒/pcm");
                GridBand bandValue1 = bandSystem1.Children.AddBand("计算值");
                GridBand bandError1 = bandSystem1.Children.AddBand("考虑的计算误差");
                GridBand bandValue2 = bandSystem2.Children.AddBand("计算值");
                GridBand bandError2 = bandSystem2.Children.AddBand("考虑的计算误差");
                GridBand bandValue3 = bandSystem3.Children.AddBand("计算值");
                GridBand bandError3 = bandSystem3.Children.AddBand("考虑的计算误差");
                GridBand bandErrer = bandedGridView1.Bands.AddBand("最大计算误差");
                #endregion
                #region 列标题对齐方式
                bandName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandSystem1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandSystem2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandSystem3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandValue1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandError1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandValue2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandError2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandValue3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                bandError3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                #endregion


                List<Idata> datas = new List<Idata>();
                datas.Add(new Idata(1, "温度效应(冷态到热态的反应性亏损)", ValueFrom4[0], "", "", "", "0", "0", "15%"));
                datas.Add(new Idata(2, "功率效应", ValueFrom4[1], "", "", "", "", "", "25%"));
                datas.Add(new Idata(3, "燃料循环后备反应性", ValueFrom4[2], "", "0", "0", "0", "0", "10%"));
                datas.Add(new Idata(4, "超功率效应", (Convert.ToDouble(ValueFrom4[1]) * 0.1).ToString(), "0", "", "0", "0", "0", ""));
                datas.Add(new Idata(5, "反应性引入事故", (Convert.ToDouble(ValueFrom4[3]) * 0.5).ToString(), "", "", "", "0", "0", "15%"));
                datas.Add(new Idata(6, "临界和裂变材料装量不确定性相关的反应性", ValueFrom3[0], "0", "0", "0", "0", "0", ""));
                double total6 = SubTotal(datas);
                datas.Add(new Idata(7, "停堆深度要求", ValueFrom3[1], "0", "1000", "0", "0", "0", ""));

                double total8 = SubTotal(datas);
                datas.Add(new Idata(8, "合计（未考虑控制棒计算误差）", total8.ToString(), "", "", "", "", "", "15%"));
                datas.Add(new Idata(9, "卡棒准则下停堆深度要求", ValueFrom3[2], "0", ValueFrom3[2], "0", "空", "空", ""));
                double total10 = total6 + Convert.ToDouble(ValueFrom3[2]);
                datas.Add(new Idata(10, "卡棒准则下合计（未考虑控制棒计算误差）", total10.ToString(), "", "", "", "空", "空", "15%"));
                datas.Add(new Idata(11, "反应性价值最低要求（考虑控制棒计算误差）", (total8 / (1 - 0.15)).ToString(), (total8 / (1 - 0.15)).ToString(), "", "", "", "", ""));
                datas.Add(new Idata(12, "卡棒准则下反应性价值最低要求（考虑控制棒计算误差）", (total10 / (1 - 0.15)).ToString(), (total10 / (1 - 0.15)).ToString(), "", "", "空", "空", ""));
                this.gridControl1.DataSource = datas;
                this.gridControl1.MainView.PopulateColumns();


                #region 将标题列和数据列对应
                bandedGridView1.Columns["Id"].OwnerBand = bandID;
                bandedGridView1.Columns["Name"].OwnerBand = bandName;
                bandedGridView1.Columns["Value1"].OwnerBand = bandValue1;
                bandedGridView1.Columns["Error1"].OwnerBand = bandError1;
                bandedGridView1.Columns["Value2"].OwnerBand = bandValue2;
                bandedGridView1.Columns["Error2"].OwnerBand = bandError2;
                bandedGridView1.Columns["Value3"].OwnerBand = bandValue3;
                bandedGridView1.Columns["Error3"].OwnerBand = bandError3;
                bandedGridView1.Columns["ErrorMax"].OwnerBand = bandErrer;
                #endregion

                #region 设置除最后一列之外的列都不可编辑
                bandedGridView1.Columns["Name"].OptionsColumn.AllowEdit = false;
                bandedGridView1.Columns["Value1"].OptionsColumn.AllowEdit = false;
                bandedGridView1.Columns["Error1"].OptionsColumn.AllowEdit = false;
                bandedGridView1.Columns["Value2"].OptionsColumn.AllowEdit = false;
                bandedGridView1.Columns["Error2"].OptionsColumn.AllowEdit = false;
                bandedGridView1.Columns["Value3"].OptionsColumn.AllowEdit = false;
                bandedGridView1.Columns["Error3"].OptionsColumn.AllowEdit = false;
                #endregion
                bandedGridView1.EndDataUpdate();//结束数据的编辑
                bandedGridView1.EndUpdate();   //结束视图的编辑


                #region 绑定数据格式
                RepositoryItemSpinEdit riSpin = new RepositoryItemSpinEdit();
                RepositoryItemTextEdit riText = new RepositoryItemTextEdit();
                gridControl1.RepositoryItems.Add(riSpin);
                gridControl1.RepositoryItems.Add(riText);
                bandedGridView1.Columns["Id"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Name"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Value1"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Error1"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Value2"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Error2"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Value3"].ColumnEdit = riSpin;
                bandedGridView1.Columns["Error3"].ColumnEdit = riSpin;
                bandedGridView1.Columns["ErrorMax"].ColumnEdit = riText;

                #endregion

            }
        }

        public double SubTotal(List<Idata> datas)
        {
            double total = 0;
            foreach (Idata data in datas)
            {
                double sum = Convert.ToDouble(data.Value1) + Convert.ToDouble(data.Error1);
                total = sum + total;
            }
            return total;
        }

        //// 计算小计
        //private float calcSubTotal(float value, float error)
        //{
        //    return value * error;
        //}
        //private void bandedGridView1_CustomUnboundColumnData(object sender, CustomColumnDataEventArgs e)
        //{

        //    ColumnView colView = sender as ColumnView;
        //    if (e.Column.FieldName == "Error1" && e.IsGetData) e.Value = calcSubTotal(
        //             Convert.ToSingle(colView.GetRowCellValue(e.RowHandle, colView.Columns["Value1"])),
        //             Convert.ToSingle(0.5));
        //}


        /// <summary>
        /// 从附件4中获取参数值
        /// </summary>
        /// <returns></returns>
        private List<string> GetValueFrom4()
        {
            List<string> values = new List<string>();
            //string strNewPath = "E:\\堆芯\\附件4-CoreEasy输出参数表.docx";
            CReadWordFile wordFile = new CReadWordFile(strPath4);
            wordFile.Open();
            string valueT = wordFile.ReadWord(tableNumber41, 4, 2);
            string valueP = wordFile.ReadWord(tableNumber41, 5, 2);
            string valueR = wordFile.ReadWord(tableNumber41, 6, 2);
            string valueRE = wordFile.ReadWord(tableNumber42, 2, 2);
            wordFile.Close();
            #region 检查附件4中的值是否为数字类型
            try
            {
                Convert.ToDouble(valueT);
            }
            catch
            {
                MessageBox.Show("第4个表中温度效应的名义值无效，请检查！");
                return null;
            }
            try
            {
                Convert.ToDouble(valueP);
            }
            catch
            {
                MessageBox.Show("第4个表中功率效应的名义值无效，请检查！");
                return null;
            }
            try
            {
                Convert.ToDouble(valueR);
            }
            catch
            {
                MessageBox.Show("第4个表中燃耗反应性的名义值无效，请检查！");
                return null;
            }
            #endregion
            #region 检查各参数值不为空
            //if (valueT == "")
            //{
            //    MessageBox.Show("第4个表中温度效应的名义值无效，请检查！");
            //    return null;
            //}

            //if (valueP == "")
            //{
            //    MessageBox.Show("第4个表中功率效应的名义值无效，请检查！");
            //    return null;
            //}
            //if (valueR == "")
            //{
            //    MessageBox.Show("第4个表中燃耗反应性的名义值无效，请检查！");
            //    return null;
            //}
            #endregion
            values.Add(valueT);
            values.Add(valueP);
            values.Add(valueR);
            values.Add(valueRE);
            return values;
        }
        /// <summary>
        /// 从附件3中获取参数值
        /// </summary>
        /// <returns></returns>
        private List<string> GetValueFrom3()
        {
            List<string> values = new List<string>();
            //string strNewPath = "E:\\堆芯\\附件3-反应性控制需求表.docx";
            CReadWordFile wordFile = new CReadWordFile(strPath3);
            wordFile.Open();
            string valueR = wordFile.ReadWord(tableNumber3, 7, 2); //临界和裂变材料装量不确定性相关的反应性
            string valueD = wordFile.ReadWord(2, 8, 2);  //停堆深度要求
            wordFile.Close();

            #region
            //if (valueR == "")
            //{
            //    MessageBox.Show("第2个表中临界和裂变材料装量不确定性相关的反应性的量化说明无效，请检查！");
            //    return null;
            //}

            //if (valueD == "")
            //{
            //    MessageBox.Show("第2个表中停堆深度要求的量化说明无效，请检查！");
            //    return null;
            //}
            #endregion
            valueD = "1000";
            valueR = "800";
            #region 检查附件3中的值是否为数字类型
            try
            {
                Convert.ToDouble(valueD);
            }
            catch
            {
                MessageBox.Show("第2个表中停堆深度要求的量化说明无效，请检查！");
                return null;
            }
            try
            {
                Convert.ToDouble(valueR);
            }
            catch
            {
                MessageBox.Show("第2个表中临界和裂变材料装量不确定性相关的反应性的量化说明无效，请检查！");
                return null;
            }
            #endregion
            values.Add(valueR);
            values.Add(valueD);
            values.Add("500");    //卡棒准则下停堆深度要求
            return values;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            List<string> values = new List<string>();
            //string strNewPath = "E:\\堆芯\\反应性控制分析计算.docx";
            CReadWordFile wordFile = new CReadWordFile(strPath5);
            wordFile.Open();

            List<string> names = new List<string>();
            names.Add("Value1");
            names.Add("Error1");
            names.Add("Value2");
            names.Add("Error2");
            names.Add("Value3");
            names.Add("Error3");
            names.Add("ErrorMax");

            for (int i = 0; i < 12; i++)
            {
                if (i != 7 && i != 9 && i != 10 && i != 11)
                {
                    for (int j = 0; j < 7; j++)
                    {
                        string valueCell = this.bandedGridView1.GetRowCellDisplayText(i, bandedGridView1.Columns[names[j]]);
                        wordFile.WriteToWord(tableNumber5, i + 3, 2 + j, valueCell);
                    }
                }
                else
                {
                    for (int j = 0; j < 7; j = j + 2)
                    {
                        string valueCell = this.bandedGridView1.GetRowCellDisplayText(i, bandedGridView1.Columns[names[j]]);
                        int jj = j + 2;
                        if (j != 0)
                            jj = 2 + j / 2;

                        wordFile.WriteToWord(tableNumber5, i + 3, jj, valueCell);
                    }
                    string valueCell7 = this.bandedGridView1.GetRowCellDisplayText(i, bandedGridView1.Columns["ErrorMax"]);
                    wordFile.WriteToWord(tableNumber5, i + 3, 5, valueCell7);
                }
            }

            wordFile.Save();
            wordFile.Close();
        }
        /// <summary>
        /// 获取当前工作目录
        /// </summary>
        /// <returns></returns>
        private string GetWorkPath()
        {
            string currentWorkPath = System.IO.Directory.GetCurrentDirectory();
            return currentWorkPath;
        }


        #region  合并单元格
        public MyMergedCell AddMergedCell(int rowHandle, GridColumn col1, GridColumn col2)
        {
            MyMergedCell cell = new MyMergedCell(rowHandle, col1, col2);
            _MergedCells.Add(cell);
            return cell;
        }
        public void AddMergedCell(int rowHandle, int col1, int col2, object value)
        {
            AddMergedCell(rowHandle, _view.Columns[col1], _view.Columns[col2], value);
        }
        public void AddMergedCell(int rowHandle, GridColumn col1, GridColumn col2, object value)
        {
            MyMergedCell cell = AddMergedCell(rowHandle, col1, col2);
            SafeSetMergedCellValue(cell, value);
        }
        public void SafeSetMergedCellValue(MyMergedCell cell, object value)
        {
            if (cell != null)
            {
                SafeSetMergedCellValue(cell.RowHandle, cell.Column1, value);
                SafeSetMergedCellValue(cell.RowHandle, cell.Column2, value);
            }
        }
        public void SafeSetMergedCellValue(int rowHandle, GridColumn column, object value)
        {
            if (_view.GetRowCellValue(rowHandle, column) != value)
            {
                _view.SetRowCellValue(rowHandle, column, value);
            }
        }


        private MyMergedCell GetMergedCell(int rowHandle, GridColumn column)
        {
            foreach (MyMergedCell cell in _MergedCells)
            {
                if (cell.RowHandle == rowHandle && (column == cell.Column1 || column == cell.Column2))
                    return cell;
            }
            return null;
        }
        private bool IsMergedCell(int rowHandle, GridColumn column)
        {
            return (GetMergedCell(rowHandle, column) != null);
        }


        private void DrawMergedCells(PaintEventArgs e)
        {
            foreach (MyMergedCell cell in _MergedCells)
            {
                //break;
                painter.DrawMergedCell(cell, e);
            }
        }

        private void bandedGridView1_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            SafeSetMergedCellValue(GetMergedCell(e.RowHandle, e.Column), e.Value);
        }

        private void gridControl1_Paint(object sender, PaintEventArgs e)
        {

            //DrawMergedCells(e);
        }

        private void bandedGridView1_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            if (IsMergedCell(e.RowHandle, e.Column))
                e.Handled = !(painter.IsCustomPainting);

        }

        #endregion

        private void bandedGridView1_CellValueChanging(object sender, CellValueChangedEventArgs e)
        {
            //Jisuan();
        }
        public void Jisuan()
        {
            //double total5 = 0;
            //for (int i = 0; i < 6; i++)
            //{
            //    total5 = total5 + Convert.ToDouble(this.bandedGridView1.GetRowCellValue(i, bandedGridView1.Columns["Value1"])) + Convert.ToDouble(this.bandedGridView1.GetRowCellValue(i, bandedGridView1.Columns["Error1"]));

            //}
            //double total7 = total5 + Convert.ToDouble(this.bandedGridView1.GetRowCellValue(6, bandedGridView1.Columns["Value1"])) + Convert.ToDouble(this.bandedGridView1.GetRowCellValue(6, bandedGridView1.Columns["Error1"]));
            //double total9 = total5 + Convert.ToDouble(this.bandedGridView1.GetRowCellValue(8, bandedGridView1.Columns["Value1"]));
            //string errorMax7 = Convert.ToString(this.bandedGridView1.GetRowCellValue(7, bandedGridView1.Columns["ErrorMax"]));
            //string errorMax9 = Convert.ToString(this.bandedGridView1.GetRowCellValue(9, bandedGridView1.Columns["ErrorMax"]));
            //double total10 = total7 / (1 - Convert.ToDouble(errorMax7.Remove(errorMax7.Length - 1)) * 0.01);
            //double total11 = total9 / (1 - Convert.ToDouble(errorMax9.Remove(errorMax7.Length - 1)) * 0.01);

        }


    }
}
