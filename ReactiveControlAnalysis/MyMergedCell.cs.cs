using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.XtraGrid.Columns;

namespace ReactiveControlAnalysis
{
   public  class MyMergedCell
    {
       public MyMergedCell(int rowHandle, GridColumn col1, GridColumn col2)
        {
            this._rowHandle = rowHandle;
            this._column1 = col1;
            this._column2 = col2;

        }
       public  int _rowHandle;

        public int RowHandle
        {
            get { return _rowHandle; }
            set { _rowHandle = value; }
        }
      public   GridColumn _column1, _column2;

        public GridColumn Column1
        {
            get { return _column1; }
            set { _column1 = value; }
        }

        public GridColumn Column2
        {
            get { return _column2; }
            set { _column2 = value; }
        }
    }
}
