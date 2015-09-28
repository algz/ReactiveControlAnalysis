using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.Utils.Drawing;
using System.Drawing;
using System.Windows.Forms;

namespace ReactiveControlAnalysis
{
    class MyGridPainter : GridPainter
    {
        GridView view;
        public MyGridPainter(GridView view)
            : base(view)
        {
            this.view = view;
        }
        public bool IsCustomPainting;

        public bool DrawMergedCell(MyMergedCell cell, PaintEventArgs e)
        {
            //return true;
            int delta = cell.Column1.VisibleIndex - cell.Column2.VisibleIndex;
            if (Math.Abs(delta) > 1)
                return false;
            GridViewInfo vi = View.GetViewInfo() as GridViewInfo;
 
            GridCellInfo gridCellInfo1 = vi.GetGridCellInfo(cell.RowHandle, cell.Column1.VisibleIndex);
            GridCellInfo gridCellInfo2 = vi.GetGridCellInfo(cell.RowHandle, cell.Column2.VisibleIndex);
            if (gridCellInfo1 == null || gridCellInfo2 == null)
                return false;
            Rectangle targetRect = Rectangle.Union(gridCellInfo1.Bounds, gridCellInfo2.Bounds);
            gridCellInfo1.Bounds = targetRect;
            gridCellInfo1.CellValueRect = targetRect;
            gridCellInfo2.Bounds = targetRect;
            gridCellInfo2.CellValueRect = targetRect;
            if (delta < 0)
                gridCellInfo1 = gridCellInfo2;

            Rectangle bounds = gridCellInfo1.ViewInfo.Bounds;
            bounds.Width = targetRect.Width;
            bounds.Height = targetRect.Height;
            gridCellInfo1.ViewInfo.Bounds = bounds;
            gridCellInfo1.ViewInfo.CalcViewInfo(e.Graphics);
            IsCustomPainting = true;
            GraphicsCache cache = new GraphicsCache(e.Graphics);

            gridCellInfo1.Appearance.FillRectangle(cache, gridCellInfo1.Bounds);
            DrawRowCell(new GridViewDrawArgs(cache, vi, vi.ViewRects.Bounds), gridCellInfo1);
            IsCustomPainting = false;

            return true;
        }

    }
}
