using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraEditors.Repository;
using DevExpress.Utils;
using System.Diagnostics;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        string XLSfileName = "C:\\Test.xls";
        string[] names = new string[] { "Nasdaq", "S&P 500", "10 Yr Bond", "Oil", "Gold" };
        decimal[] summaries = new decimal[] { 10038.38m, 2147.87m, 1068.13m, 3.6940m, 1079.60m };

        Random r = new Random();

        decimal GetRandomValue(int i)
        {
            return summaries[i] * r.Next(100) / 100;
        }

         private DataTable CreateTable()
        {
            DataTable tbl = new DataTable();
            tbl.Columns.Add("Index", typeof(string));
            tbl.Columns.Add("Week1", typeof(decimal));
            tbl.Columns.Add("Week2", typeof(decimal));
            tbl.Columns.Add("Week3", typeof(decimal));
            tbl.Columns.Add("Week4", typeof(decimal));
            for (int i = 0; i < names.Length; i++)
            {
                tbl.Rows.Add(new object[] { names[i], GetRandomValue(i), GetRandomValue(i), GetRandomValue(i), GetRandomValue(i) });
            }
            return tbl;
        }
        

        public Form1()
        {
            InitializeComponent();
            gridControl1.DataSource = CreateTable();
        }

        void CreateUnboundColumns()
        {
            RepositoryItemPictureEdit pictureEdit = new RepositoryItemPictureEdit();
            int count = gridView1.Columns.Count;
            for (int i = 0; i < count - 2; i++)
            {
                GridColumn column = gridView1.Columns.Add();
                column.FieldName = "unbound" + i.ToString();
                column.UnboundType = DevExpress.Data.UnboundColumnType.Object;
                column.ColumnEdit = pictureEdit;
                column.VisibleIndex = (2 * i) + 2;
            }
        }

        void ShowAnalys()
        {
            CreateUnboundColumns();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShowAnalys();
            button1.Enabled = false;
            gridControl1.ExportToXls(XLSfileName);
            Process.Start(XLSfileName);
        }

        Image GetImageByValue(int value)
        {
            int offset = 15;
            Bitmap b = new Bitmap(60, 32);
            Graphics g = Graphics.FromImage(b);
            Image image;
            Brush brush;
            Font font = new Font(AppearanceObject.DefaultFont.FontFamily, 10, FontStyle.Bold);
            if (value > 0)
            {
                brush = Brushes.Green;
                image = new Bitmap(global::WindowsApplication1.Properties.Resources.sort_ascending_32);
            }
            else
            {
                brush = Brushes.Red;
                image = new Bitmap(global::WindowsApplication1.Properties.Resources.sort_descending_32);
            }
            g.DrawImageUnscaled(image, new Point(0, 0));
            g.FillRectangle(Brushes.LightYellow, new Rectangle(offset, 0, b.Width, b.Height));
            g.DrawString(Math.Abs(value) + "%", font, brush, new PointF(offset, offset/2));
            return b;
        }

        int GetValue(int rowHanlde, int columnIndex)
        {
            int value1 = Convert.ToInt32(gridView1.GetRowCellValue(rowHanlde, gridView1.VisibleColumns[columnIndex - 1]));
            int value2 = Convert.ToInt32(gridView1.GetRowCellValue(rowHanlde, gridView1.VisibleColumns[columnIndex + 1]));
            if (value1 == 0) return 0;
            return (value2 - value1) * 100/ value1;
        }

        private void gridView1_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
            if (e.IsGetData)
                e.Value = GetImageByValue(GetValue(e.RowHandle, e.Column.VisibleIndex));
        }
    }
}