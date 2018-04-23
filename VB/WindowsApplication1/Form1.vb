Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.Utils
Imports System.Diagnostics

Namespace WindowsApplication1
	Partial Public Class Form1
		Inherits Form
		Private XLSfileName As String = "C:\Test.xls"
		Private names() As String = { "Nasdaq", "S&P 500", "10 Yr Bond", "Oil", "Gold" }
		Private summaries() As Decimal = { 10038.38D, 2147.87D, 1068.13D, 3.6940D, 1079.60D }

		Private r As New Random()

		Private Function GetRandomValue(ByVal i As Integer) As Decimal
			Return summaries(i) * r.Next(100) / 100
		End Function

		 Private Function CreateTable() As DataTable
			Dim tbl As New DataTable()
			tbl.Columns.Add("Index", GetType(String))
			tbl.Columns.Add("Week1", GetType(Decimal))
			tbl.Columns.Add("Week2", GetType(Decimal))
			tbl.Columns.Add("Week3", GetType(Decimal))
			tbl.Columns.Add("Week4", GetType(Decimal))
			For i As Integer = 0 To names.Length - 1
				tbl.Rows.Add(New Object() { names(i), GetRandomValue(i), GetRandomValue(i), GetRandomValue(i), GetRandomValue(i) })
			Next i
			Return tbl
		 End Function


		Public Sub New()
			InitializeComponent()
			gridControl1.DataSource = CreateTable()
		End Sub

		Private Sub CreateUnboundColumns()
			Dim pictureEdit As New RepositoryItemPictureEdit()
			Dim count As Integer = gridView1.Columns.Count
			For i As Integer = 0 To count - 2 - 1
				Dim column As GridColumn = gridView1.Columns.Add()
				column.FieldName = "unbound" & i.ToString()
				column.UnboundType = DevExpress.Data.UnboundColumnType.Object
				column.ColumnEdit = pictureEdit
				column.VisibleIndex = (2 * i) + 2
			Next i
		End Sub

		Private Sub ShowAnalys()
			CreateUnboundColumns()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			ShowAnalys()
			button1.Enabled = False
			gridControl1.ExportToXls(XLSfileName)
			Process.Start(XLSfileName)
		End Sub

		Private Function GetImageByValue(ByVal value As Integer) As Image
			Dim offset As Integer = 15
			Dim b As New Bitmap(60, 32)
			Dim g As Graphics = Graphics.FromImage(b)
			Dim image As Image
			Dim brush As Brush
			Dim font As New Font(AppearanceObject.DefaultFont.FontFamily, 10, FontStyle.Bold)
			If value > 0 Then
				brush = Brushes.Green
				image = New Bitmap(My.Resources.sort_ascending_32)
			Else
				brush = Brushes.Red
				image = New Bitmap(My.Resources.sort_descending_32)
			End If
			g.DrawImageUnscaled(image, New Point(0, 0))
			g.FillRectangle(Brushes.LightYellow, New Rectangle(offset, 0, b.Width, b.Height))
			g.DrawString(Math.Abs(value) & "%", font, brush, New PointF(offset, offset\2))
			Return b
		End Function

		Private Function GetValue(ByVal rowHanlde As Integer, ByVal columnIndex As Integer) As Integer
			Dim value1 As Integer = Convert.ToInt32(gridView1.GetRowCellValue(rowHanlde, gridView1.VisibleColumns(columnIndex - 1)))
			Dim value2 As Integer = Convert.ToInt32(gridView1.GetRowCellValue(rowHanlde, gridView1.VisibleColumns(columnIndex + 1)))
			If value1 = 0 Then
				Return 0
			End If
			Return (value2 - value1) * 100\ value1
		End Function

		Private Sub gridView1_CustomUnboundColumnData(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs) Handles gridView1.CustomUnboundColumnData
			If e.IsGetData Then
				e.Value = GetImageByValue(GetValue(e.RowHandle, e.Column.VisibleIndex))
			End If
		End Sub
	End Class
End Namespace