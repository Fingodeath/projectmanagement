
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Namespace nExcel
    Public Class DatagridToExcel

        Public Shared Sub ExportDataGridViewToExcel(ByVal myDataGridView As DataGridView)
            Dim oExcel As Excel.Application = Nothing
            'Excel Application 
            Dim oBook As Excel.Workbook = Nothing
            ' Excel Workbook 
            Dim oSheetsColl As Excel.Sheets = Nothing
            ' Excel Worksheets collection 
            Dim oSheet As Excel.Worksheet = Nothing
            ' Excel Worksheet 
            Dim oRange As Excel.Range = Nothing
            ' Cell or Range in worksheet 
            Dim oMissing As Object = System.Reflection.Missing.Value
            Try
                ' Create an instance of Excel. 
                oExcel = New Excel.Application
                ' Make Excel visible to the user. 
                oExcel.Visible = True

                ' Set the UserControl property so Excel won't shut down. 
                oExcel.UserControl = True
                ' System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US"); 
                ' Add a workbook. 
                oBook = oExcel.Workbooks.Add(oMissing)
                ' Get worksheets collection 
                oSheetsColl = oExcel.Worksheets
                ' Get Worksheet "Sheet1" 
                oSheet = DirectCast(oSheetsColl.Item("Sheet1"), Excel.Worksheet)
                ' Export titles 
                For j As Integer = 0 To myDataGridView.Columns.Count - 1
                    oRange = DirectCast(oSheet.Cells(1, j + 1), Excel.Range)
                    oRange.Value2 = myDataGridView.Columns(j).HeaderText
                Next

                ' Export data 
                For i As Integer = 0 To myDataGridView.Rows.Count - 1
                    For j As Integer = 0 To myDataGridView.Columns.Count - 1
                        oRange = DirectCast(oSheet.Cells(i + 2, j + 1), Excel.Range)
                        oRange.Value2 = myDataGridView(j, i).Value
                    Next
                Next

                ' Release the variables. 
                oBook = Nothing

                'oExcel.Quit()
                oExcel = Nothing

                ' Collect garbage. 
                GC.Collect()
            Catch ex As Exception
                MsgBox(ex.Message + " : Medical Sandbox " + " : DataGridtoExcel - " + " ExportDataGridViewToExcel")
            Finally
                ' Release the variables. 
                oBook = Nothing
                'oExcel.Quit()
                oExcel = Nothing
                ' Collect garbage. 
                GC.Collect()
            End Try
        End Sub
    End Class
End Namespace
