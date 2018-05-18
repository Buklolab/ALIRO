Imports Microsoft.Office.Tools.Ribbon

Public Class Menu

    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnSumar_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSumar.Click
        Dim ActiveWorksheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(1)
        Dim Worksheet As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(ActiveWorksheet)

        Dim celda As Excel.Range = Worksheet.Range("A1")
        celda.Value = Globals.ThisAddIn.Sumar()
    End Sub
End Class
