Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Public Class Menu

    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub bnHelloWorld_Click(sender As Object, e As RibbonControlEventArgs) Handles bnHelloWorld.Click
        Dim ActiveWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(1)
        Dim worksheet As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(ActiveWorkSheet)
        Dim cells As Range = worksheet.Range("A1")
        cells.Value = "Hello World!"
        cells.WrapText = True
    End Sub
End Class
