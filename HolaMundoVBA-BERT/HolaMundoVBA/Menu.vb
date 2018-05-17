Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Public Class Menu

    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub bnHelloWorld_Click(sender As Object, e As RibbonControlEventArgs) Handles bnHelloWorld.Click
        Dim cells As Range = getSingleCell()
        cells.Value = "Hello World!"
        cells.WrapText = True
    End Sub

    Private Sub bnSum_Click(sender As Object, e As RibbonControlEventArgs) Handles bnSum.Click
        Dim cells As Range = getSingleCell()
        cells.Value = Globals.ThisAddIn.suma()
    End Sub

    Function getSingleCell()
        Dim ActiveWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(1)
        Dim worksheet As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(ActiveWorkSheet)
        Dim cells As Range = worksheet.Range("A1")
        Return cells
    End Function

    Private Sub bnGraph_Click(sender As Object, e As RibbonControlEventArgs) Handles bnGraph.Click
        Dim cells As Range = getSingleCell()
        cells.Value = Globals.ThisAddIn.grafica()
    End Sub
End Class
