Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub bHelloWorld_Click(sender As Object, e As RibbonControlEventArgs) Handles bHelloWorld.Click
        Dim ActiveWorksheet As Microsoft.Office.Interop.Excel.Worksheet =
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(1)

        Dim Worksheet As Microsoft.Office.Tools.Excel.Worksheet =
            Globals.Factory.GetVstoObject(ActiveWorksheet)

        Dim cellB2 As Excel.Range = Worksheet.Range("B2")

        cellB2.Value = "Hello World"
    End Sub

    Private Sub bSum_Click(sender As Object, e As RibbonControlEventArgs) Handles bSum.Click
        Dim ActiveWorksheet As Microsoft.Office.Interop.Excel.Worksheet =
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(1)

        Dim Worksheet As Microsoft.Office.Tools.Excel.Worksheet =
            Globals.Factory.GetVstoObject(ActiveWorksheet)

        Dim cellB2 As Excel.Range = Worksheet.Range("B3")

        cellB2.Value = Globals.ThisAddIn.sum()
    End Sub
End Class
