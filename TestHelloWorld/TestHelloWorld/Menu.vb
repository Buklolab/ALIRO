Imports Microsoft.Office.Tools.Ribbon

Public Class Menu

    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTouchme.Click
        Dim ActiveWorksheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(1)
        Dim Worksheets As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(ActiveWorksheet)

        Dim celda As Excel.Range = Worksheets.Range("A1")

        celda.Value = "Hello  World"

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPresioname.Click
        MsgBox("Hola mundo")
    End Sub
End Class
