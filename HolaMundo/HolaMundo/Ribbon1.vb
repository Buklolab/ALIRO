﻿Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnHola_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnHola.Click
        MsgBox("Hola mundo", MsgBoxStyle.Information, "Ejemplo")

    End Sub
End Class
