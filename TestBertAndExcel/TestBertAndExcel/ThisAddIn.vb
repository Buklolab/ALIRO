Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Function Sumar()

        Dim resultado As Double
        resultado = Application.Run("BERT.Call", "sum", 5, 7)

        Return resultado
    End Function
End Class
