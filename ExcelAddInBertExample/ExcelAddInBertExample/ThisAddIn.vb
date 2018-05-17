Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Function sum()
        Dim result As Double
        result = Application.Run("BERT.Call", "sum", 1, 2, 3)
        Return result
    End Function

End Class
