Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup


    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Function suma()
        Dim result As Double
        result = Application.Run("BERT.Call", "sum", 1, 2, 3, 4)
        Return result
    End Function

    Function grafica()
        Dim result As Double
        result = Application.Run("BERT.Exec", "plot( sort( rnorm( 1000 )))")
        Return result
    End Function
End Class
