Imports ReportsApplication.Form1.GlobalVariables

Public Class Form6
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Enter the value for parameter " + tempsqlparameter.ParamVar
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempsqlparameter.Parameter = DateTimePicker1.Value
        SQLParameters.Add(tempsqlparameter)
        Close()
    End Sub
End Class