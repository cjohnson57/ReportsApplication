Imports ReportsApplication.Form1.GlobalVariables

Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Enter the value for parameter " + tempsqlparameter.ParamVar
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempsqlparameter.Parameter = TextBox1.Text
        SQLParameters.Add(tempsqlparameter)
        Close()
    End Sub
End Class