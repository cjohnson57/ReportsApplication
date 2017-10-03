Imports ReportsApplication.Form1.GlobalVariables

Public Class Form6
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Enter the value for parameter " + paramvarsql(countparamssql)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        paramsql.Add(DateTimePicker1.Value) 'Gets the parameter from the datetimepicker
        Close()
    End Sub
End Class