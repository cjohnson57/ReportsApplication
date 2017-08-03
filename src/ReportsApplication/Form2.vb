Imports ReportsApplication.Form1.GlobalVariables

Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Enter the value for parameter " + paramvarsql(countparamssql)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        paramsql.Add(TextBox1.Text) 'Gets the parameter from the text box
        Close()
    End Sub
End Class