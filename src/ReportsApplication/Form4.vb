Imports ReportsApplication.Form1.GlobalVariables

Public Class Form4
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Enter the value for parameter " + tempadomdparameter.ParamVar.Replace("@", "")
        Dim firstset As Boolean = True
        For i As Integer = 0 To (adomdvalues.Count - 1)
            ComboBox1.Items.Add(adomdvalues(i))
            If firstset Then
                ComboBox1.Text = adomdvalues(i)
                firstset = False
            End If
        Next
        If ComboBox1.Items.Count = 1 Then
            tempadomdparameter.Parameter = ComboBox1.Items(0).ToString
            tempadomdparameter.QueryValues = ComboBox1.Items(0).ToString
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempadomdparameter.Parameter = ComboBox1.Items(0).ToString
        tempadomdparameter.QueryValues = ComboBox1.Items(0).ToString
        AdomdParameters.Add(tempadomdparameter)
        Close()
    End Sub
End Class