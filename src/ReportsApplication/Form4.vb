Imports ReportsApplication.Form1.GlobalVariables

Public Class Form4
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Enter the value for parameter " + paramvaradomd(countparamsadomd).Replace("@", "")
        Dim firstset As Boolean = True
        For i As Integer = 0 To (adomdvalues.Count - 1)
            ComboBox1.Items.Add(adomdvalues(i))
            If firstset Then
                ComboBox1.Text = adomdvalues(i)
                firstset = False
            End If
        Next
        If ComboBox1.Items.Count = 1 Then
            paramadomd.Add(ComboBox1.Items(0).ToString)
            adomdqueryvalues.Add(ComboBox1.Items(0).ToString)
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        paramadomd.Add(ComboBox1.Text) 'Gets the parameter from the combo box
        adomdqueryvalues.Add(ComboBox1.Text)
        Close()
    End Sub
End Class