Imports ReportsApplication.Form1.GlobalVariables
Imports System.Data.SqlClient

Public Class Form7
    Private tbl As DataTable
    Private k As Integer
    Private skippedone As Boolean
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Loads a table based on the parameter data set's query, in order to populate the combo box
        Label1.Text = "Enter the value for parameter " + tempsqlparameter.ParamVar.Replace("@", "")
        Dim cn = New SqlConnection(tempsqlparameter.ConnectionString)
        Dim da As SqlDataAdapter = New SqlDataAdapter()
        Dim cmd = New SqlCommand(tempsqlparameter.Query, cn)
        tbl = New DataTable

        If (SQLParameters.Count > 0) Then 'If there are parameters, adds them to the command using the parameters found in the SetParameters function
            For l As Integer = 0 To (SQLParameters.Count - 1)
                cmd.Parameters.AddWithValue(SQLParameters(l).ParamVar, SQLParameters(l).Parameter)
            Next
        End If

        If (AdomdParameters.Count() > 0) Then 'It's possible for a SQL connection to use AS parameters, so they too must be added
            For l As Integer = 0 To (AdomdParameters.Count() - 1)
                If (AdomdParameters(l).ConnectionString = "NODATASET") Then
                    cmd.Parameters.AddWithValue(AdomdParameters(l).ParamVar, AdomdParameters(l).Parameter)
                Else
                    cmd.Parameters.AddWithValue(AdomdParameters(l).ParamVar, AdomdParameters(l).QueryValues)
                End If
            Next
        End If

        da.SelectCommand = cmd
        Try
            cn.Open()
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception
            If (renderingmultiple) Then
                wereerrors = True
                errormessages.Add("Error in connection for parameter " + tempsqlparameter.ParamVar.Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for parameter " + tempsqlparameter.ParamVar.Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            End If
        End Try

        Dim firstset As Boolean = True
        Dim j As Integer = 0
        k = 0
        For i As Integer = 0 To (tbl.Columns.Count - 1)
            If tempsqlparameter.LabelField.Contains(Form1.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                j = i
            End If
            If tempsqlparameter.ValueField.Contains(Form1.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                k = i
            End If
        Next
        skippedone = False
        For i As Integer = 0 To (tbl.Rows.Count - 1)
            If (tbl.Rows(i).ItemArray(j).ToString <> "All") Then
                ComboBox1.Items.Add(tbl.Rows(i).ItemArray(j).ToString())
                If firstset Then
                    ComboBox1.Text = tbl.Rows(i).ItemArray(j).ToString()
                    firstset = False
                End If
            Else
                skippedone = True
            End If
        Next
        If ComboBox1.Items.Count = 1 Then
            tempsqlparameter.Parameter = ComboBox1.Items(0).ToString()
            tempsqlparameter.QueryValues = tbl.Rows(0).ItemArray(k).ToString()
            SQLParameters.Add(tempsqlparameter)
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempsqlparameter.Parameter = ComboBox1.Text
        If (skippedone) Then
            tempsqlparameter.QueryValues = tbl.Rows(ComboBox1.SelectedIndex + 1).ItemArray(k).ToString
        Else
            tempsqlparameter.QueryValues = tbl.Rows(ComboBox1.SelectedIndex).ItemArray(k).ToString
        End If
        SQLParameters.Add(tempsqlparameter)
        Close()
    End Sub
End Class