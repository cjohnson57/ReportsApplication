Imports ReportsApplication.Form1.GlobalVariables
Imports Microsoft.AnalysisServices.AdomdClient
Imports ReportsApplication.Form1

Public Class Form3
    Private tbl As DataTable
    Private k As Integer
    Private skippedone As Boolean
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Loads a table based on the parameter data set's query, in order to populate the combo box
        Label1.Text = "Enter the value for parameter " + tempadomdparameter.ParamVar.Replace("@", "")
        Dim cn = New AdomdConnection(tempadomdparameter.ConnectionString)
        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(tempadomdparameter.Query, cn)
        tbl = New DataTable

        If (SQLParameters.Count > 0) Then 'It's possible for an AS connection to use SQL parameters, so they too must be added
            For l As Integer = 0 To (SQLParameters.Count - 1)
                Dim p As New AdomdParameter(SQLParameters(l).ParamVar, SQLParameters(l).Parameter)
                cmd.Parameters.Add(p)
            Next
        End If

        If (AdomdParameters.Count() > 0) Then 'If there are parameters, adds them to the command using the parameters found in the SetParameters function
            For l As Integer = 0 To (AdomdParameters.Count() - 1)
                If (AdomdParameters(l).ConnectionString = "NODATASET") Then
                    Dim p As New AdomdParameter(AdomdParameters(l).ParamVar, AdomdParameters(l).Parameter)
                    cmd.Parameters.Add(p)
                Else
                    Dim p As New AdomdParameter(AdomdParameters(l).ParamVar, AdomdParameters(l).QueryValues)
                    cmd.Parameters.Add(p)
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
                errormessages.Add("Error in connection for parameter " + tempadomdparameter.ParamVar.Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for parameter " + tempadomdparameter.ParamVar.Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            End If
        End Try

        Dim firstset As Boolean = True
        Dim j As Integer = 0
        k = 0
        For i As Integer = 0 To (tbl.Columns.Count - 1)
            If tempadomdparameter.LabelField.Contains(Form1.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                j = i
            End If
            If tempadomdparameter.ValueField.Contains(Form1.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
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
            tempadomdparameter.Parameter = ComboBox1.Items(0).ToString()
            tempadomdparameter.QueryValues = tbl.Rows(0).ItemArray(k).ToString()
            AdomdParameters.Add(tempadomdparameter)
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempadomdparameter.Parameter = ComboBox1.Text
        If (skippedone) Then
            tempadomdparameter.QueryValues = tbl.Rows(ComboBox1.SelectedIndex + 1).ItemArray(k).ToString
        Else
            tempadomdparameter.QueryValues = tbl.Rows(ComboBox1.SelectedIndex).ItemArray(k).ToString
        End If
        AdomdParameters.Add(tempadomdparameter)
        Close()
    End Sub
End Class