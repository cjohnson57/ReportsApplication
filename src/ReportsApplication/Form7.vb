Imports ReportsApplication.Form1.GlobalVariables
Imports System.Data.SqlClient

Public Class Form7
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Loads a table based on the parameter data set's query, in order to populate the combo box
        Label1.Text = "Enter the value for parameter " + tempsqlparameter.ParamVar.Replace("@", "")
        Dim cn = New SqlConnection(tempsqlparameter.ConnectionString)
        Dim da As SqlDataAdapter = New SqlDataAdapter()
        Dim cmd = New SqlCommand(tempsqlparameter.Query, cn)
        Dim tbl = New DataTable

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
        For i As Integer = 0 To (tbl.Columns.Count - 1)
            If tbl.Columns(i).ColumnName.Contains("ParameterCaption") Then
                j = i
                Exit For
            End If
        Next

        For i As Integer = 0 To (tbl.Rows.Count - 1)
            If (tbl.Rows(i).ItemArray(j).ToString <> "All") Then
                ComboBox1.Items.Add(tbl.Rows(i).ItemArray(j).ToString())
                If firstset Then
                    ComboBox1.Text = tbl.Rows(i).ItemArray(j).ToString()
                    firstset = False
                End If
            End If
        Next
        If ComboBox1.Items.Count = 1 Then
            tempsqlparameter.Parameter = ComboBox1.Items(0).ToString
            FindValue()
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempsqlparameter.Parameter = ComboBox1.Text
        FindValue()
        Close()
    End Sub

    Private Sub FindValue()
        Dim cn = New SqlConnection(tempsqlparameter.ConnectionString)
        Dim da As SqlDataAdapter = New SqlDataAdapter()
        Dim cmd = New SqlCommand(tempsqlparameter.Query, cn)
        Dim tbl = New DataTable

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
        Try
            For i As Integer = 0 To tbl.Columns.Count - 1
                If tbl.Columns(i).ColumnName.Contains("ParameterCaption") Then
                    tbl.Columns(i).ColumnName = "ParameterCaption"
                ElseIf tbl.Columns(i).ColumnName.Contains("ParameterValue") Then
                    tbl.Columns(i).ColumnName = "ParameterValue"
                End If
            Next
            Dim row As DataRow = tbl.Select("ParameterCaption = '" + tempsqlparameter.Parameter + "'").FirstOrDefault()
            Dim value As String
            value = row.Item("ParameterValue")
            value = ReplaceEscapeCharacters(value)
            tempsqlparameter.QueryValues = value
        Catch ex As Exception
            tempsqlparameter.QueryValues = tempsqlparameter.Parameter
        End Try
        SQLParameters.Add(tempsqlparameter)
    End Sub

    Private Function ReplaceEscapeCharacters(finaloutcome As String)
        Dim chars = New String() {"'", "&", ">", "<"}
        Dim escs = New String() {"&apos;", "&amp;", "&gt;", "&lt;"}
        For i As Integer = 0 To (escs.Length - 1)
            finaloutcome = finaloutcome.Replace(escs(i), chars(i)) 'Replaces escape characters with their normal characters
        Next
        Return finaloutcome
    End Function
End Class