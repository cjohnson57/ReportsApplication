Imports ReportsApplication.Form1.GlobalVariables
Imports System.Data.SqlClient

Public Class Form7
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Loads a table based on the parameter data set's query, in order to populate the combo box
        Label1.Text = "Enter the value for parameter " + paramvarsql(countparamssql).Replace("@", "")
        Dim cn = New SqlConnection(sqlparamconnectionstrings(countparamssql))
        Dim da As SqlDataAdapter = New SqlDataAdapter()
        Dim cmd = New SqlCommand(sqlparamcommands(countparamssql), cn)
        Dim tbl = New DataTable

        If (sqlparams) Then 'If there are parameters, adds them to the query
            For l As Integer = 0 To (countparamssql - 1)
                If (sqlparamconnectionstrings(l) = "NODATASET") Then
                    Dim p As New SqlParameter(paramvarsql(l), paramsql(l))
                    cmd.Parameters.Add(p)
                Else
                    Dim p As New SqlParameter(paramvarsql(l), sqlqueryvalues(l))
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
                errormessages.Add("Error in connection for parameter " + paramvarsql(countparamssql).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for parameter " + paramvarsql(countparamssql).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
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
            paramsql.Add(ComboBox1.Items(0).ToString)
            FindValue()
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        paramsql.Add(ComboBox1.Text) 'Gets the parameter from the combo box
        FindValue()
        Close()
    End Sub

    Private Sub FindValue()
        Dim cn = New SqlConnection(sqlparamconnectionstrings(countparamssql))
        Dim da As SqlDataAdapter = New SqlDataAdapter()
        Dim cmd = New SqlCommand(sqlparamcommands(countparamssql), cn)
        Dim tbl = New DataTable

        If (sqlparams) Then 'If there are parameters, adds them to the report
            For l As Integer = 0 To (countparamssql - 1) 'This loop queries the data base for the parameter table, then sets it based on the user's entered value
                If (sqlparamconnectionstrings(l) = "NODATASET") Then
                    Dim p As New SqlParameter(paramvarsql(l), paramsql(l))
                    cmd.Parameters.Add(p)
                Else
                    Dim p As New SqlParameter(paramvarsql(l), sqlqueryvalues(l))
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
                errormessages.Add("Error in connection for parameter " + paramvarsql(countparamssql).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for parameter " + paramvarsql(countparamssql).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
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
            Dim row As DataRow = tbl.Select("ParameterCaption = '" + paramsql(countparamssql) + "'").FirstOrDefault()
            Dim value As String
            value = row.Item("ParameterValue")
            value = ReplaceEscapeCharacters(value)
            sqlqueryvalues.Add(value)
        Catch ex As Exception
            sqlqueryvalues.Add(paramsql(countparamssql))
        End Try
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