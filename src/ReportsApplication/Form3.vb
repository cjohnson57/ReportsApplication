Imports ReportsApplication.Form1.GlobalVariables
Imports Microsoft.AnalysisServices.AdomdClient

Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Loads a table based on the parameter data set's query, in order to populate the combo box
        Label1.Text = "Enter the value for parameter " + paramvaradomd(countparamsadomd).Replace("@", "")
        Dim cn = New AdomdConnection(adomdparamconnectionstrings(countparamsadomd))
        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(paramcommands(countparamsadomd), cn)
        Dim tbl = New DataTable

        If (adomdparams) Then 'If there are parameters, adds them to the query
            For l As Integer = 0 To (countparamsadomd - 1)
                If (adomdparamconnectionstrings(l) = "NODATASET") Then
                    Dim p As New AdomdParameter(paramvaradomd(l), paramadomd(l))
                    cmd.Parameters.Add(p)
                Else
                    Dim p As New AdomdParameter(paramvaradomd(l), adomdqueryvalues(l))
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
                errormessages.Add("Error in connection for parameter " + paramvaradomd(countparamsadomd).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for parameter " + paramvaradomd(countparamsadomd).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
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
                ComboBox1.Items.Add(tbl.Rows(i).ItemArray(j))
                If firstset Then
                    ComboBox1.Text = tbl.Rows(i).ItemArray(j)
                    firstset = False
                End If
            End If
        Next
        If ComboBox1.Items.Count = 1 Then
            paramadomd.Add(ComboBox1.Items(0).ToString)
            FindValue()
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        paramadomd.Add(ComboBox1.Text) 'Gets the parameter from the combo box
        FindValue()
        Close()
    End Sub

    Private Sub FindValue()
        Dim cn = New AdomdConnection(adomdparamconnectionstrings(countparamsadomd))
        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(paramcommands(countparamsadomd), cn)
        Dim tbl = New DataTable

        If (adomdparams) Then 'If there are parameters, adds them to the report
            For l As Integer = 0 To (countparamsadomd - 1) 'This loop queries the data base for the parameter table, then sets it based on the user's entered value
                If (adomdparamconnectionstrings(l) = "NODATASET") Then
                    Dim p As New AdomdParameter(paramvaradomd(l), paramadomd(l))
                    cmd.Parameters.Add(p)
                Else
                    Dim p As New AdomdParameter(paramvaradomd(l), adomdqueryvalues(l))
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
                errormessages.Add("Error in connection for parameter " + paramvaradomd(countparamsadomd).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for parameter " + paramvaradomd(countparamsadomd).Replace("@", "") + "'s dataset in report " + filenametemp + Environment.NewLine + ex.Message)
            End If
        End Try
        Try
            tbl.Columns(1).ColumnName = "ParameterCaption"
            tbl.Columns(2).ColumnName = "ParameterValue"
            Dim row As DataRow = tbl.Select("ParameterCaption = '" + paramadomd(countparamsadomd) + "'").FirstOrDefault()
            Dim value As String
            value = row.Item("ParameterValue")
            value = ReplaceEscapeCharacters(value)
            adomdqueryvalues.Add(value)
        Catch ex As Exception
            adomdqueryvalues.Add(paramadomd(countparamsadomd))
        End Try
    End Sub

    Private Function ReplaceEscapeCharacters(finaloutcome As String)
        Dim chars = New String() {"'", "&", ">", "<"}
        Dim escs = New String() {" &apos;", "&amp;", "&gt;", "&lt;"}
        For i As Integer = 0 To (escs.Length - 1)
            finaloutcome = finaloutcome.Replace(escs(i), chars(i)) 'Replaces escape characters with their normal characters
        Next
        Return finaloutcome
    End Function
End Class