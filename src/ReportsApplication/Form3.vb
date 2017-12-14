Imports ReportsApplication.Form1.GlobalVariables
Imports Microsoft.AnalysisServices.AdomdClient

Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Loads a table based on the parameter data set's query, in order to populate the combo box
        Label1.Text = "Enter the value for parameter " + tempadomdparameter.ParamVar.Replace("@", "")
        Dim cn = New AdomdConnection(tempadomdparameter.ConnectionString)
        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(tempadomdparameter.Query, cn)
        Dim tbl = New DataTable

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
            tempadomdparameter.Parameter = ComboBox1.Items(0).ToString
            FindValue()
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tempadomdparameter.Parameter = ComboBox1.Text
        FindValue()
        Close()
    End Sub

    Private Sub FindValue()
        Dim cn = New AdomdConnection(tempadomdparameter.ConnectionString)
        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(tempadomdparameter.Query, cn)
        Dim tbl = New DataTable

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
        Try
            For i As Integer = 0 To tbl.Columns.Count - 1
                If tbl.Columns(i).ColumnName.Contains("ParameterCaption") Then
                    tbl.Columns(i).ColumnName = "ParameterCaption"
                ElseIf tbl.Columns(i).ColumnName.Contains("ParameterValue") Then
                    tbl.Columns(i).ColumnName = "ParameterValue"
                End If
            Next
            Dim row As DataRow = tbl.Select("ParameterCaption = '" + tempadomdparameter.Parameter + "'").FirstOrDefault()
            Dim value As String
            value = row.Item("ParameterValue")
            value = ReplaceEscapeCharacters(value)
            tempadomdparameter.QueryValues = value
        Catch ex As Exception
            tempadomdparameter.QueryValues = tempadomdparameter.Parameter
        End Try
        AdomdParameters.Add(tempadomdparameter)
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