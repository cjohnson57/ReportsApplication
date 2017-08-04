﻿Imports System.Data.SqlClient
Imports ReportsApplication.Form1.GlobalVariables
Imports Microsoft.Reporting.WinForms
Imports Microsoft.AnalysisServices.AdomdClient
Imports System.IO
Imports System.Text.RegularExpressions
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO


Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
        ReportViewer1.RefreshReport()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Report Files|*.rdlc;*.rdl"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.ShowDialog()
        If (OpenFileDialog1.FileName <> "") Then 'Makes sure a file was chosen, if not the dialog box just closes
            ReportViewer1.LocalReport.DataSources.Clear()
            ReportViewer1.LocalReport.ReportPath = OpenFileDialog1.FileName 'Sets the report equal to the file chosen by the user
            SetData(False, OpenFileDialog1.SafeFileName) 'Sets the source for the data, finds the query, and populates the data table
            ReportViewer1.RefreshReport()
            Text = "Report Viewer - " + OpenFileDialog1.SafeFileName 'Places the name of the report in the control's title bar
        End If
        DeleteFilesFromFolder()
    End Sub

    Private Sub SetData(saveparameters As Boolean, filename As String)
        Dim filereaderdatasources As String = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data for data source purposes
        Dim filereaderdatasets As String = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data for data set purposes
        Dim numdatasources As Integer = NumTimes("DataSource Name", "Report") 'Finds the number of data sources in the report
        Dim numdatasets As Integer = NumTimes("DataSet Name", "Report") 'Finds the number of data sets in the report
        Dim i As Integer = -1 'Used as an index when finding strings in the report
        Dim count As Integer = 0
        Dim datasetnames(numdatasets - 1) As String 'Array of data set names
        Dim connectionstrings(numdatasources - 1) As String 'Array of connection strings for each data source
        Dim datasourcenames(numdatasources - 1) As String 'Array of data source names
        Dim currentdatasourcename As String
        Dim query As String
        Dim datasource As Integer
        Dim isanalysis(numdatasources - 1) As Boolean

        If (numdatasources = 0) Then 'If there is no data source, the report simply renders since no data has to be provided
            SetParameters(datasourcenames, connectionstrings, numdatasources, filename)
            ReplaceFields(filename) 'Replaces fields, required for AS reports 
            If (Not saveparameters) Then 'Clears variables unless rendering multiple reports
                ClearGlobalVariables()
            End If
            Exit Sub

        ElseIf (numdatasources > 1) Then
            While count < numdatasources 'This while loop finds each connection string in the report if there are multiple data sources
                i = filereaderdatasources.IndexOf("<DataSource Name", i + 1)
                connectionstrings(count) = FindString("<ConnectString", "</ConnectString", i)
                If (FindString("<DataProvider", "</DataProvider", i) <> "SQL") Then
                    isanalysis(count) = True
                End If
                If (Not connectionstrings(count).Contains("Integrated Security")) Then
                    connectionstrings(count) += ";Integrated Security=True"
                End If
                datasourcenames(count) = FindDataSourceName(i, True)
                count += 1
            End While

        Else
            connectionstrings(0) = FindString("<ConnectString", "</ConnectString", 0) 'If there is only one data source, simply finds the name of the one string
            If (FindString("<DataProvider", "</DataProvider", 0) <> "SQL") Then
                isanalysis(0) = True
            End If
            If (Not connectionstrings(0).Contains("Integrated Security")) Then
                connectionstrings(0) += ";Integrated Security=True"
            End If
            datasourcenames(0) = FindDataSourceName(0, True)
            datasourcenames(0) = datasourcenames(0).Remove(0, 19)
        End If

        i = -1 'Resets i and count to their original values for use in future functions
        count = 0

        SetParameters(datasourcenames, connectionstrings, numdatasources, filename)

        While count < numdatasets 'This while loop iterates once for each dataset, since each dataset has to be set and filled individually

            i = filereaderdatasets.IndexOf("<DataSet Name", i + 1) 'This block finds which data source the dataset is using, so it knows which connection string to use
            currentdatasourcename = FindDataSourceName(i, False)
            For j As Integer = 0 To (numdatasources - 1)
                If (currentdatasourcename = datasourcenames(j)) Then
                    datasource = j
                    Exit For
                End If
            Next

            query = FindString("<CommandText", "</CommandText", i) 'Records the xml data of the query

            If (isanalysis(datasource)) Then 'Checks whether or not the data source uses analysis services
                connectionstrings(datasource) = connectionstrings(datasource).Replace(";Integrated Security=True", "")

                Dim cn = New AdomdConnection(connectionstrings(datasource))

                ConnectAndFillAdomd(count, i, query, cn, datasetnames, filename) 'Connects to the data source and fills the data set
            Else
                Dim cn = New SqlConnection(connectionstrings(datasource)) 'Sets the connection using the connection string found in the last block

                ConnectAndFillSQL(count, i, query, cn, datasetnames, filename) 'Connects to the data source and fills the data set
            End If

            count += 1
        End While

        ReplaceFields(filename) 'Replaces fields, required for AS reports 
        If (Not saveparameters) Then 'Clears variables unless rendering multiple reports
            ClearGlobalVariables()
        End If
    End Sub

    Private Sub ConnectAndFillSQL(count As Integer, i As Integer, query As String, cn As SqlConnection, datasetnames As String(), filename As String)
        Dim cmd = New SqlCommand(query, cn) 'Sets the sql command

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

        If (sqlparams) Then 'If there are parameters, adds them to the command using the parameters found in the SetParameters function
            For l As Integer = 0 To (countparamssql - 1)
                If (l = 0) Then
                    cmd.Parameters.AddWithValue(paramvarsql(l), paramsql(l))
                ElseIf (CheckArray(paramvarsql, paramvarsql(l), l)) Then
                    cmd.Parameters.AddWithValue(paramvarsql(l), paramsql(l))
                End If
            Next
        End If

        Dim da = New SqlDataAdapter(cmd) 'Sets the data adapter
        Dim tbl = New DataTable()

        Try
            cn.Open()
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception
            If (renderingmultiple) Then
                wereerrors = True
                errormessages.Add("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            End If
        End Try

        datasetnames(count) = FindDataSetName(i)
        Dim rptData = New ReportDataSource(datasetnames(count), tbl) 'Sets the data set with the data table

        ReportViewer1.LocalReport.DataSources.Add(rptData) 'Adds the report data to the report

    End Sub

    Private Sub ConnectAndFillAdomd(count As Integer, i As Integer, query As String, cn As AdomdConnection, datasetnames As String(), filename As String)
        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(query, cn) 'Sets the command
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

        If (sqlparams) Then 'If there are parameters, adds them to the command using the parameters found in the SetParameters function
            For l As Integer = 0 To (countparamssql - 1)
                If (l = 0) Then
                    cmd.Parameters.Add(paramvarsql(l), paramsql(l))
                ElseIf (CheckArray(paramvarsql, paramvarsql(l), l)) Then
                    cmd.Parameters.Add(paramvarsql(l), paramsql(l))
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
                errormessages.Add("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            End If
        End Try

        datasetnames(count) = FindDataSetName(i)
        Dim rptData = New ReportDataSource(datasetnames(count), tbl) 'Sets the data set with the data table

        ReportViewer1.LocalReport.DataSources.Add(rptData) 'Adds the report data to the report

    End Sub

    Private Sub ReplaceFields(filename As String)

        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report

        Dim i As Integer = 0
        Dim failedtochange As Boolean = True
        Dim isparametercolumn As Boolean = False
        'This loop replaces each <Datafield> tag with the name of the column. This is because analysis services reports have some stuff in the data field which makes it not work with this application. However, replacing these makes it not work with the report builder, which is why the changes are saved to a temp file rather than the original file.

        For Each item In report.Root.Descendants("{http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition}Fields")
            While True 'Makes sure the dataset has datafields
                If (ReportViewer1.LocalReport.DataSources(i).Value.Columns.Count = 0) Then
                    i += 1
                Else
                    Exit While
                End If
            End While
            For Each item2 In item.Descendants("{http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition}DataField")
                Dim columnname As String = item2.Parent.FirstAttribute.Value
                failedtochange = True
                isparametercolumn = False
                For j As Integer = 0 To (ReportViewer1.LocalReport.DataSources(i).Value.Columns.Count - 1)
                    Dim valuecheck As String = ReportViewer1.LocalReport.DataSources(i).Value.Columns(j).ColumnName
                    valuecheck = ValueCheckFunction(valuecheck)
                    If (String.Compare(columnname, valuecheck, True) = 0) Then
                        item2.Value = ReportViewer1.LocalReport.DataSources(i).Value.Columns(j).ColumnName
                        failedtochange = False
                        Exit For
                    End If
                Next
                For l As Integer = 0 To (countparamsadomd - 1)
                    If (paramvaradomd(l) = columnname) Then
                        isparametercolumn = True
                    End If
                Next
                If (failedtochange And item2.Value.Contains("http://") And (Not isparametercolumn)) Then
                    If (renderingmultiple) Then
                        wereerrors = True
                        errormessages.Add("Failed to set datafield " + columnname + " in dataset " + item.Parent.FirstAttribute.Value + Environment.NewLine + "Data in report " + filename + " may not be correct. (Check for zeroes or blanks where they shouldn't be.)" + Environment.NewLine + "This could be caused by a field in the report definition that data is not returned for when the query is ran with specified parameters. It could also be caused by failing to follow standard MDX naming for data fields.")
                    Else
                        MsgBox("Failed to set datafield " + columnname + " in dataset " + item.Parent.FirstAttribute.Value + Environment.NewLine + "Data in report " + filename + " may not be correct. (Check for zeroes or blanks where they shouldn't be.)" + Environment.NewLine + "This could be caused by a field in the report definition that data is not returned for when the query is ran with specified parameters. It could also be caused by failing to follow standard MDX naming for data fields.")
                    End If
                End If
            Next
            i += 1
        Next

        report.Save(filename)
        ReportViewer1.LocalReport.ReportPath = filename

        If (sqlparams) Then 'If there are sql parameters, adds them to the report (they have to be added to the report separately from the dataset)
            For l As Integer = 0 To (countparamssql - 1)
                Dim Namespce As XNamespace = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition"
                For Each item In report.Root.Descendants(Namespce + "ReportParameter") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
                    If (item.FirstAttribute.Value = paramvarsql(l)) Then
                        Dim p As New ReportParameter(paramvarsql(l), paramsql(l))
                        Try
                            ReportViewer1.LocalReport.SetParameters(p)
                        Catch ex As Exception
                            If (renderingmultiple) Then
                                wereerrors = True
                                errormessages.Add(ex.Message + Environment.NewLine + "Error in report definition.")
                            Else
                                Dim response = MsgBox(ex.Message + Environment.NewLine + "Error in report definition." + Environment.NewLine + Environment.NewLine + "Would you like to see more information?", MsgBoxStyle.YesNo)
                                If response = MsgBoxResult.Yes Then
                                    MsgBox(ex.InnerException.InnerException.Message)
                                End If
                            End If
                        End Try
                    End If
                Next
            Next
        End If

        Dim m As Integer = 0

        If (adomdparams) Then 'If there are adomd parameters, adds them to the report (they have to be added to the report separately from the dataset)
            Dim Namespce As XNamespace = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition"
            For Each item In report.Root.Descendants(Namespce + "ReportParameter") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
                For m = 0 To (countparamsadomd - 1)
                    If (item.FirstAttribute.Value = paramvaradomd(m)) Then
                        If (adomdparamconnectionstrings(m) = "NODATASET") Then
                            Dim p As New ReportParameter(paramvaradomd(m), paramadomd(m))
                            ReportViewer1.LocalReport.SetParameters(p)
                        Else
                            Dim p As New ReportParameter(paramvaradomd(m), adomdqueryvalues(m))
                            Dim fileReader As String
                            fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                            fileReader = fileReader.Insert((fileReader.IndexOf("</ReportParameters>")), "    <ReportParameter Name=""" + paramvaradomd(m) + "Name"">
                              <DataType>String</DataType>
                              <DefaultValue>
                                <Values>
                                  <Value>" + paramadomd(m) + "</Value>
                                </Values>
                              </DefaultValue>
                              <Prompt></Prompt>
                            </ReportParameter>"
                                  )
                            fileReader = fileReader.Replace("Parameters!" + paramvaradomd(m) + ".Label", "Parameters!" + paramvaradomd(m) + "Name.Value")
                            My.Computer.FileSystem.DeleteFile(ReportViewer1.LocalReport.ReportPath)
                            My.Computer.FileSystem.WriteAllText(ReportViewer1.LocalReport.ReportPath, fileReader, True)
                            Try
                                ReportViewer1.LocalReport.SetParameters(p)
                            Catch ex As Exception
                                If (renderingmultiple) Then
                                    wereerrors = True
                                    errormessages.Add(ex.Message + Environment.NewLine + "Error in report definition.")
                                Else
                                    Dim response = MsgBox(ex.Message + Environment.NewLine + "Error in report definition." + Environment.NewLine + Environment.NewLine + "Would you like to see more information?", MsgBoxStyle.YesNo)
                                    If response = MsgBoxResult.Yes Then
                                        MsgBox(ex.InnerException.InnerException.Message)
                                    End If
                                End If
                            End Try
                        End If
                    End If
                Next
            Next
        End If
    End Sub

    Private Function ValueCheckFunction(valuecheck As String)
        If (valuecheck.Contains("[Measures].[") And (NumTimes("]", valuecheck) = 2)) Then
            valuecheck = valuecheck.Replace("[Measures].[", "")
            valuecheck = valuecheck.Replace("]", "")
            valuecheck = ReplaceWithUnderscores(valuecheck)
        ElseIf (Not valuecheck.Contains("[")) Then
            valuecheck = ReplaceWithUnderscores(valuecheck)
        ElseIf (valuecheck.Contains("[MEMBER_CAPTION]") Or valuecheck.Contains("[MEMBER_UNIQUE_NAME]")) Then
            Dim StartPos As Integer = valuecheck.IndexOf("[", 1)
            Dim EndPos As Integer = valuecheck.IndexOf("]", StartPos)
            StartPos = valuecheck.IndexOf("[", StartPos + 1)
            EndPos = valuecheck.IndexOf("]", StartPos)
            valuecheck = valuecheck.Substring(StartPos + 1, EndPos - StartPos - 1)
            valuecheck = ReplaceWithUnderscores(valuecheck)
        End If
        Return valuecheck
    End Function

    Private Sub SetParameters(datasourcenames As String(), connectionstrings As String(), numdatasources As Integer, filename As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        Dim paramdataset As String
        Dim paramvar As String
        Dim issql As Boolean = True
        Dim adomdbutnodataset As Boolean = True
        Dim yeardatasetnotset = True

        Dim Namespce As XNamespace = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition"
        For Each item In report.Root.Descendants(Namespce + "ReportParameter") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
            paramvar = item.FirstAttribute.Value
            issql = True
            adomdbutnodataset = True
            For Each item3 In item.Descendants(Namespce + "ValidValues") 'Checks each parameter to see if it's AS or SQL, then runs the appropriate function. If it's a yeardataset parameter, takes the value from DateYear
                issql = False
                For Each item2 In item3.Descendants(Namespce + "DataSetName")
                    adomdbutnodataset = False
                    paramdataset = item2.Value
                    paramdatasets.Add(paramdataset)
                    'adomdparams = True
                    'SetParametersAdomd(datasourcenames, connectionstrings, numdatasources, paramdataset, paramvar)
                    If ((String.Compare(paramdataset, "YearDataSet", True) = 0)) Then
                        yeardatasetnotset = True
                        For l As Integer = 0 To (countparamsadomd - 1)
                            If (paramvaradomd(l) = "DateYear") Then
                                paramvarsql.Add("Year")
                                paramsql.Add(paramadomd(l))
                                countparamssql += 1
                                yeardatasetnotset = False
                                sqlparams = True
                            End If
                        Next
                        If (yeardatasetnotset) Then
                            sqlparams = True
                            SetParametersSQL(paramvar)
                        End If
                    Else
                        If ((String.Compare(paramdataset, "YearDataSet", True) = 0)) Then
                            yeardatasetnotset = True
                            For l As Integer = 0 To (countparamsadomd - 1)
                                If (paramvaradomd(l) = "DateYear") Then
                                    paramvarsql.Add("Year")
                                    paramsql.Add(paramadomd(l))
                                    countparamssql += 1
                                    yeardatasetnotset = False
                                    sqlparams = True
                                End If
                            Next
                            If (yeardatasetnotset) Then
                                sqlparams = True
                                SetParametersSQL(paramvar)
                            End If
                        Else
                            adomdparams = True
                            SetParametersAdomd(datasourcenames, connectionstrings, numdatasources, paramdataset, paramvar, filename)
                        End If
                    End If
                Next
                If (adomdbutnodataset) Then
                    For Each item2 In item3.Descendants(Namespce + "Label")
                        adomdvalues.Add(item2.Value)
                    Next
                    SetParametersAdomd(datasourcenames, connectionstrings, numdatasources, "adomdbutnodataset", paramvar, filename)
                    adomdvalues.Clear()
                End If
            Next
            If (issql) Then
                sqlparams = True
                SetParametersSQL(paramvar)
            End If
        Next
    End Sub

    Private Sub SetParametersSQL(paramvar As String)
        Dim tempcount As Integer = countparamssql
        paramvarsql.Add(paramvar)
        If countparamssql = 0 Then
            Dim frm2 As Form2 = New Form2
            frm2.ShowDialog() 'Opens the 2nd form used for setting the paramete
            countparamssql += 1
        ElseIf (CheckArray(paramvarsql, paramvarsql(countparamssql), countparamssql)) Then
            Dim frm2 As Form2 = New Form2
            frm2.ShowDialog() 'Opens the 2nd form used for setting the parameter
            countparamssql += 1
        End If
        If (tempcount >= countparamssql) Then
            paramvarsql.RemoveAt(paramvaradomd.Count - 1)
        End If
    End Sub

    Private Sub SetParametersAdomd(datasourcenames As String(), connectionstrings As String(), numdatasources As Integer, paramdataset As String, paramvar As String, filename As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        Dim Namespce As XNamespace = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition"
        If (paramdataset = "adomdbutnodataset") Then
            paramvaradomd.Add(paramvar)
            adomdparamconnectionstrings.Add("NODATASET")
            If countparamsadomd = 0 Then
                Dim frm4 As Form4 = New Form4
                frm4.ShowDialog()
                countparamsadomd += 1
            ElseIf (CheckArray(paramvaradomd, paramvaradomd(countparamsadomd), countparamsadomd)) Then
                Dim frm4 As Form4 = New Form4
                frm4.ShowDialog()
                countparamsadomd += 1
            End If
        Else
            For Each item In report.Root.Descendants(Namespce + "DataSet") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
                If (item.FirstAttribute.Value = paramdataset) Then
                    Dim tempcount As Integer = countparamsadomd
                    paramvaradomd.Add(paramvar)
                    Dim fileReader As String
                    fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                    Dim i As Integer = fileReader.IndexOf("<DataSet Name=""" + paramdataset)
                    paramcommands.Add(FindString("<CommandText", "</CommandText", i))
                    Dim currentdatasourcename As String = FindDataSourceName(i, False)
                    For j As Integer = 0 To (numdatasources - 1) 'Gets the connection string for the parameter data set
                        If (currentdatasourcename = datasourcenames(j)) Then
                            adomdparamconnectionstrings.Add(connectionstrings(j))
                            If (adomdparamconnectionstrings(countparamsadomd).Contains("Integrated Security")) Then
                                adomdparamconnectionstrings(countparamsadomd) = adomdparamconnectionstrings(countparamsadomd).Replace(";Integrated Security=True", "")
                            End If
                            Exit For
                        End If
                    Next
                    If countparamsadomd = 0 Then
                        filenametemp = filename
                        Dim frm3 As Form3 = New Form3
                        frm3.ShowDialog()
                        filenametemp = ""
                        countparamsadomd += 1
                    ElseIf (CheckArray(paramvaradomd, paramvaradomd(countparamsadomd), countparamsadomd)) Then
                        filenametemp = filename
                        Dim frm3 As Form3 = New Form3
                        frm3.ShowDialog()
                        filenametemp = ""
                        countparamsadomd += 1
                    End If
                    If (tempcount >= countparamsadomd) Then
                        paramvaradomd.RemoveAt(paramvaradomd.Count - 1)
                        paramcommands.RemoveAt(paramvaradomd.Count - 1)
                    End If
                End If
            Next
        End If
    End Sub

    Private Function NumTimes(stringtofind As String, stringtosearch As String) 'Finds the number of times a string occurs in the report
        Dim num As Integer
        If (stringtosearch = "Report") Then
            stringtosearch = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
        End If
        num = Regex.Split(stringtosearch, stringtofind).Length - 1
        Return num
    End Function

    Private Function FindString(OpenTag As String, ClosingTag As String, StartingIndex As Integer) 'Used to find strings between an opening and closing tag in xml
        Dim fileReader As String
        Dim len As Integer = OpenTag.Length
        fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
        Dim StartPos As Integer = fileReader.IndexOf(OpenTag, StartingIndex)
        Dim EndPos As Integer = fileReader.IndexOf(ClosingTag, StartPos)
        Dim finaloutcome As String = fileReader.Substring(StartPos + len + 1, EndPos - StartPos - len - 1) 'After finding the index of the opening and closing tag, gets the substring between those two indexes, plus the length of the tags, because otherwise the tags would be included in the string
        finaloutcome = ReplaceEscapeCharacters(finaloutcome)
        Return finaloutcome
    End Function

    Private Function FindDataSetName(StartingIndex As Integer) 'Similar to FindString(), but specifically for data set names, since they're stored differently
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
        Dim StartPos As Integer = fileReader.IndexOf("DataSet Name=", StartingIndex)
        Dim EndPos As Integer = fileReader.IndexOf(">", StartPos)
        Dim finaloutcome As String = fileReader.Substring(StartPos + 14, EndPos - StartPos - 15)
        finaloutcome = ReplaceEscapeCharacters(finaloutcome)
        Return finaloutcome
    End Function

    Private Function FindDataSourceName(StartingIndex As Integer, bool As Boolean) 'Similar to FindDataSetName(), but for data source names
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
        Dim StartPos As Integer = fileReader.IndexOf("DataSource", StartingIndex)
        Dim EndPos As Integer = fileReader.IndexOf(">", StartPos + 15)
        Dim finaloutcome As String
        If (bool) Then 'Checks whether the data source name is in the data source definitions, or in a dataset definition
            finaloutcome = fileReader.Substring(StartPos + 17, EndPos - StartPos - 18)
        Else
            finaloutcome = fileReader.Substring(StartPos + 15, EndPos - StartPos - 31)
        End If
        finaloutcome = ReplaceEscapeCharacters(finaloutcome)
        Return finaloutcome
    End Function

    Private Function ReplaceEscapeCharacters(finaloutcome As String)
        Dim chars = New String() {"'", "&", ">", "<"}
        Dim escs = New String() {" &apos;", "&amp;", "&gt;", "&lt;"}
        For i As Integer = 0 To (escs.Length - 1)
            finaloutcome = finaloutcome.Replace(escs(i), chars(i)) 'Replaces escape characters with their normal characters
        Next
        Return finaloutcome
    End Function

    Private Function ReplaceWithUnderscores(str As String) 'Replaces characters with underscores for column names
        If (str.Contains("%") And (str.IndexOf("%") = 0)) Then
            str = str.Replace("%", "ID_")
        End If
        Dim startswithnumber As Boolean = False
        For Each m As Match In Regex.Matches(str, "([\d]+)")
            If (m.Index = 0) Then
                startswithnumber = True
            End If
        Next
        If (startswithnumber) Then
            str = str.Insert(0, "ID")
        End If
        Dim rgx As New Regex("[^a-zA-Z0-9]")
        str = rgx.Replace(str, "_")
        Return str
    End Function

    Public Class GlobalVariables 'Necessary because of these variables' use in a different form, also helpful because of how many different functions use them when parameters are involved
        'Variables for SQL parameters
        Public Shared paramsql As New List(Of String)
        Public Shared paramvarsql As New List(Of String)
        Public Shared countparamssql As Integer = 0
        Public Shared sqlparams As Boolean = False
        'Variables for Adomd parameters
        Public Shared paramadomd As New List(Of String)
        Public Shared paramvaradomd As New List(Of String)
        Public Shared adomdvalues As New List(Of String) 'For parameters with a list of values
        Public Shared adomdparamconnectionstrings As New List(Of String)
        Public Shared paramcommands As New List(Of String)
        Public Shared adomdqueryvalues As New List(Of String)
        Public Shared paramdatasets As New List(Of String)
        Public Shared countparamsadomd As Integer = 0
        Public Shared adomdparams As Boolean = False
        'Variable for rendering multiple reports into one PDF
        Public Shared combinedfilename
        'Variables for error handling
        Public Shared renderingmultiple As Boolean = False
        Public Shared wereerrors As Boolean = False
        Public Shared errormessages As New List(Of String)
        Public Shared filenametemp As String = ""
    End Class

    Private Sub ClearGlobalVariables() 'Clears all the global variables
        countparamssql = 0
        countparamsadomd = 0
        paramsql.Clear()
        paramvarsql.Clear()
        paramadomd.Clear()
        paramvaradomd.Clear()
        adomdvalues.Clear()
        paramdatasets.Clear()
        paramcommands.Clear()
        adomdqueryvalues.Clear()
        adomdparamconnectionstrings.Clear()
        adomdparams = False
        sqlparams = False
    End Sub

    Private Function CheckArray(arr As List(Of String), value As String, count As Integer) 'Makes sure the parameter is not one that has already been assigned
        For i As Integer = 0 To (count - 1)
            If (arr(i) = value) Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub DeleteFilesFromFolder() 'Deletes files from the program folder so they're not left over
        For Each fil As String In Directory.GetFiles(Application.StartupPath)
            If ((Path.GetExtension(fil) = ".rdl") Or (Path.GetExtension(fil) = ".rdlc")) Then  'Checks extension
                File.Delete(fil)
            End If
        Next
    End Sub

    Private Sub PDFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PDFToolStripMenuItem.Click
        ExportMultiple("PDF", ".pdf")
    End Sub

    Private Sub ExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem.Click
        ExportMultiple("Excel", ".xls")
    End Sub

    Private Sub WordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WordToolStripMenuItem.Click
        ExportMultiple("Word", ".doc")
    End Sub

    Private Sub ExportMultiple(filetype As String, fileextension As String) 'Exports multiple reports using selected file type
        OpenFileDialog2.Filter = "Report Files|*.rdlc;*.rdl"
        Dim dr As DialogResult = OpenFileDialog2.ShowDialog()
        If (dr = System.Windows.Forms.DialogResult.OK) Then
            FolderBrowserDialog1.ShowDialog()
            Dim filepath As String = ""
            filepath = FolderBrowserDialog1.SelectedPath
            If (filepath <> "") Then
                Dim file As String
                Dim filenum As Integer = 0
                Dim bytes As Byte()
                Dim pdfDoc As New PdfDocument()

                If (CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Checked And filetype = "PDF") Then
                    Dim frm5 As Form5 = New Form5
                    frm5.ShowDialog()
                End If

                renderingmultiple = True

                For Each file In OpenFileDialog2.FileNames 'For each file, sets the report/parameters, renders them, saves them to the drive
                    ReportViewer1.LocalReport.DataSources.Clear()
                    ReportViewer1.LocalReport.ReportPath = OpenFileDialog2.FileNames(filenum) 'Sets the report equal to the file chosen by the user
                    SetData(UseSameValuesForSameParametersToolStripMenuItem.Checked, OpenFileDialog2.SafeFileNames(filenum)) 'Sets the source for the data, finds the query, and populates the data table
                    bytes = ReportViewer1.LocalReport.Render(filetype)

                    If (CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Checked And filetype = "PDF") Then
                        Dim MS As MemoryStream = New MemoryStream(bytes)
                        Dim tempPDFDoc As PdfDocument = PdfReader.Open(MS, PdfDocumentOpenMode.Import)
                        Dim pagestorender As Integer
                        If (OnlyRenderTheFirstPageOfEachReportToolStripMenuItem.Checked) Then
                            pagestorender = 0
                        Else
                            pagestorender = tempPDFDoc.Pages.Count - 1
                        End If
                        For i As Integer = 0 To pagestorender
                            Dim page As PdfPage = tempPDFDoc.Pages(i)
                            pdfDoc.AddPage(page)
                        Next

                        If (combinedfilename.Contains(".pdf")) Then
                            pdfDoc.Save(filepath + "\" + combinedfilename)
                        Else
                            pdfDoc.Save(filepath + "\" + combinedfilename + ".pdf")
                        End If
                    Else
                        Dim filename As String
                        filename = OpenFileDialog2.SafeFileNames(filenum).Replace(".rdlc", fileextension)
                        filename = filename.Replace(".rdl", fileextension)
                        Dim fs As New FileStream(filepath + "\" + filename, FileMode.Create)
                        fs.Write(bytes, 0, bytes.Length)
                    End If

                    filenum += 1
                Next
                combinedfilename = ""
                ClearGlobalVariables()
                DeleteFilesFromFolder()
                If (wereerrors) Then
                    Dim response = MsgBox("Reports finished rendering." + Environment.NewLine + "There were errors during rendering. Would you like to view them?", MsgBoxStyle.YesNo)
                    If response = MsgBoxResult.Yes Then
                        For i As Integer = 0 To (errormessages.Count - 1)
                            MsgBox(errormessages(i))
                        Next
                    End If
                Else
                    MsgBox("Reports finished rendering.")
                End If
                renderingmultiple = False
                wereerrors = False
                errormessages.Clear()
            End If
        End If
    End Sub
End Class