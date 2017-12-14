Imports System.Data.SqlClient
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
            SetData(False, OpenFileDialog1.SafeFileName) 'Sets the parameters, adds data sets, replaces fields, renders report
            ReportViewer1.RefreshReport() 'Used to show the new report
            Text = "Report Viewer - " + OpenFileDialog1.SafeFileName 'Places the name of the report in the control's title bar
        End If
        DeleteFilesFromFolder() 'Clears all .rdl and .rdlc files from the application's folder
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
        Dim currentdatasourcename As String 'Used to find data sources later
        Dim query As String 'Later contains the command text for data sets
        Dim isanalysisdatasources(numdatasources - 1) As Boolean 'An array used to determine if each datasource in datasources is analysis services
        Dim isanalysisdatasets(numdatasets - 1) As Boolean 'An array used to determine if each dataset's datasource is analysis services
        Dim datasetsdatasource(numdatasets - 1) As Integer 'An array that stores which data source a data set uses
        Dim datasetqueries(numdatasets - 1) As String 'An array for data set queries

        FindNameSpace(filereaderdatasets)

        If (numdatasources = 0) Then 'If there is no data source, the report simply renders since no data has to be provided
            SetParameters(datasourcenames, connectionstrings, numdatasources, filename, isanalysisdatasets, numdatasets, datasetnames) 'Sets parameters, as even if a report has no data it may have parameters
            ReplaceFields(filename) 'Replaces fields, required for AS reports 
            If (Not saveparameters) Then 'Clears variables unless rendering multiple reports
                ClearGlobalVariables()
            End If
            Exit Sub

        Else
            While count < numdatasources 'This while loop finds each connection string in the report if there are multiple data sources
                i = filereaderdatasources.IndexOf("<DataSource Name", i + 1)
                connectionstrings(count) = FindString("<ConnectString", "</ConnectString", i)
                If (FindString("<DataProvider", "</DataProvider", i) <> "SQL") Then 'If the data provider is not SQL, then it is Analysis Services
                    isanalysisdatasources(count) = True
                End If
                If (Not connectionstrings(count).Contains("Integrated Security")) Then 'Adds integrated security if it's not already there
                    connectionstrings(count) += ";Integrated Security=True"
                End If
                datasourcenames(count) = FindDataSourceName(i, True)
                count += 1
            End While
        End If

        i = -1 'Resets i and count to their original values for use in future functions
        count = 0

        For k As Integer = 0 To (numdatasets - 1)
            i = filereaderdatasets.IndexOf("<DataSet Name", i + 1) 'This block finds which data source the dataset is using, so it knows which connection string to use
            currentdatasourcename = FindDataSourceName(i, False)
            datasetnames(k) = FindDataSetName(i)
            datasetqueries(k) = FindString("<CommandText", "</CommandText", i)
            For j As Integer = 0 To (numdatasources - 1)
                If (currentdatasourcename = datasourcenames(j)) Then
                    isanalysisdatasets(k) = isanalysisdatasources(j)
                    datasetsdatasource(k) = j
                    Exit For
                End If
            Next
        Next
        SetParameters(datasourcenames, connectionstrings, numdatasources, filename, isanalysisdatasets, numdatasets, datasetnames) 'Sets the parameters so they can be used when adding datasets

        While count < numdatasets 'This while loop iterates once for each dataset, since each dataset has to be set and filled individually

            If (isanalysisdatasets(count)) Then 'Checks whether or not the data source uses analysis services
                connectionstrings(datasetsdatasource(count)) = connectionstrings(datasetsdatasource(count)).Replace(";Integrated Security=True", "") 'Removes integrated security, since it doesn't work with AS

                Dim cn = New AdomdConnection(connectionstrings(datasetsdatasource(count))) 'Sets the connection using the connection string found in the last block

                ConnectAndFillAdomd(count, datasetqueries(count), cn, datasetnames, filename) 'Connects to the data source and fills the data set
            Else
                Dim cn = New SqlConnection(connectionstrings(datasetsdatasource(count))) 'Sets the connection using the connection string found in the last block

                ConnectAndFillSQL(count, datasetqueries(count), cn, datasetnames, filename) 'Connects to the data source and fills the data set
            End If

            count += 1
        End While

        ReplaceFields(filename) 'Replaces fields, required for AS reports 

        If (Not saveparameters) Then 'Clears variables unless rendering multiple reports
            ClearGlobalVariables()
        End If
    End Sub

    Private Sub ConnectAndFillSQL(count As Integer, query As String, cn As SqlConnection, datasetnames As String(), filename As String)

        Dim cmd = New SqlCommand(query, cn) With {.CommandTimeout = 120} 'Sets the sql command using the query

        If (sqlparams) Then 'If there are parameters, adds them to the command using the parameters found in the SetParameters function
            For l As Integer = 0 To (countparamssql - 1)
                If (l = 0) Then
                    cmd.Parameters.AddWithValue(paramvarsql(l), paramsql(l))
                ElseIf (CheckArray(paramvarsql, paramvarsql(l), l)) Then
                    cmd.Parameters.AddWithValue(paramvarsql(l), paramsql(l))
                End If
            Next
        End If

        If (adomdparams) Then 'It's possible for a SQL connection to use AS parameters, so they too must be added
            For l As Integer = 0 To (countparamsadomd - 1)
                If (adomdparamconnectionstrings(l) = "NODATASET") Then
                    cmd.Parameters.AddWithValue(paramvaradomd(l), paramadomd(l))
                Else
                    cmd.Parameters.AddWithValue(paramvaradomd(l), adomdqueryvalues(l))
                End If
            Next
        End If

        Dim da = New SqlDataAdapter(cmd) 'Sets the data adapter
        Dim tbl = New DataTable()

        Try 'Populates the table
            cn.Open()
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception 'Catches if there's an error connecting to the dataset
            If (renderingmultiple) Then
                wereerrors = True
                errormessages.Add("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            End If
        End Try

        Dim rptData = New ReportDataSource(datasetnames(count), tbl) 'Sets the data set with the data table

        ReportViewer1.LocalReport.DataSources.Add(rptData) 'Adds the report data to the report

    End Sub

    Private Sub ConnectAndFillAdomd(count As Integer, query As String, cn As AdomdConnection, datasetnames As String(), filename As String)

        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(query, cn) With {.CommandTimeout = 120} 'Sets the command
        Dim tbl = New DataTable

        If (sqlparams) Then 'It's possible for an AS connection to use SQL parameters, so they too must be added
            For l As Integer = 0 To (countparamssql - 1)
                If (l = 0) Then
                    Dim p As New AdomdParameter(paramvarsql(l), paramsql(l))
                    cmd.Parameters.Add(p)
                ElseIf (CheckArray(paramvarsql, paramvarsql(l), l)) Then
                    Dim p As New AdomdParameter(paramvarsql(l), paramsql(l))
                    cmd.Parameters.Add(p)
                End If
            Next
        End If

        If (adomdparams) Then 'If there are parameters, adds them to the command using the parameters found in the SetParameters function
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

        Try 'Populates the table
            cn.Open()
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception 'Catches if there's an error connecting to the dataset
            If (renderingmultiple) Then
                wereerrors = True
                errormessages.Add("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for dataset " + datasetnames(count) + " in report " + filename + Environment.NewLine + ex.Message)
            End If
        End Try

        Dim rptData = New ReportDataSource(datasetnames(count), tbl) 'Sets the data set with the data table

        ReportViewer1.LocalReport.DataSources.Add(rptData) 'Adds the report data to the report

    End Sub

    Private Sub ReplaceFields(filename As String)

        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report

        Dim i As Integer = 0
        Dim failedtochange As Boolean = True
        Dim isparametercolumn As Boolean = False
        'This loop replaces each <Datafield> tag with the name of the column. This is because analysis services reports have some stuff in the data field which makes it not work with this application. However, replacing these makes it not work with the report builder, which is why the changes are saved to a temp file rather than the original file.

        For Each item In report.Root.Descendants(NS + "Fields")
            While True 'Makes sure the dataset has datafields
                If (ReportViewer1.LocalReport.DataSources(i).Value.Columns.Count = 0) Then 'If there are no fields, just increments i by 1 so it can go to the next source
                    i += 1
                Else
                    Exit While
                End If
            End While
            For Each item2 In item.Descendants(NS + "DataField")
                Dim columnname As String = item2.Parent.FirstAttribute.Value
                failedtochange = True
                isparametercolumn = False
                For j As Integer = 0 To (ReportViewer1.LocalReport.DataSources(i).Value.Columns.Count - 1)
                    Dim valuecheck As String = ReportViewer1.LocalReport.DataSources(i).Value.Columns(j).ColumnName
                    valuecheck = ValueCheckFunction(valuecheck)
                    If (String.Compare(columnname, valuecheck, True) = 0) Then 'Checks if the column name is what it should be
                        item2.Value = ReportViewer1.LocalReport.DataSources(i).Value.Columns(j).ColumnName 'If it is, renames the field for rendering
                        failedtochange = False
                        Exit For
                    End If
                Next
                For l As Integer = 0 To (countparamsadomd - 1) 'Checks if the column is a parameter
                    If (paramvaradomd(l) = columnname) Then
                        isparametercolumn = True
                    End If
                Next
                If (failedtochange And item2.Value.Contains("http://") And (Not isparametercolumn)) Then 'If it didn't change it, and the field contains http (meaning it's not what it should be,) and it's not a parameter column (since the fields of parameter data sets don't matter), gives an error.
                    If (renderingmultiple) Then
                        wereerrors = True
                        errormessages.Add("Failed to set datafield " + columnname + " in dataset " + item.Parent.FirstAttribute.Value + Environment.NewLine + "Data in report " + filename + " may not be correct. (Check for zeroes or blanks where they shouldn't be.)" + Environment.NewLine + "This could be caused by a field in the report definition that data is not returned for when the query is ran with specified parameters. It could also be caused by failing to follow standard MDX naming for data fields.")
                    Else
                        MsgBox("Failed to set datafield " + columnname + " in dataset " + item.Parent.FirstAttribute.Value + Environment.NewLine + "Data in report " + filename + " may not be correct. (Check for zeroes or blanks where they shouldn't be.)" + Environment.NewLine + "This could be caused by a field in the report definition that data is not returned for when the query is ran with specified parameters. It could also be caused by failing to follow standard MDX naming for data fields.")
                    End If
                End If
            Next
            i += 1 'Increments i by 1 for the next data set
        Next

        report.Save(filename) 'Saves the new, changed report and switches the report to it
        ReportViewer1.LocalReport.ReportPath = filename

        SetParameterDataTypes() 'Changes the parameters to use the datatype specified in the original report, since parameters added with code are type string by default

        report.Save(filename) 'Saves the new, changed report and switches the report to it
        ReportViewer1.LocalReport.ReportPath = filename

        If (sqlparams) Then 'If there are sql parameters, adds them to the report (they have to be added to the report separately from the dataset)
            For l As Integer = 0 To (countparamssql - 1)
                For Each item In report.Root.Descendants(NS + "ReportParameter")
                    If (item.FirstAttribute.Value = paramvarsql(l)) Then
                        Dim p As New ReportParameter(paramvarsql(l), paramsql(l))
                        Try 'During testing this was a common place for errors. Not because of the parameters, but because if you try to add parameters to a report which has a bad definition, it will give an error.
                            ReportViewer1.LocalReport.SetParameters(p)
                        Catch ex As Exception
                            If (renderingmultiple) Then
                                wereerrors = True
                                errormessages.Add(ex.Message + Environment.NewLine + "Error in report definition.")
                            Else
                                Dim response = MsgBox(ex.Message + Environment.NewLine + "Error in report definition." + Environment.NewLine + Environment.NewLine + "Would you like to see more information?", MsgBoxStyle.YesNo)
                                If response = MsgBoxResult.Yes Then
                                    MsgBox(ex.InnerException.Message)
                                End If
                            End If
                        End Try
                    End If
                Next
            Next
        End If

        Dim m As Integer = 0

        If (adomdparams) Then 'If there are adomd parameters, adds them to the report (they have to be added to the report separately from the dataset)
            For Each item In report.Root.Descendants(NS + "ReportParameter")
                For m = 0 To (countparamsadomd - 1)
                    If (item.FirstAttribute.Value = paramvaradomd(m)) Then
                        If (adomdparamconnectionstrings(m) = "NODATASET") Then
                            Dim p As New ReportParameter(paramvaradomd(m), paramadomd(m))
                            ReportViewer1.LocalReport.SetParameters(p)
                        Else
                            Dim p As New ReportParameter(paramvaradomd(m), adomdqueryvalues(m))
                            Dim fileReader As String
                            fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                            'Adds a parameter for each adomd parameter. This is because adomd parameters in report builder have two values, their value and their label (x.value and x.label). Here, adomd parameters are not that complex. So it makes one parameter for the value and one for the label. Then replaces all references to x.label with xname.value
                            fileReader = fileReader.Insert((fileReader.IndexOf("</ReportParameters>")), "<ReportParameter Name=""" + paramvaradomd(m) + "Name"">
                                <DataType>" + adomddatatypes(m) + "</DataType>
                                  <DefaultValue>
                                    <Values>
                                      <Value>" + paramadomd(m) + "</Value>
                                    </Values>
                                  </DefaultValue>
                              </ReportParameter>
                            ")
                            fileReader = fileReader.Replace("Parameters!" + paramvaradomd(m) + ".Label", "Parameters!" + paramvaradomd(m) + "Name.Value")
                            'Microsoft's 2016 report builder added a section for parameter layouts. Despite the application not using these, they must exist for each parameter. So this block adds one to the number of columns, then adds the required information for the parameter
                            If fileReader.Contains("<ReportParametersLayout>") Then
                                Dim i1 As Integer = fileReader.IndexOf("<NumberOfColumns>") + 17
                                Dim i2 As Integer = fileReader.IndexOf("<", i1)
                                Dim columnsstring As String = fileReader.Substring(i1, i2 - i1)
                                Dim columns As Integer = Integer.Parse(columnsstring)
                                fileReader = fileReader.Replace("<NumberOfColumns>" + columns.ToString() + "</NumberOfColumns>", "<NumberOfColumns>" + (columns + 1).ToString() + "</NumberOfColumns>")
                                fileReader = fileReader.Insert((fileReader.IndexOf("</CellDefinitions>")), "<CellDefinition>
                                <ColumnIndex>" + columns.ToString() + "</ColumnIndex>
                                <RowIndex>0</RowIndex>
                                <ParameterName>" + paramvaradomd(m) + "Name</ParameterName>
                            </CellDefinition>
                            ")
                            End If

                            My.Computer.FileSystem.DeleteFile(ReportViewer1.LocalReport.ReportPath) 'Deletes and rewrites the report with its new values
                            My.Computer.FileSystem.WriteAllText(ReportViewer1.LocalReport.ReportPath, fileReader, True)
                            Try 'During testing this was a common place for errors. Not because of the parameters, but because if you try to add parameters to a report which has a bad definition, it will give an error.
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

    Private Function ValueCheckFunction(valuecheck As String) 'Converts the name of the column to how it should be using MDX naming.
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

    Private Sub SetParameters(datasourcenames As String(), connectionstrings As String(), numdatasources As Integer, filename As String, isanalysis As Boolean(), numdatasets As Integer, datasetnames As String())
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        Dim paramdataset As String
        Dim paramvar As String
        Dim datatype As String
        Dim issql As Boolean = True
        Dim adomdbutnodataset As Boolean = True
        Dim yeardatasetnotset = True
        Dim HSRG As Boolean = False

        For Each item In report.Root.Descendants(NS + "ReportParameter") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
            paramvar = item.FirstAttribute.Value
            adomdbutnodataset = True
            datatype = "String" 'String is the default value
            For Each item4 In item.Descendants(NS + "DataType")
                datatype = item4.Value
            Next
            For Each item3 In item.Descendants(NS + "ValidValues") 'Checks each parameter to see if it's AS or SQL, then runs the appropriate function. If it's a yeardataset parameter, takes the value from DateYear
                For Each item2 In item3.Descendants(NS + "DataSetName")
                    adomdbutnodataset = False
                    paramdataset = item2.Value
                    adomdparamdatasets.Add(paramdataset)
                    For i As Integer = 0 To numdatasources - 1
                        If datasourcenames(i) = "CRSHDWHSRG" Then
                            HSRG = True
                            Exit For
                        End If
                    Next
                    For i As Integer = 0 To numdatasets - 1
                        If datasetnames(i) = paramdataset Then
                            issql = Not isanalysis(i)
                            Exit For
                        End If
                    Next
                    If (HSRG And (Not issql)) Then 'This is for HSRG specifically since we assume YearDataSet == DateYear. Otherwise it will ask for the value as normal
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
                                SetParametersSQL(paramvar, datatype)
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
                                    SetParametersSQL(paramvar, datatype)
                                End If
                            Else
                                adomdparams = True
                                SetParametersAdomd(datasourcenames, connectionstrings, numdatasources, paramdataset, paramvar, filename, datatype)
                            End If
                        End If
                    ElseIf (Not issql) Then
                        adomdparams = True
                        SetParametersAdomd(datasourcenames, connectionstrings, numdatasources, paramdataset, paramvar, filename, datatype)
                    ElseIf (issql) Then
                        sqlparams = True
                        SetParametersSQLQuery(datasourcenames, connectionstrings, numdatasources, paramdataset, paramvar, filename, datatype)
                    End If
                Next
                If (adomdbutnodataset) Then 'This is for the case where there is an ADOMD parameter that uses a list of values rather than a data set
                    For Each item2 In item3.Descendants(NS + "Label")
                        adomdvalues.Add(item2.Value)
                    Next
                    SetParametersAdomd(datasourcenames, connectionstrings, numdatasources, "adomdbutnodataset", paramvar, filename, datatype)
                    adomdvalues.Clear()
                End If
            Next
            If (issql) Then
                sqlparams = True
                SetParametersSQL(paramvar, datatype)
            End If
        Next
    End Sub

    Private Sub SetParametersSQL(paramvar As String, datatype As String)
        Dim tempcount As Integer = countparamssql 'Saves the original value of countparamssql
        paramvarsql.Add(paramvar) 'Adds the parameter before incrementing countparamssql
        sqldatatypes.Add(datatype)
        If (datatype = "DateTime") Then 'Checks if the parameter is datetime or not, since it requires a different form with a datetime picker
            If countparamssql = 0 Then 'If there are no parameters currently just adds it. Otherwise checks to make sure the parameter doesn't already exist.
                Dim frm6 As Form6 = New Form6
                frm6.ShowDialog() 'Opens the 2nd form used for setting the parameter
                countparamssql += 1
            ElseIf (CheckArray(paramvarsql, paramvarsql(countparamssql), countparamssql)) Then
                Dim frm6 As Form6 = New Form6
                frm6.ShowDialog() 'Opens the 2nd form used for setting the parameter
                countparamssql += 1
            End If
        Else
            If countparamssql = 0 Then 'If there are no parameters currently just adds it. Otherwise checks to make sure the parameter doesn't already exist.
                Dim frm2 As Form2 = New Form2
                frm2.ShowDialog() 'Opens the 2nd form used for setting the parameter
                countparamssql += 1
            ElseIf (CheckArray(paramvarsql, paramvarsql(countparamssql), countparamssql)) Then
                Dim frm2 As Form2 = New Form2
                frm2.ShowDialog() 'Opens the 2nd form used for setting the parameter
                countparamssql += 1
            End If
        End If
        If (tempcount >= countparamssql) Then 'If countparamssql was not incremented, then that parameter must have already existed, so it removes the copy
            paramvarsql.RemoveAt(paramvarsql.Count - 1)
            sqldatatypes.RemoveAt(sqldatatypes.Count - 1)
        End If
    End Sub

    Private Sub SetParametersSQLQuery(datasourcenames As String(), connectionstrings As String(), numdatasources As Integer, paramdataset As String, paramvar As String, filename As String, datatype As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        For Each item In report.Root.Descendants(NS + "DataSet") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
            If (item.FirstAttribute.Value = paramdataset) Then
                Dim tempcount As Integer = countparamssql 'Saves the original value of countparamssql
                paramvarsql.Add(paramvar) 'Adds the parameter before incrementing countparamssql
                sqldatatypes.Add(datatype)

                Dim fileReader As String
                fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                Dim i As Integer = fileReader.IndexOf("<DataSet Name=""" + paramdataset)
                sqlparamcommands.Add(FindString("<CommandText", "</CommandText", i)) 'Finds the command for the parameter's data set
                Dim currentdatasourcename As String = FindDataSourceName(i, False)
                For j As Integer = 0 To (numdatasources - 1) 'Gets the connection string for the parameter data set
                    If (currentdatasourcename = datasourcenames(j)) Then
                        sqlparamconnectionstrings.Add(connectionstrings(j))
                        Exit For
                    End If
                Next
                If countparamssql = 0 Then 'If there are no parameters currently just adds it. Otherwise checks to make sure the parameter doesn't already exist.
                    filenametemp = filename
                    Dim frm7 As Form7 = New Form7
                    frm7.ShowDialog()
                    filenametemp = ""
                    countparamssql += 1
                ElseIf (CheckArray(paramvarsql, paramvarsql(countparamssql), countparamssql)) Then
                    filenametemp = filename
                    Dim frm7 As Form7 = New Form7
                    frm7.ShowDialog()
                    filenametemp = ""
                    countparamssql += 1
                End If
                If (tempcount >= countparamssql) Then 'If countparamssql was not incremented, then that parameter must have already existed, so it removes the copy and all its attributes
                    paramvarsql.RemoveAt(paramvarsql.Count - 1)
                    sqlparamconnectionstrings.RemoveAt(sqlparamconnectionstrings.Count - 1)
                    sqldatatypes.RemoveAt(sqldatatypes.Count - 1)
                    sqlparamdatasets.RemoveAt(sqlparamdatasets.Count - 1)
                    sqlparamcommands.RemoveAt(sqlparamcommands.Count - 1)
                End If
            End If
        Next
    End Sub

    Private Sub SetParametersAdomd(datasourcenames As String(), connectionstrings As String(), numdatasources As Integer, paramdataset As String, paramvar As String, filename As String, datatype As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        If (paramdataset = "adomdbutnodataset") Then 'For the case where there is no dataset and the parameter uses a list instead.
            paramvaradomd.Add(paramvar)
            adomddatatypes.Add(datatype)
            adomdparamconnectionstrings.Add("NODATASET") 'Since it uses no data set
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
            For Each item In report.Root.Descendants(NS + "DataSet") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
                If (item.FirstAttribute.Value = paramdataset) Then
                    Dim tempcount As Integer = countparamsadomd 'Saves the original value of countparamsadomd
                    paramvaradomd.Add(paramvar) 'Adds the parameter before incrementing countparamsadomd
                    adomddatatypes.Add(datatype)

                    Dim fileReader As String
                    fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                    Dim i As Integer = fileReader.IndexOf("<DataSet Name=""" + paramdataset)
                    adomdparamcommands.Add(FindString("<CommandText", "</CommandText", i)) 'Finds the command for the parameter's data set
                    Dim currentdatasourcename As String = FindDataSourceName(i, False)
                    For j As Integer = 0 To (numdatasources - 1) 'Gets the connection string for the parameter data set
                        If (currentdatasourcename = datasourcenames(j)) Then
                            adomdparamconnectionstrings.Add(connectionstrings(j))
                            If (adomdparamconnectionstrings(countparamsadomd).Contains("Integrated Security")) Then 'Takes away integrated security since it doesn't work with AS connection strings
                                adomdparamconnectionstrings(countparamsadomd) = adomdparamconnectionstrings(countparamsadomd).Replace(";Integrated Security=True", "")
                            End If
                            Exit For
                        End If
                    Next
                    If countparamsadomd = 0 Then 'If there are no parameters currently just adds it. Otherwise checks to make sure the parameter doesn't already exist.
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
                    If (tempcount >= countparamsadomd) Then 'If countparamsadomd was not incremented, then that parameter must have already existed, so it removes the copy and all its attributes
                        paramvaradomd.RemoveAt(paramvaradomd.Count - 1)
                        adomdparamconnectionstrings.RemoveAt(adomdparamconnectionstrings.Count - 1)
                        adomddatatypes.RemoveAt(adomddatatypes.Count - 1)
                        adomdparamdatasets.RemoveAt(adomdparamdatasets.Count - 1)
                        adomdparamcommands.RemoveAt(adomdparamcommands.Count - 1)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub SetParameterDataTypes()
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report

        If (sqlparams) Then 'If there are sql parameters, changes their data type to what it should be
            For l As Integer = 0 To (countparamssql - 1)
                For Each item In report.Root.Descendants(NS + "ReportParameter")
                    If (item.FirstAttribute.Value = paramvarsql(l)) Then
                        Dim datatypeindex = fileReader.IndexOf("<ReportParameter Name=""" + paramvarsql(l) + """>")
                        Dim fileReadertemp1 As String = Replace(fileReader, "<DataType>String</DataType>", "<DataType>" + sqldatatypes(l) + "</DataType>", datatypeindex, 1)
                        Dim fileReadertemp2 As String = fileReader.Substring(datatypeindex) 'This is necessary because the above line, using the Replace method, will start at the datatypeindex (which is required) and cuts out everything before it. So instead of setting filereader equal to it, the original part of the report starting at that index is replaced by the new one.
                        fileReader = fileReader.Replace(fileReadertemp2, fileReadertemp1)
                    End If
                Next
            Next
        End If
        Dim m As Integer = 0

        If (adomdparams) Then 'If there are adomd parameters, changes their data type to what it should be
            For Each item In report.Root.Descendants(NS + "ReportParameter")
                For m = 0 To (countparamsadomd - 1)
                    If (item.FirstAttribute.Value = paramvaradomd(m)) Then
                        Dim datatypeindex As Integer = fileReader.IndexOf("<ReportParameter Name=""" + paramvaradomd(m))
                        Dim fileReadertemp1 = Replace(fileReader, "<DataType>String</DataType>", "<DataType>" + adomddatatypes(m) + "</DataType>", datatypeindex, 1)
                        Dim fileReadertemp2 As String = fileReader.Substring(datatypeindex)  'This is necessary because the above line, using the Replace method, will start at the datatypeindex (which is required) and cuts out everything before it. So instead of setting filereader equal to it, the original part of the report starting at that index is replaced by the new one.
                        fileReader = fileReader.Replace(fileReadertemp2, fileReadertemp1)
                    End If
                Next
            Next
        End If

        My.Computer.FileSystem.DeleteFile(ReportViewer1.LocalReport.ReportPath) 'Deletes and rewrites the report with its new values
        My.Computer.FileSystem.WriteAllText(ReportViewer1.LocalReport.ReportPath, fileReader, True)

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
        Dim escs = New String() {"&apos;", "&amp;", "&gt;", "&lt;"}
        For i As Integer = 0 To (escs.Length - 1)
            finaloutcome = finaloutcome.Replace(escs(i), chars(i)) 'Replaces escape characters with their normal characters
        Next
        Return finaloutcome
    End Function

    Private Function ReplaceWithUnderscores(str As String) 'Replaces characters with underscores for MDX column names; used in ValueCheckFunction
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

    Public Class GlobalVariables 'Necessary because of these variables' use in different forms, also helpful because of how many different functions use them
        'Variables for SQL parameters
        Public Shared paramsql As New List(Of String)
        Public Shared paramvarsql As New List(Of String)
        Public Shared sqlparamconnectionstrings As New List(Of String)
        Public Shared sqlparamcommands As New List(Of String)
        Public Shared sqlqueryvalues As New List(Of String)
        Public Shared sqlparamdatasets As New List(Of String)
        Public Shared countparamssql As Integer = 0
        Public Shared sqlparams As Boolean = False
        Public Shared sqldatatypes As New List(Of String)
        'Variables for Adomd parameters
        Public Shared paramadomd As New List(Of String)
        Public Shared paramvaradomd As New List(Of String)
        Public Shared adomdvalues As New List(Of String) 'For parameters with a list of values
        Public Shared adomdparamconnectionstrings As New List(Of String)
        Public Shared adomdparamcommands As New List(Of String)
        Public Shared adomdqueryvalues As New List(Of String)
        Public Shared adomdparamdatasets As New List(Of String)
        Public Shared countparamsadomd As Integer = 0
        Public Shared adomdparams As Boolean = False
        Public Shared adomddatatypes As New List(Of String)
        'Variable for rendering multiple reports into one PDF
        Public Shared combinedfilename
        'Variable for the namespace of the report
        Public Shared NS As XNamespace
        'Variables for error handling
        Public Shared renderingmultiple As Boolean = False
        Public Shared wereerrors As Boolean = False
        Public Shared errormessages As New List(Of String)
        Public Shared filenametemp As String = ""
    End Class

    Private Sub ClearGlobalVariables() 'Clears all the global variables
        paramsql.Clear()
        paramvarsql.Clear()
        sqlparamconnectionstrings.Clear()
        sqlparamcommands.Clear()
        sqlqueryvalues.Clear()
        sqlparamdatasets.Clear()
        countparamssql = 0
        sqlparams = False
        sqldatatypes.Clear()
        paramadomd.Clear()
        paramvaradomd.Clear()
        adomdvalues.Clear()
        adomdparamconnectionstrings.Clear()
        adomdparamcommands.Clear()
        adomdqueryvalues.Clear()
        adomdparamdatasets.Clear()
        countparamsadomd = 0
        adomdparams = False
        adomddatatypes.Clear()
        NS = ""
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

    Private Sub FindNameSpace(report As String) 'Finds the namespace of the report, which depends on what version of Microsoft's report builder the report was made in
        Dim i1 As Integer = report.IndexOf("xmlns=") + 7
        Dim i2 As Integer = report.IndexOf("""", i1)
        NS = report.Substring(i1, i2 - i1)
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
                If (OpenFileDialog2.FileNames.Count > 1) Then
                    renderingmultiple = True
                End If

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
                    Dim response = MsgBox("Reports finished rendering." + Environment.NewLine + "There were " + errormessages.Count().ToString() + " errors during rendering." + Environment.NewLine + "These may or may not have affected the outcome of your reports." + Environment.NewLine + "Would you like to view the errors?", MsgBoxStyle.YesNo)
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