Imports System.Data.SqlClient
Imports ReportsApplication.Form1.GlobalVariables
Imports Microsoft.Reporting.WinForms
Imports Microsoft.AnalysisServices.AdomdClient
Imports System.IO
Imports System.Text.RegularExpressions
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO

Public Class Form1

    Private Datasets As List(Of Dataset)
    Private filename As String

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        OnloadUI() 'Set some positionings
        ReportViewer1.RefreshReport()
        Refresh_ListView()
        AddHandler ListView1.SelectedIndexChanged, Sub(s, ea)
                                                       Dim filename = ListView1.Items(ListView1.FocusedItem.Index).Name
                                                       Open(filename)
                                                   End Sub

        Me.Text = System.Configuration.ConfigurationSettings.AppSettings.Get("title")
        'Me.ForeColor = Color.FromName(System.Configuration.ConfigurationSettings.AppSettings.Get("Background Color"))
        If Not String.IsNullOrEmpty(System.Configuration.ConfigurationSettings.AppSettings.Get("Icon").ToString()) Then
            Me.Icon = New System.Drawing.Icon(System.Configuration.ConfigurationSettings.AppSettings.Get("Icon"))
        End If
        FactbookMenuItem.Checked = Boolean.Parse(System.Configuration.ConfigurationSettings.AppSettings.Get("Factbook Default"))
    End Sub

    Public Sub Refresh_ListView()
        If Not String.IsNullOrEmpty(System.Configuration.ConfigurationSettings.AppSettings.Get("Report Directory")) Then
            Dim Folder As New System.IO.DirectoryInfo(System.Configuration.ConfigurationSettings.AppSettings.Get("Report Directory"))
            ListView1.Items.Clear()
            For Each File As System.IO.FileInfo In Folder.GetFiles("*.rdl", System.IO.SearchOption.TopDirectoryOnly)
                Dim item As New ListViewItem
                item.Text = File.Name
                item.Name = File.FullName
                ListView1.Items.Add(item)
            Next
            For Each File As System.IO.FileInfo In Folder.GetFiles("*.rdlc", System.IO.SearchOption.TopDirectoryOnly)
                Dim item As New ListViewItem
                item.Text = File.Name
                item.Name = File.FullName
                ListView1.Items.Add(item)
            Next
        End If
    End Sub

    Public Structure Dataset
        Public Name As String
        Public IsAnalysis As Boolean
        Public ConnectionString As String
        Public Query As String
        Public DataSource As String
    End Structure
    Public Structure Datasource
        Public Name As String
        Public IsAnalysis As Boolean
        Public ConnectionString As String
    End Structure
    Public Class ParameterSQL 'I know the SQL and Adomd structs are exactly the same, but I do this to make sure I don't accidentally add a parameter to the wrong list.
        Public Parameter As String
        Public ParamVar As String
        Public ConnectionString As String
        Public Query As String
        Public QueryValues As String
        Public Dataset As String
        Public DataType As String
        Public ValueField As String
        Public LabelField As String

        Public Function Clone() As ParameterSQL
            Dim copy As ParameterSQL = New ParameterSQL()
            copy.Parameter = Parameter
            copy.ParamVar = ParamVar
            copy.ConnectionString = ConnectionString
            copy.Query = Query
            copy.QueryValues = QueryValues
            copy.Dataset = Dataset
            copy.DataType = DataType
            copy.ValueField = ValueField
            copy.LabelField = LabelField
            Return copy
        End Function

        Public Overrides Function ToString() As String
            Return Parameter
        End Function

    End Class
    Public Class ParameterAdomd
        Public Parameter As String
        Public ParamVar As String
        Public ConnectionString As String
        Public Query As String
        Public QueryValues As String
        Public Dataset As String
        Public DataType As String
        Public ValueField As String
        Public LabelField As String

        Public Function Clone() As ParameterAdomd
            Dim copy As ParameterAdomd = New ParameterAdomd()
            copy.Parameter = Parameter
            copy.ParamVar = ParamVar
            copy.ConnectionString = ConnectionString
            copy.Query = Query
            copy.QueryValues = QueryValues
            copy.Dataset = Dataset
            copy.DataType = DataType
            copy.ValueField = ValueField
            copy.LabelField = LabelField
            Return copy
        End Function

        Public Overrides Function ToString() As String
            Return Parameter
        End Function

    End Class

    Private Sub Open(filename)
        FlowLayoutPanel1.Controls.Clear()
        ClearGlobalVariables()
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.ReportPath = filename 'Sets the report equal to the file chosen by the user
        Me.filename = Path.GetFileName(filename)
        SetData(False, Me.filename) 'Sets the parameters, adds data sets, replaces fields, renders report
        Text = "Report Viewer - " + Me.filename 'Places the name of the report in the control's title bar
        ReportViewer1.RefreshReport() 'Used to show the new report
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Report Files|*.rdlc;*.rdl"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.ShowDialog()
        If (OpenFileDialog1.FileName <> "") Then 'Makes sure a file was chosen, if not the dialog box just closes
            Dim directory = Path.GetDirectoryName(OpenFileDialog1.FileName)
            System.Configuration.ConfigurationSettings.AppSettings.Set("Report Directory", directory)
            Refresh_ListView()
            Open(OpenFileDialog1.FileName)
        End If
        If System.Configuration.ConfigurationSettings.AppSettings.Get("MODE") <> "Viewer" Then
            DeleteFilesFromFolder() 'Clears all .rdl and .rdlc files from the application's folder
        End If
        PictureBox1.Visible = False 'Hides loading icon
        Button2.Enabled = True
    End Sub

    Private Sub SetData(saveparameters As Boolean, filename As String)
        PictureBox1.Visible = True 'Displays loading icon
        SetVisibility()

        Dim filereaderdatasources As String = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data for data source purposes
        Dim filereaderdatasets As String = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data for data set purposes
        Dim numdatasources As Integer = NumTimes("DataSource Name", "Report") 'Finds the number of data sources in the report
        Dim numdatasets As Integer = NumTimes("DataSet Name", "Report") 'Finds the number of data sets in the report
        Dim i As Integer = -1 'Used as an index when finding strings in the report
        Dim count As Integer = 0

        Dim DataSources As List(Of Datasource) = New List(Of Datasource)

        Me.Datasets = New List(Of Dataset)

        FindNameSpace(filereaderdatasets)

        If (numdatasources = 0) Then 'If there is no data source, the report simply renders since no data has to be provided
            SetParameters(DataSources, DataSets, filename) 'Sets the parameters so they can be used when adding datasets
            ReplaceFields(filename) 'Replaces fields, required for AS reports 
            If (Not saveparameters) Then 'Clears variables unless rendering multiple reports
                ClearGlobalVariables()
            End If
            Exit Sub

        Else
            While count < numdatasources 'This while loop finds each connection string in the report if there are multiple data sources
                i = filereaderdatasources.IndexOf("<DataSource Name", i + 1)
                Dim tempdatasource As Datasource = New Datasource
                tempdatasource.ConnectionString = FindString("<ConnectString", "</ConnectString", i)
                tempdatasource.IsAnalysis = False
                If (FindString("<DataProvider", "</DataProvider", i) <> "SQL") Then 'If the data provider is not SQL, then it is Analysis Services
                    tempdatasource.IsAnalysis = True
                End If

                Dim AlreadyIntegrated As Boolean = tempdatasource.ConnectionString.Contains("Integrated Security")
                Dim SecurityTypeIntegrated As Boolean = FindString("<rd:SecurityType", "</rd:SecurityType", i) = "Integrated" And (Not tempdatasource.IsAnalysis)
                Dim UserSet As Boolean = Not String.IsNullOrEmpty((System.Configuration.ConfigurationSettings.AppSettings.Get("user")).ToString())
                Dim PasswordSet As Boolean = Not String.IsNullOrEmpty(System.Configuration.ConfigurationSettings.AppSettings.Get("pwd").ToString())


                If (Not AlreadyIntegrated Or FactbookMenuItem.Checked) And UserSet And PasswordSet Then
                    tempdatasource.ConnectionString += ";UID=" + System.Configuration.ConfigurationSettings.AppSettings.Get("user").ToString() + ";PWD=" + System.Configuration.ConfigurationSettings.AppSettings.Get("pwd").ToString()
                ElseIf Not AlreadyIntegrated And SecurityTypeIntegrated Then
                    tempdatasource.ConnectionString += ";Integrated Security=True"
                End If

                tempdatasource.Name = FindDataSourceName(i, True)
                DataSources.Add(tempdatasource)
                count += 1
            End While
        End If

        i = -1 'Resets i and count to their original values for use in future functions
        count = 0

        For k As Integer = 0 To (numdatasets - 1)
            i = filereaderdatasets.IndexOf("<DataSet Name", i + 1) 'This block finds which data source the dataset is using, so it knows which connection string to use
            Dim tempdataset As Dataset = New Dataset
            tempdataset.DataSource = FindDataSourceName(i, False)
            tempdataset.Name = FindDataSetName(i)
            tempdataset.Query = FindString("<CommandText", "</CommandText", i)
            For j As Integer = 0 To (DataSources.Count - 1)
                If (tempdataset.DataSource = DataSources(j).Name) Then
                    tempdataset.IsAnalysis = DataSources(j).IsAnalysis
                    tempdataset.ConnectionString = DataSources(j).ConnectionString
                    Exit For
                End If
            Next
            DataSets.Add(tempdataset)
        Next

        SetParameters(DataSources, DataSets, filename) 'Sets the parameters so they can be used when adding datasets

        For count = 0 To (DataSets.Count - 1) 'This while loop iterates once for each dataset, since each dataset has to be set and filled individually
            If (DataSets(count).IsAnalysis) Then 'Checks whether or not the data source uses analysis services
                Dim cn = New AdomdConnection(DataSets(count).ConnectionString) 'Sets the connection using the connection string found in the last block
                ConnectAndFillAdomd(DataSets(count), cn, filename) 'Connects to the data source and fills the data set
            Else
                Dim cn = New SqlConnection(DataSets(count).ConnectionString) 'Sets the connection using the connection string found in the last block
                ConnectAndFillSQL(DataSets(count), cn, filename) 'Connects to the data source and fills the data set
            End If
        Next

        ReplaceFields(filename) 'Replaces fields, required for AS reports 

        If (Not saveparameters And System.Configuration.ConfigurationSettings.AppSettings.Get("MODE") <> "Viewer") Then 'Clears variables unless rendering multiple reports
            ClearGlobalVariables()
        End If

        PictureBox1.Visible = False 'Hides loading icon
    End Sub

    Private Sub ConnectAndFillSQL(DataSet As Dataset, cn As SqlConnection, filename As String)

        Dim cmd = New SqlCommand(DataSet.Query, cn) With {.CommandTimeout = 120} 'Sets the sql command using the query

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

        Dim da = New SqlDataAdapter(cmd) 'Sets the data adapter
        Dim tbl = New DataTable()

        Try 'Populates the table
            cn.Open()
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception 'Catches if there's an error connecting to the dataset
            If (renderingmultiple) Then
                wereerrors = True
                errormessages.Add("Error in connection for dataset " + DataSet.Name + " in report " + filename + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for dataset " + DataSet.Name + " in report " + filename + Environment.NewLine + ex.Message)
            End If
        End Try

        Dim rptData = New ReportDataSource(DataSet.Name, tbl) 'Sets the data set with the data table

        ReportViewer1.LocalReport.DataSources.Add(rptData) 'Adds the report data to the report

    End Sub

    Private Sub ConnectAndFillAdomd(DataSet As Dataset, cn As AdomdConnection, filename As String)

        Dim da As AdomdDataAdapter = New AdomdDataAdapter()
        Dim cmd = New AdomdCommand(DataSet.Query, cn) With {.CommandTimeout = 120} 'Sets the command
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

        Try 'Populates the table
            cn.Open()
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception 'Catches if there's an error connecting to the dataset
            If (renderingmultiple) Then
                wereerrors = True
                errormessages.Add("Error in connection for dataset " + DataSet.Name + " in report " + filename + Environment.NewLine + ex.Message)
            Else
                MsgBox("Error in connection for dataset " + DataSet.Name + " in report " + filename + Environment.NewLine + ex.Message)
            End If
        End Try

        Dim rptData = New ReportDataSource(DataSet.Name, tbl) 'Sets the data set with the data table

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
                If (ReportViewer1.LocalReport.DataSources Is Nothing) Then 'If there are no fields, just increments i by 1 so it can go to the next source
                    i += 1
                ElseIf (ReportViewer1.LocalReport.DataSources(i).Value.Columns.Count = 0) Then
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
                For l As Integer = 0 To (AdomdParameters.Count() - 1) 'Checks if the column is a parameter
                    If (AdomdParameters(l).ParamVar = columnname) Then
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

        If (Not Directory.Exists("TempFiles")) Then
            Directory.CreateDirectory("TempFiles")
        End If

        report.Save("TempFiles/" + filename) 'Saves the new, changed report and switches the report to it
        ReportViewer1.LocalReport.ReportPath = "TempFiles/" + filename

        SetParameterDataTypes() 'Changes the parameters to use the datatype specified in the original report, since parameters added with code are type string by default

        report.Save("TempFiles/" + filename) 'Saves the new, changed report and switches the report to it
        ReportViewer1.LocalReport.ReportPath = "TempFiles/" + filename

        If (SQLParameters.Count() > 0) Then 'If there are sql parameters, adds them to the report (they have to be added to the report separately from the dataset)
            For l As Integer = 0 To (SQLParameters.Count() - 1)
                For Each item In report.Root.Descendants(NS + "ReportParameter")
                    If (item.FirstAttribute.Value = SQLParameters(l).ParamVar) Then
                        Dim p As New ReportParameter(SQLParameters(l).ParamVar, SQLParameters(l).Parameter)
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

        If (AdomdParameters.Count() > 0) Then 'If there are adomd parameters, adds them to the report (they have to be added to the report separately from the dataset)
            For Each item In report.Root.Descendants(NS + "ReportParameter")
                For m = 0 To (AdomdParameters.Count - 1)
                    If (item.FirstAttribute.Value = AdomdParameters(m).ParamVar) Then
                        If (AdomdParameters(m).ConnectionString = "NODATASET") Then
                            Dim p As New ReportParameter(AdomdParameters(m).ParamVar, AdomdParameters(m).Parameter)
                            ReportViewer1.LocalReport.SetParameters(p)
                        Else
                            Dim p As New ReportParameter(AdomdParameters(m).ParamVar, AdomdParameters(m).QueryValues)
                            Dim fileReader As String
                            fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                            'Adds a parameter for each adomd parameter. This is because adomd parameters in report builder have two values, their value and their label (x.value and x.label). Here, adomd parameters are not that complex. So it makes one parameter for the value and one for the label. Then replaces all references to x.label with xname.value
                            fileReader = fileReader.Insert((fileReader.IndexOf("</ReportParameters>")), "<ReportParameter Name=""" + AdomdParameters(m).ParamVar + "Name"">
                                <DataType>" + AdomdParameters(m).DataType + "</DataType>
                                  <DefaultValue>
                                    <Values>
                                      <Value>" + AdomdParameters(m).Parameter + "</Value>
                                    </Values>
                                  </DefaultValue>
                              </ReportParameter>
                            ")
                            fileReader = fileReader.Replace("Parameters!" + AdomdParameters(m).ParamVar + ".Label", "Parameters!" + AdomdParameters(m).ParamVar + "Name.Value")
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
                                <ParameterName>" + AdomdParameters(m).ParamVar + "Name</ParameterName>
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

    Public Function ValueCheckFunction(valuecheck As String) 'Converts the name of the column to how it should be using MDX naming.
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
        ElseIf (valuecheck.Contains("&")) Then
            Dim StartPos As Integer = valuecheck.IndexOf("&") + 2
            Dim EndPos As Integer = valuecheck.IndexOf("]", StartPos)
            valuecheck = valuecheck.Substring(StartPos, EndPos - StartPos)
        End If
        Return valuecheck
    End Function

    Private Sub SetParameters(DataSources As List(Of Datasource), DataSets As List(Of Dataset), filename As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        Dim paramdataset As String
        Dim paramvar As String
        Dim datatype As String
        Dim issql As Boolean = True
        Dim adomdbutnodataset As Boolean = True
        Dim alreadyset As Boolean = False
        'Dim yeardatasetnotset = True
        'Dim HSRG As Boolean = False

        For Each item In report.Root.Descendants(NS + "ReportParameter") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
            paramvar = item.FirstAttribute.Value
            adomdbutnodataset = True
            issql = True
            alreadyset = False
            datatype = "String" 'String is the default value
            For Each item4 In item.Descendants(NS + "DataType")
                datatype = item4.Value
            Next
            For Each item3 In item.Descendants(NS + "ValidValues") 'Checks each parameter to see if it's AS or SQL, then runs the appropriate function. If it's a yeardataset parameter, takes the value from DateYear
                Dim valuefield As String = ""
                Dim labelfield As String = ""
                For Each item2 In item3.Descendants(NS + "ValueField")
                    valuefield = item2.Value
                Next
                For Each item2 In item3.Descendants(NS + "LabelField")
                    labelfield = item2.Value
                Next
                For Each item2 In item3.Descendants(NS + "DataSetName")
                    adomdbutnodataset = False
                    paramdataset = item2.Value
                    For i As Integer = 0 To DataSets.Count() - 1
                        If DataSets(i).Name = paramdataset Then
                            issql = Not DataSets(i).IsAnalysis
                            Exit For
                        End If
                    Next
                    If (Not issql) Then
                        SetParametersAdomd(DataSets, paramdataset, paramvar, filename, datatype, valuefield, labelfield)
                    ElseIf (issql) Then
                        alreadyset = True
                        SetParametersSQLQuery(DataSets, paramdataset, paramvar, filename, datatype, valuefield, labelfield)
                    End If
                Next
                If (adomdbutnodataset) Then 'This is for the case where there is an ADOMD parameter that uses a list of values rather than a data set
                    For Each item2 In item3.Descendants(NS + "Value")
                        adomdvalues.Add(item2.Value)
                    Next
                    For Each item2 In item3.Descendants(NS + "Label")
                        adomdlabels.Add(item2.Value)
                    Next
                    issql = False
                    SetParametersAdomd(DataSets, "adomdbutnodataset", paramvar, filename, datatype, valuefield, labelfield)
                    adomdvalues.Clear()
                    adomdlabels.Clear()
                End If
            Next
            If (issql And (Not alreadyset)) Then
                SetParametersSQL(paramvar, datatype)
            End If
        Next
    End Sub

    Private Sub SetParametersSQL(paramvar As String, datatype As String)
        tempsqlparameter = New ParameterSQL
        tempsqlparameter.ParamVar = paramvar
        tempsqlparameter.DataType = datatype
        If (datatype = "DateTime") Then 'Checks if the parameter is datetime or not, since it requires a different form with a datetime picker
            If (CheckArray("SQL", tempsqlparameter.ParamVar, SQLParameters.Count())) Then 'If not already added
                Dim mode = System.Configuration.ConfigurationSettings.AppSettings.Get("MODE")
                If mode.ToUpper() <> "VIEWER" Then
                    Dim frm6 As Form6 = New Form6
                    frm6.ShowDialog() 'Opens the 2nd form used for setting the parameter`
                Else
                    Dim datepickerparam = New DateTimePicker
                    datepickerparam.Name = paramvar
                    ' Add Label
                    Dim label = New Label
                    label.AutoSize = True
                    label.Padding = New Padding(10, 5, 5, 10)
                    label.Text = tempsqlparameter.ParamVar
                    FlowLayoutPanel1.Controls.Add(label)
                    ' Add control to flow panel and define its handler
                    FlowLayoutPanel1.Controls.Add(datepickerparam)
                    datepickerparam.Value = DateTime.Now
                    tempsqlparameter.Parameter = datepickerparam.Value
                    SQLParameters.Add(tempsqlparameter)
                    AddHandler datepickerparam.ValueChanged, Sub(s, ea)
                                                                 Dim oldvalue As ParameterSQL
                                                                 oldvalue = SQLParameters.Find(Function(x) x.ParamVar = datepickerparam.Name)
                                                                 Dim newvalue = oldvalue
                                                                 newvalue.Parameter = datepickerparam.Value
                                                                 SQLParameters.Remove(oldvalue)
                                                                 SQLParameters.Add(newvalue)
                                                             End Sub
                End If
            End If
        Else
            If (CheckArray("SQL", tempsqlparameter.ParamVar, SQLParameters.Count())) Then 'If not already added
                Dim mode = System.Configuration.ConfigurationSettings.AppSettings.Get("MODE")
                If mode.ToUpper() <> "VIEWER" Then
                    Dim frm2 As Form2 = New Form2
                    frm2.ShowDialog() 'Opens the 2nd form used for setting the parameter
                Else
                    Dim textboxparam = New TextBox
                    textboxparam.Name = paramvar
                    ' Add Label
                    Dim label = New Label
                    label.AutoSize = True
                    label.Padding = New Padding(10, 5, 5, 10)
                    label.Text = tempsqlparameter.ParamVar
                    FlowLayoutPanel1.Controls.Add(label)
                    ' Add control to flow panel and define its handler
                    FlowLayoutPanel1.Controls.Add(textboxparam)
                    textboxparam.Text = ""
                    tempsqlparameter.Parameter = textboxparam.Text
                    SQLParameters.Add(tempsqlparameter)
                    AddHandler textboxparam.TextChanged, Sub(s, ea)
                                                             Dim oldvalue As ParameterSQL
                                                             oldvalue = SQLParameters.Find(Function(x) x.ParamVar = textboxparam.Name)
                                                             Dim newvalue = oldvalue
                                                             newvalue.Parameter = textboxparam.Text
                                                             SQLParameters.Remove(oldvalue)
                                                             SQLParameters.Add(newvalue)
                                                         End Sub
                End If
            End If
        End If
    End Sub

    Private Sub SetParametersSQLQuery(DataSets As List(Of Dataset), paramdataset As String, paramvar As String, filename As String, datatype As String, valuefield As String, labelfield As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        Dim tbl As DataTable
        Dim k As Integer
        Dim skippedone As Boolean
        For Each item In report.Root.Descendants(NS + "DataSet") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
            If (item.FirstAttribute.Value = paramdataset) Then
                tempsqlparameter = New ParameterSQL
                tempsqlparameter.ParamVar = paramvar
                tempsqlparameter.DataType = datatype
                tempsqlparameter.Dataset = paramdataset
                tempsqlparameter.ValueField = valuefield
                tempsqlparameter.LabelField = labelfield
                Dim fileReader As String
                fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                For j As Integer = 0 To (DataSets.Count() - 1) 'Gets the connection string for the parameter data set
                    If (paramdataset = DataSets(j).Name) Then
                        tempsqlparameter.Query = DataSets(j).Query
                        tempsqlparameter.ConnectionString = DataSets(j).ConnectionString
                        Exit For
                    End If
                Next
                If (CheckArray("SQL", tempsqlparameter.ParamVar, SQLParameters.Count())) Then 'If not already added
                    Dim mode = System.Configuration.ConfigurationSettings.AppSettings.Get("MODE")
                    If mode.ToUpper() <> "VIEWER" Then
                        Dim frm7 As Form7 = New Form7
                        frm7.ShowDialog()
                    Else
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
                            If tempsqlparameter.LabelField.Contains(Me.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                                j = i
                            End If
                            If tempsqlparameter.ValueField.Contains(Me.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                                k = i
                            End If
                        Next
                        skippedone = False

                        Dim combobox = New ComboBox()
                        combobox.DropDownStyle = ComboBoxStyle.DropDownList
                        For i As Integer = 0 To (tbl.Rows.Count - 1)
                            If (tbl.Rows(i).ItemArray(j).ToString <> "All") Then
                                Dim newasqlparam As ParameterSQL = tempsqlparameter.Clone()
                                newasqlparam.Parameter = tbl.Rows(i).ItemArray(j).ToString()
                                newasqlparam.QueryValues = tbl.Rows(i).ItemArray(0).ToString()
                                combobox.Items.Add(newasqlparam)
                            Else
                                skippedone = True
                            End If
                        Next
                        combobox.SelectedIndex = 0
                        Dim initialitem As ParameterSQL = combobox.SelectedItem
                        SQLParameters.Add(combobox.SelectedItem)
                        ' Add Label
                        Dim label = New Label
                        label.AutoSize = True
                        label.Padding = New Padding(10, 5, 5, 10)
                        label.Text = tempsqlparameter.ParamVar
                        FlowLayoutPanel1.Controls.Add(label)
                        ' Add control to flow panel and define its handler
                        FlowLayoutPanel1.Controls.Add(combobox)
                        AddHandler combobox.SelectedIndexChanged, Sub(s, ea)
                                                                      Dim index As Integer = combobox.SelectedIndex
                                                                      Dim param As ParameterSQL = combobox.SelectedItem
                                                                      Dim oldvalue As ParameterSQL
                                                                      oldvalue = SQLParameters.Find(Function(x) x.ParamVar = param.ParamVar)
                                                                      SQLParameters.Remove(oldvalue)
                                                                      SQLParameters.Add(param)

                                                                  End Sub
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub SetParametersAdomd(DataSets As List(Of Dataset), paramdataset As String, paramvar As String, filename As String, datatype As String, valuefield As String, labelfield As String)
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report
        Dim tbl As DataTable
        Dim k As Integer
        Dim skippedone As Boolean
        If (paramdataset = "adomdbutnodataset") Then 'For the case where there is no dataset and the parameter uses a list instead.
            tempadomdparameter = New ParameterAdomd
            tempadomdparameter.ParamVar = paramvar
            tempadomdparameter.DataType = datatype
            tempadomdparameter.Dataset = "NODATASET"
            tempadomdparameter.ValueField = valuefield
            tempadomdparameter.LabelField = labelfield
            If (CheckArray("Adomd", tempadomdparameter.ParamVar, AdomdParameters.Count())) Then 'If not already added
                Dim mode = System.Configuration.ConfigurationSettings.AppSettings.Get("MODE")
                If mode.ToUpper() <> "VIEWER" Then
                    Dim frm4 As Form4 = New Form4
                    frm4.ShowDialog()
                Else
                    Dim firstset As Boolean = True
                    Dim combobox = New ComboBox()
                    combobox.DropDownStyle = ComboBoxStyle.DropDownList
                    For i As Integer = 0 To (adomdlabels.Count - 1)
                        Dim item As New ParameterAdomd
                        item.Parameter = adomdlabels(i)
                        item.QueryValues = adomdvalues(i)
                        item.ParamVar = tempadomdparameter.ParamVar
                        item.LabelField = tempadomdparameter.LabelField
                        item.DataType = tempadomdparameter.DataType
                        item.Dataset = tempadomdparameter.Dataset
                        item.ConnectionString = tempadomdparameter.ConnectionString
                        item.Query = tempadomdparameter.Query
                        item.ValueField = tempadomdparameter.ValueField
                        combobox.Items.Add(item)
                        If firstset Then
                            combobox.Text = item.ToString()
                            firstset = False
                        End If
                    Next
                    Dim initialitem As ParameterAdomd = combobox.SelectedItem
                    tempadomdparameter.Parameter = initialitem.Parameter
                    tempadomdparameter.QueryValues = initialitem.QueryValues
                    AdomdParameters.Add(tempadomdparameter)
                    ' Add Label
                    Dim label = New Label
                    label.AutoSize = True
                    label.Padding = New Padding(10, 5, 5, 10)
                    label.Text = tempadomdparameter.ParamVar
                    FlowLayoutPanel1.Controls.Add(label)
                    ' Add control to flow panel and define its handler
                    FlowLayoutPanel1.Controls.Add(combobox)
                    AddHandler combobox.SelectedIndexChanged, Sub(s, ea)
                                                                  Dim index As Integer = combobox.SelectedIndex
                                                                  Dim item As ParameterAdomd = combobox.SelectedItem
                                                                  Dim oldvalue As ParameterAdomd
                                                                  oldvalue = AdomdParameters.Find(Function(x) x.ParamVar = item.ParamVar)
                                                                  AdomdParameters.Remove(oldvalue)
                                                                  AdomdParameters.Add(item)

                                                              End Sub
                End If
            End If
        Else
            For Each item In report.Root.Descendants(NS + "DataSet") 'Goes through the xml to find the data sets that are actually parameters to get their names and queries
                If (item.FirstAttribute.Value = paramdataset) Then
                    tempadomdparameter = New ParameterAdomd
                    tempadomdparameter.ParamVar = paramvar
                    tempadomdparameter.DataType = datatype
                    tempadomdparameter.Dataset = paramdataset
                    tempadomdparameter.ValueField = valuefield
                    tempadomdparameter.LabelField = labelfield
                    Dim fileReader As String
                    fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
                    For j As Integer = 0 To (DataSets.Count() - 1) 'Gets the connection string for the parameter data set
                        If (paramdataset = DataSets(j).Name) Then
                            tempadomdparameter.Query = DataSets(j).Query
                            tempadomdparameter.ConnectionString = DataSets(j).ConnectionString
                            Exit For
                        End If
                    Next
                    If (CheckArray("Adomd", tempadomdparameter.ParamVar, AdomdParameters.Count())) Then 'If not already added
                        Dim mode = System.Configuration.ConfigurationSettings.AppSettings.Get("MODE")
                        If mode.ToUpper() <> "VIEWER" Then
                            Dim frm3 As Form3 = New Form3
                            frm3.ShowDialog()
                        Else
                            'Label1.Text = "Enter the value for parameter " + tempadomdparameter.ParamVar.Replace("@", "")
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
                                If tempadomdparameter.LabelField.Contains(Me.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                                    j = i
                                End If
                                If tempadomdparameter.ValueField.Contains(Me.ValueCheckFunction(tbl.Columns(i).ColumnName)) Then
                                    k = i
                                End If
                            Next
                            skippedone = False

                            Dim combobox = New ComboBox()
                            combobox.DropDownStyle = ComboBoxStyle.DropDownList
                            For i As Integer = 0 To (tbl.Rows.Count - 1)
                                If (tbl.Rows(i).ItemArray(j).ToString <> "All") Then
                                    Dim newadomdparam As ParameterAdomd = tempadomdparameter.Clone()
                                    newadomdparam.Parameter = tbl.Rows(i).ItemArray(j).ToString()
                                    newadomdparam.QueryValues = tbl.Rows(i).ItemArray(2).ToString()
                                    combobox.Items.Add(newadomdparam)
                                Else
                                    skippedone = True
                                End If
                            Next
                            combobox.SelectedIndex = 0
                            Dim initialitem As ParameterAdomd = combobox.SelectedItem
                            AdomdParameters.Add(combobox.SelectedItem)
                            ' Add Label
                            Dim label = New Label
                            label.AutoSize = True
                            label.Padding = New Padding(10, 5, 5, 10)
                            label.Text = tempadomdparameter.ParamVar
                            FlowLayoutPanel1.Controls.Add(label)
                            ' Add control to flow panel and define its handler
                            FlowLayoutPanel1.Controls.Add(combobox)
                            AddHandler combobox.SelectedIndexChanged, Sub(s, ea)
                                                                          Dim index As Integer = combobox.SelectedIndex
                                                                          Dim param As ParameterAdomd = combobox.SelectedItem
                                                                          Dim oldvalue As ParameterAdomd
                                                                          oldvalue = AdomdParameters.Find(Function(x) x.ParamVar = param.ParamVar)
                                                                          AdomdParameters.Remove(oldvalue)
                                                                          AdomdParameters.Add(param)

                                                                      End Sub

                        End If
                    ElseIf (parishreports And tempadomdparameter.ParamVar = "Parish") Then 'Must reset parish for each report
                        AdomdParameters.RemoveAll(AddressOf ParishParamPredicate)
                        filenametemp = filename
                        Dim frm3 As Form3 = New Form3
                        frm3.ShowDialog()
                        filenametemp = ""
                    End If
                End If
            Next
        End If
    End Sub


    Private Sub SetParameterDataTypes()
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(ReportViewer1.LocalReport.ReportPath) 'Reads the report's xml data
        Dim report = XDocument.Load(ReportViewer1.LocalReport.ReportPath) 'This xml file is a copy of the original report

        If (SQLParameters.Count() > 0) Then 'If there are sql parameters, changes their data type to what it should be
            For l As Integer = 0 To (SQLParameters.Count() - 1)
                For Each item In report.Root.Descendants(NS + "ReportParameter")
                    If (item.FirstAttribute.Value = SQLParameters(l).ParamVar) Then
                        Dim datatypeindex = fileReader.IndexOf("<ReportParameter Name=""" + SQLParameters(l).ParamVar + """>")
                        Dim fileReadertemp1 As String = Replace(fileReader, "<DataType>String</DataType>", "<DataType>" + SQLParameters(l).DataType + "</DataType>", datatypeindex, 1)
                        Dim fileReadertemp2 As String = fileReader.Substring(datatypeindex) 'This is necessary because the above line, using the Replace method, will start at the datatypeindex (which is required) and cuts out everything before it. So instead of setting filereader equal to it, the original part of the report starting at that index is replaced by the new one.
                        fileReader = fileReader.Replace(fileReadertemp2, fileReadertemp1)
                    End If
                Next
            Next
        End If
        Dim m As Integer = 0

        If (AdomdParameters.Count() > 0) Then 'If there are adomd parameters, changes their data type to what it should be
            For Each item In report.Root.Descendants(NS + "ReportParameter")
                For m = 0 To (AdomdParameters.Count() - 1)
                    If (item.FirstAttribute.Value = AdomdParameters(m).ParamVar) Then
                        Dim datatypeindex As Integer = fileReader.IndexOf("<ReportParameter Name=""" + AdomdParameters(m).ParamVar)
                        Dim fileReadertemp1 = Replace(fileReader, "<DataType>String</DataType>", "<DataType>" + AdomdParameters(m).DataType + "</DataType>", datatypeindex, 1)
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

    Private Shared Function ParishParamPredicate(param As ParameterAdomd) As Boolean
        Return param.ParamVar = "Parish"
    End Function

    Public Class GlobalVariables 'Necessary because of these variables' use in different forms, also helpful because of how many different functions use them
        'Variables for SQL parameters
        Public Shared SQLParameters As List(Of ParameterSQL) = New List(Of ParameterSQL)
        Public Shared tempsqlparameter As ParameterSQL = New ParameterSQL
        'Variables for Adomd parameters
        Public Shared AdomdParameters As List(Of ParameterAdomd) = New List(Of ParameterAdomd)
        Public Shared tempadomdparameter As ParameterAdomd = New ParameterAdomd
        Public Shared adomdvalues As List(Of String) = New List(Of String)
        Public Shared adomdlabels As List(Of String) = New List(Of String)
        'Variable for rendering multiple reports into one PDF
        Public Shared combinedfilename As String
        'Variable for the namespace of the report
        Public Shared NS As XNamespace
        'Variables for error handling
        Public Shared renderingmultiple As Boolean = False
        Public Shared wereerrors As Boolean = False
        Public Shared errormessages As New List(Of String)
        Public Shared filenametemp As String = ""
        'For factbook parish reports
        Public Shared parishreports As Boolean = False
        Public Shared parishiteration As Integer = 0
    End Class

    Private Sub ClearGlobalVariables() 'Clears all the global variables
        SQLParameters.Clear()
        AdomdParameters.Clear()
        NS = ""
        If (Me.Datasets IsNot Nothing) Then
            Me.Datasets.Clear()
        End If
        Me.filename = ""
    End Sub

    Private Function CheckArray(whicharray As String, value As String, count As Integer) 'Makes sure the parameter is not one that has already been assigned
        If count = 0 Then
            Return True
        End If
        If (whicharray = "Adomd") Then
            For i As Integer = 0 To (count - 1)
                If (AdomdParameters(i).ParamVar = value) Then
                    Return False
                End If
            Next
        ElseIf (whicharray = "SQL") Then
            For i As Integer = 0 To (count - 1)
                If (SQLParameters(i).ParamVar = value) Then
                    Return False
                End If
            Next
        End If
        Return True
    End Function

    Private Sub DeleteFilesFromFolder() 'Deletes files from the program folder so they're not left over
        For Each fil As String In Directory.GetFiles(Application.StartupPath + "/TempFiles")
            If ((Path.GetExtension(fil) = ".rdl") Or (Path.GetExtension(fil) = ".rdlc") And Not fil.Contains(System.Configuration.ConfigurationSettings.AppSettings.Get("Default Report"))) Then  'Checks extension and that it's not the default report
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

                If (CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Checked And filetype = "PDF") Or FactbookMenuItem.Checked Then
                    Dim frm5 As Form5 = New Form5
                    frm5.ShowDialog()
                End If
                If (OpenFileDialog2.FileNames.Count > 1) Then
                    renderingmultiple = True
                End If

                Dim previousnum As Integer = 0 'Number of previous report, used to detect gaps where blank page must be inserted
                Dim Blank As PdfPage = New PdfPage
                If (FactbookMenuItem.Checked And OpenFileDialog2.SafeFileNames.Contains("Blank Report.rdl")) Then 'If doing factbook, render and store blank page
                    ReportViewer1.LocalReport.DataSources.Clear()
                    ReportViewer1.LocalReport.ReportPath = (From f In OpenFileDialog2.FileNames Where f.Contains("Blank Report.rdl"))(0) 'Sets the report equal to the blank report
                    SetData(UseSameValuesForSameParametersToolStripMenuItem.Checked, "Blank Report.rdl")
                    Dim blankbytes As Byte() = ReportViewer1.LocalReport.Render(filetype)
                    Dim BlankMS As MemoryStream = New MemoryStream(blankbytes)
                    Dim BlankDoc As PdfDocument = PdfReader.Open(BlankMS, PdfDocumentOpenMode.Import)
                    Blank = BlankDoc.Pages(0) 'Blank report is now stored in here
                End If

                For Each file In OpenFileDialog2.FileNames 'For each file, sets the report/parameters, renders them, saves them to the drive
                    Dim newnum As Integer = 0
                    If FactbookMenuItem.Checked And Not OpenFileDialog2.SafeFileNames(filenum) = "Blank Report.rdl" Then
                        newnum = Integer.Parse(OpenFileDialog2.SafeFileNames(filenum).Split("-"c)(0)) 'Gets page number
                    End If
                    If newnum = 75 And FactbookMenuItem.Checked Then 'The part of factbook where it renders same report for each parish
                        parishreports = True
                        While (parishiteration < 65) 'Must render report once for each parish + Louisiana as a whole
                            ReportViewer1.LocalReport.DataSources.Clear()
                            ReportViewer1.LocalReport.ReportPath = OpenFileDialog2.FileNames(filenum) 'Sets the report equal to the file chosen by the user
                            SetData(UseSameValuesForSameParametersToolStripMenuItem.Checked, OpenFileDialog2.SafeFileNames(filenum)) 'Sets the source for the data, finds the query, and populates the data table
                            bytes = ReportViewer1.LocalReport.Render(filetype)
                            HandleRendered(bytes, newnum, filetype, Blank, pdfDoc, filenum, filepath, fileextension, previousnum)
                            parishiteration += 1 'Determines which parish to use as parameter
                        End While
                        parishreports = False
                        parishiteration = 0
                    Else 'Normal render
                        ReportViewer1.LocalReport.DataSources.Clear()
                        ReportViewer1.LocalReport.ReportPath = OpenFileDialog2.FileNames(filenum) 'Sets the report equal to the file chosen by the user
                        SetData(UseSameValuesForSameParametersToolStripMenuItem.Checked, OpenFileDialog2.SafeFileNames(filenum)) 'Sets the source for the data, finds the query, and populates the data table
                        bytes = ReportViewer1.LocalReport.Render(filetype)
                        HandleRendered(bytes, newnum, filetype, Blank, pdfDoc, filenum, filepath, fileextension, previousnum)
                    End If
                    filenum += 1
                Next

                combinedfilename = ""
                ClearGlobalVariables()
                DeleteFilesFromFolder()
                PictureBox1.Visible = False 'Hides loading icon
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

    Private Sub HandleRendered(bytes As Byte(), newnum As Integer, filetype As String, Blank As PdfPage, ByRef pdfDoc As PdfDocument, filenum As Integer, filepath As String, fileextension As String, ByRef previousnum As Integer)
        If ((CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Checked And filetype = "PDF" Or FactbookMenuItem.Checked) And Not (FactbookMenuItem.Checked And OpenFileDialog2.SafeFileNames(filenum) = "Blank Report.rdl")) Then


            Dim MS As MemoryStream = New MemoryStream(bytes)
            Dim tempPDFDoc As PdfDocument = PdfReader.Open(MS, PdfDocumentOpenMode.Import)
            Dim pagestorender As Integer
            If (OnlyRenderTheFirstPageOfEachReportToolStripMenuItem.Checked Or FactbookMenuItem.Checked) Then
                pagestorender = 0
            Else
                pagestorender = tempPDFDoc.Pages.Count - 1
            End If
            If (FactbookMenuItem.Checked And Not (Blank Is New PdfPage)) Then 'Must check to see if blank pages should be added
                If newnum - previousnum > 1 And Not (newnum = 140) Then 'If there is a difference of greater than one and it's not from the parish reports
                    For i As Integer = 0 To newnum - previousnum - 2 'Run once for every difference greater than 1
                        pdfDoc.AddPage(Blank) 'Add to pdf
                    Next
                End If
                previousnum = newnum
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
        ElseIf (Not (FactbookMenuItem.Checked And OpenFileDialog2.SafeFileNames(filenum) = "Blank Report.rdl")) Then
            Dim filename As String
            filename = OpenFileDialog2.SafeFileNames(filenum).Replace(".rdlc", fileextension)
            filename = filename.Replace(".rdl", fileextension)
            Dim fs As New FileStream(filepath + "\" + filename, FileMode.Create)
            fs.Write(bytes, 0, bytes.Length)
            fs.Close()
        End If
    End Sub

    Private Sub PDFToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PDFToolStripMenuItem1.Click
        ExportCurrent("PDF", ".pdf")
    End Sub

    Private Sub ExcelToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem1.Click
        ExportCurrent("Excel", ".xls")
    End Sub

    Private Sub WordToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles WordToolStripMenuItem1.Click
        ExportCurrent("Word", ".doc")
    End Sub

    Private Sub ExportCurrent(filetype As String, fileextension As String) 'Exports multiple reports using selected file type
        FolderBrowserDialog1.ShowDialog()
        Dim filepath As String = ""
        filepath = FolderBrowserDialog1.SelectedPath
        If (filepath <> "") Then
            Dim filenum As Integer = 0
            Dim bytes As Byte()

            bytes = ReportViewer1.LocalReport.Render(filetype)
            Dim filename As String
            If (ReportViewer1.LocalReport.ReportPath IsNot Nothing) Then
                filename = ReportViewer1.LocalReport.ReportPath.Replace(".rdlc", fileextension)
                filename = filename.Replace(".rdl", fileextension)
            Else
                filename = "Welcome." + fileextension
            End If

            Using fs As New FileStream(filepath + "\" + filename, FileMode.Create)
                fs.Write(bytes, 0, bytes.Length)
            End Using
            ClearGlobalVariables()
            DeleteFilesFromFolder()
        End If
    End Sub

    Private Sub ReportViewer1_Load(sender As Object, e As EventArgs) Handles ReportViewer1.Load
        ReportViewer1.LocalReport.ReportPath = System.Configuration.ConfigurationSettings.AppSettings.Get("Default Report")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ReportViewer1.LocalReport.DataSources.Clear()

        For count = 0 To (Datasets.Count - 1) 'This while loop iterates once for each dataset, since each dataset has to be set and filled individually
            If (Datasets(count).IsAnalysis) Then 'Checks whether or not the data source uses analysis services
                Dim cn = New AdomdConnection(Datasets(count).ConnectionString) 'Sets the connection using the connection string found in the last block
                ConnectAndFillAdomd(Datasets(count), cn, filename) 'Connects to the data source and fills the data set
            Else
                Dim cn = New SqlConnection(Datasets(count).ConnectionString) 'Sets the connection using the connection string found in the last block
                ConnectAndFillSQL(Datasets(count), cn, filename) 'Connects to the data source and fills the data set
            End If
        Next

        ReplaceFields(filename) 'Replaces fields, required for AS reports 
        ReportViewer1.RefreshReport()
    End Sub

    Private Sub OnloadUI()
        FlowLayoutPanel1.Visible = False
        ReportViewer1.Location = New Point(ReportViewer1.Location.X, 0)
        ListView1.Location = New Point(ListView1.Location.X, 0)
        Button1.Location = New Point(Button1.Location.X, ReportViewer1.Location.Y + 1)
        Button2.Location = New Point(Button2.Location.X, ReportViewer1.Location.Y + 1)
        MenuStrip1.Location = New Point(MenuStrip1.Location.X, ReportViewer1.Location.Y + 1)
        ReportViewer1.Height = Me.Height - 40
        ListView1.Height = ReportViewer1.Height
        If System.Configuration.ConfigurationSettings.AppSettings.Get("MODE") = "Viewer" Then
            MenuStrip1.Items(0).Visible = False
            MenuStrip1.Items(2).Visible = False
        Else
            Button2.Visible = False
        End If
    End Sub

    'Sets visibility and positioning depending on the mode
    Private Sub SetVisibility()
        If (ReportViewer1.LocalReport.GetParameters().Count > 0 And System.Configuration.ConfigurationSettings.AppSettings.Get("MODE") = "Viewer") Then
            FlowLayoutPanel1.Visible = True
            ReportViewer1.Location = New Point(ReportViewer1.Location.X, FlowLayoutPanel1.Location.Y + FlowLayoutPanel1.Size.Height + 3)
            ReportViewer1.Height = 872
            ListView1.Height = ReportViewer1.Height
            ListView1.Location = New Point(ListView1.Location.X, FlowLayoutPanel1.Location.Y + FlowLayoutPanel1.Size.Height + 3)
            Button1.Location = New Point(Button1.Location.X, ReportViewer1.Location.Y + 1)
            Button2.Location = New Point(Button2.Location.X, ReportViewer1.Location.Y + 1)
            MenuStrip1.Location = New Point(MenuStrip1.Location.X, ReportViewer1.Location.Y + 1)
        Else
            FlowLayoutPanel1.Visible = False
            ReportViewer1.Location = New Point(ReportViewer1.Location.X, 0)
            ListView1.Location = New Point(ListView1.Location.X, 0)
            Button1.Location = New Point(Button1.Location.X, ReportViewer1.Location.Y + 1)
            Button2.Location = New Point(Button2.Location.X, ReportViewer1.Location.Y + 1)
            MenuStrip1.Location = New Point(MenuStrip1.Location.X, ReportViewer1.Location.Y + 1)
        End If
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DeleteFilesFromFolder() 'Clears all .rdl and .rdlc files from the application's folder
    End Sub
End Class