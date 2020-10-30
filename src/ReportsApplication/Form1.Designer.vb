<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PDFToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WordToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.PDFToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcelToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.WordToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportOptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UseSameValuesForSameParametersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OnlyRenderTheFirstPageOfEachReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FactbookMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(546, 64)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Choose Report"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Title = "Select Report"
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.ReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ReportViewer1.DocumentMapWidth = 1
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "ReportsApplication.Report1.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(126, 63)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.ServerReport.BearerToken = Nothing
        Me.ReportViewer1.ShowBackButton = False
        Me.ReportViewer1.ShowExportButton = False
        Me.ReportViewer1.ShowFindControls = False
        Me.ReportViewer1.Size = New System.Drawing.Size(882, 872)
        Me.ReportViewer1.TabIndex = 0
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.Multiselect = True
        Me.OpenFileDialog2.Title = "Select Reports"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(434, 444)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(300, 100)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(648, 64)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(49, 23)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "RUN"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(1, -3)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1007, 65)
        Me.FlowLayoutPanel1.TabIndex = 12
        '
        'ListView1
        '
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(1, 63)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(125, 872)
        Me.ListView1.TabIndex = 13
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.List
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExportToolStripMenuItem, Me.ExportToolStripMenuItem1, Me.ExportOptionsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(700, 63)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(307, 24)
        Me.MenuStrip1.TabIndex = 14
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ExportToolStripMenuItem
        '
        Me.ExportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PDFToolStripMenuItem, Me.ExcelToolStripMenuItem, Me.WordToolStripMenuItem})
        Me.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem"
        Me.ExportToolStripMenuItem.Size = New System.Drawing.Size(107, 20)
        Me.ExportToolStripMenuItem.Text = "Open and Export"
        '
        'PDFToolStripMenuItem
        '
        Me.PDFToolStripMenuItem.Name = "PDFToolStripMenuItem"
        Me.PDFToolStripMenuItem.Size = New System.Drawing.Size(103, 22)
        Me.PDFToolStripMenuItem.Text = "PDF"
        '
        'ExcelToolStripMenuItem
        '
        Me.ExcelToolStripMenuItem.Name = "ExcelToolStripMenuItem"
        Me.ExcelToolStripMenuItem.Size = New System.Drawing.Size(103, 22)
        Me.ExcelToolStripMenuItem.Text = "Excel"
        '
        'WordToolStripMenuItem
        '
        Me.WordToolStripMenuItem.Name = "WordToolStripMenuItem"
        Me.WordToolStripMenuItem.Size = New System.Drawing.Size(103, 22)
        Me.WordToolStripMenuItem.Text = "Word"
        '
        'ExportToolStripMenuItem1
        '
        Me.ExportToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PDFToolStripMenuItem1, Me.ExcelToolStripMenuItem1, Me.WordToolStripMenuItem1})
        Me.ExportToolStripMenuItem1.Name = "ExportToolStripMenuItem1"
        Me.ExportToolStripMenuItem1.Size = New System.Drawing.Size(95, 20)
        Me.ExportToolStripMenuItem1.Text = "Export Current"
        '
        'PDFToolStripMenuItem1
        '
        Me.PDFToolStripMenuItem1.Name = "PDFToolStripMenuItem1"
        Me.PDFToolStripMenuItem1.Size = New System.Drawing.Size(103, 22)
        Me.PDFToolStripMenuItem1.Text = "PDF"
        '
        'ExcelToolStripMenuItem1
        '
        Me.ExcelToolStripMenuItem1.Name = "ExcelToolStripMenuItem1"
        Me.ExcelToolStripMenuItem1.Size = New System.Drawing.Size(103, 22)
        Me.ExcelToolStripMenuItem1.Text = "Excel"
        '
        'WordToolStripMenuItem1
        '
        Me.WordToolStripMenuItem1.Name = "WordToolStripMenuItem1"
        Me.WordToolStripMenuItem1.Size = New System.Drawing.Size(103, 22)
        Me.WordToolStripMenuItem1.Text = "Word"
        '
        'ExportOptionsToolStripMenuItem
        '
        Me.ExportOptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UseSameValuesForSameParametersToolStripMenuItem, Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem, Me.FactbookMenuItem})
        Me.ExportOptionsToolStripMenuItem.Name = "ExportOptionsToolStripMenuItem"
        Me.ExportOptionsToolStripMenuItem.Size = New System.Drawing.Size(97, 20)
        Me.ExportOptionsToolStripMenuItem.Text = "Export Options"
        '
        'UseSameValuesForSameParametersToolStripMenuItem
        '
        Me.UseSameValuesForSameParametersToolStripMenuItem.Checked = True
        Me.UseSameValuesForSameParametersToolStripMenuItem.CheckOnClick = True
        Me.UseSameValuesForSameParametersToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UseSameValuesForSameParametersToolStripMenuItem.Name = "UseSameValuesForSameParametersToolStripMenuItem"
        Me.UseSameValuesForSameParametersToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.UseSameValuesForSameParametersToolStripMenuItem.Text = "Use same values for same parameters in different reports"
        '
        'CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem
        '
        Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.CheckOnClick = True
        Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OnlyRenderTheFirstPageOfEachReportToolStripMenuItem})
        Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Name = "CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem"
        Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem.Text = "Combine multiple reports into one (PDF only)"
        '
        'OnlyRenderTheFirstPageOfEachReportToolStripMenuItem
        '
        Me.OnlyRenderTheFirstPageOfEachReportToolStripMenuItem.CheckOnClick = True
        Me.OnlyRenderTheFirstPageOfEachReportToolStripMenuItem.Name = "OnlyRenderTheFirstPageOfEachReportToolStripMenuItem"
        Me.OnlyRenderTheFirstPageOfEachReportToolStripMenuItem.Size = New System.Drawing.Size(285, 22)
        Me.OnlyRenderTheFirstPageOfEachReportToolStripMenuItem.Text = "Only render the first page of each report"
        '
        'FactbookMenuItem
        '
        Me.FactbookMenuItem.CheckOnClick = True
        Me.FactbookMenuItem.Name = "FactbookMenuItem"
        Me.FactbookMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.FactbookMenuItem.Text = "Export Factbook"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1009, 943)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report Viewer"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents OpenFileDialog2 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Button2 As Button
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents ListView1 As ListView
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ExportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PDFToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents WordToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents PDFToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ExcelToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents WordToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ExportOptionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents UseSameValuesForSameParametersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CombineMultipleReportsIntoOnePDFOnlyToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OnlyRenderTheFirstPageOfEachReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FactbookMenuItem As ToolStripMenuItem
End Class
