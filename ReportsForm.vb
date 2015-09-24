Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration



Public Class ReportsForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportsForm))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(1012, 533)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'ReportsForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1012, 533)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ReportsForm"
        Me.ShowInTaskbar = False
        Me.Text = "Backflow Testing Reporting"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public ReportType As String

    'Protected connectionString As String = Connection.SQL_CONNECTION_STRING

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Dim reppath As String = ConfigurationManager.AppSettings("reportpath")


        CrystalReportViewer1.ReportSource = reppath & "RetestLetters.rpt"

        Select Case ReportType
            Case "ManagerTest"
                CrystalReportViewer1.ReportSource = reppath & "ManagerTest.rpt"

            Case "ManagerList"
                '   Dim report As New ManagerCallSheet
                CrystalReportViewer1.ReportSource = reppath & "ManagerCallSheet.rpt"
            Case "ManagerList"
                '    Dim report As New ManagerCallSheet
                CrystalReportViewer1.ReportSource = reppath & "ManagerCallSheet.rpt"
            Case "MunicipalityList"
                '  Dim report As New Municipalities
                CrystalReportViewer1.ReportSource = reppath & "Municipalities.rpt"
            Case "RetestList"
                ' Dim report As New RetestList
                CrystalReportViewer1.ReportSource = reppath & "RetestList.rpt"
            Case "RetestLetters"
                ' Dim report As New RetestLetters
                CrystalReportViewer1.ReportSource = reppath & "RetestLetters.rpt"
            Case "TestReports"
                CrystalReportViewer1.ReportSource = reppath & "TestReport.rpt"
            Case "TestsByManager"
                CrystalReportViewer1.ReportSource = reppath & "TestsByManager.rpt"
            Case "TestsByMunicipality"
                CrystalReportViewer1.ReportSource = reppath & "TestsByMunicipality.rpt"
            Case "TopManagers"
                CrystalReportViewer1.ReportSource = reppath & "AssemblyTestsPerManager.rpt"
            Case "TestComparisonBetweenYears"
                CrystalReportViewer1.ReportSource = reppath & "TestComparisonBetweenYears.rpt"
        End Select
    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
