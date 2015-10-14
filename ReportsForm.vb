Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine

Public Class ReportsForm

    Public ReportType As String


    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Dim reppath As String = ConfigurationManager.AppSettings("reportpath")

        'CrystalReportViewer1.ReportSource = reppath & "ManagerTest.rpt"
        '        CrystalReportViewer1.ReportSource = reppath & "TestReport.rpt"

        Select Case ReportType
            Case "ManagerTest"
                CrystalReportViewer1.ReportSource = reppath & "ManagerTest.rpt"

            Case "ManagerList"
                Dim report As New ManagerCallSheet
                CrystalReportViewer1.ReportSource = reppath & "ManagerCallSheet.rpt"
            Case "MunicipalityList"
                Dim report As New Municipalities
                CrystalReportViewer1.ReportSource = reppath & "Municipalities.rpt"
            Case "RetestList"
                Dim report As New RetestList
                CrystalReportViewer1.ReportSource = reppath & "RetestList.rpt"
            Case "RetestLetters"
                Dim report As New RetestLetters
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
        CrystalReportViewer1.RefreshReport()

    End Sub


End Class