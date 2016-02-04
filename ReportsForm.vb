Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine

Public Class ReportsForm

    Public ReportType As String


    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Dim reppath As String = ConfigurationManager.AppSettings("reportpath")
        Dim repsrc As String

        '   Dim rep As CrystalDecisions.CrystalReports.Engine.ReportClass
        Dim username As String = ConfigurationManager.AppSettings("username")
        Dim password As String = ConfigurationManager.AppSettings("password")
        Dim server As String = ConfigurationManager.AppSettings("server")

        Dim rep As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Select Case ReportType

            Case "ManagerTest"
                repsrc = reppath & "ManagerTest.rpt"
            Case "ManagerList"
                Dim report As New ManagerCallSheet
                repsrc = reppath & "ManagerCallSheet.rpt"
            Case "MunicipalityList"
                Dim report As New Municipalities
                repsrc = reppath & "Municipalities.rpt"
            Case "RetestList"
                Dim report As New RetestList
                repsrc = reppath & "RetestList.rpt"
            Case "RetestLetters"
                Dim report As New RetestLetters
                repsrc = reppath & "RetestLetters.rpt"
            Case "TestReports"
                repsrc = reppath & "TestReport.rpt"
            Case "TestsByManager"
                repsrc = reppath & "TestsByManager.rpt"
            Case "TestsByMunicipality"
                repsrc = reppath & "TestsByMunicipality.rpt"
            Case "TopManagers"
                repsrc = reppath & "AssemblyTestsPerManager.rpt"
            Case "TestComparisonBetweenYears"
                repsrc = reppath & "TestComparisonBetweenYears.rpt"
        End Select

        Try
            rep.Load(repsrc, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy)
            Dim coninfo As New CrystalDecisions.Shared.ConnectionInfo()
            coninfo.ServerName = server
            coninfo.Password = password
            coninfo.UserID = username

            For Each tab As CrystalDecisions.CrystalReports.Engine.Table In rep.Database.Tables
                Dim tabLogin As CrystalDecisions.Shared.TableLogOnInfo = tab.LogOnInfo
                tabLogin.ConnectionInfo = coninfo
                tab.ApplyLogOnInfo(tabLogin)
            Next


            For Each dsc As CrystalDecisions.Shared.IConnectionInfo In rep.DataSourceConnections
                dsc.SetConnection(server, "qb", False)
                dsc.SetLogon(username, password)
            Next
        Catch Ex As Exception
            Throw Ex
        End Try
        CrystalReportViewer1.ReportSource = rep

        CrystalReportViewer1.RefreshReport()

    End Sub


End Class