Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class CompanyInfoForm
    Protected connectionString As String = Connection.SQL_CONNECTION_STRING

    Private Sub CompanyInfoForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand("SELECT name, phone, testKit, calibrationDate, cccdi FROM Company", cn)

        Try
            Dim name As String = ""
            Dim phone As String = ""
            Dim testKit As String = ""
            Dim calibrationDate As Date
            Dim cccdi As String = ""


            ' Run the query; get the DataReader object.
            cn.Open()
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            If dr.Read() Then
                Name = dr.Item("name").ToString() & ""
                phone = dr.Item("phone").ToString() & ""
                testKit = dr.Item("testKit").ToString() & ""
                calibrationDate = CDate(dr.Item("calibrationDate"))
                cccdi = dr.Item("cccdi").ToString() & ""
            End If
            dr.Close()

            Me.txtName.Text = name
            Me.txtPhone.Text = phone
            Me.txtTestKit.Text = testKit
            Me.dtpCalibration.Value = calibrationDate
            Me.txtCCCDI.Text = cccdi



        Catch ex As SqlException
            'A SqlException has occured - display details
            Dim i As Integer, msg As String
            For i = 0 To ex.Errors.Count - 1
                Dim sqlErr As SqlError = ex.Errors(i)
                msg = "Message = " & sqlErr.Message & ControlChars.CrLf
                msg &= "Source = " & sqlErr.Source & ControlChars.CrLf
            Next
            MessageBox.Show(msg)

        Catch ex As Exception
            ' A generic exception has occured
            MessageBox.Show(ex.Message)
        Finally
            ' Close the connection
            cn.Close()
        End Try

 
    End Sub

    Private Sub btnCompanyInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyInfo.Click
        SaveCompanyInfo()
    End Sub

    Private Sub SaveCompanyInfo()
        Dim CalibrationDate As Date = dtpCalibration.Value.Date

        Dim strUpdate As String = "UPDATE Company " _
            & "SET name='" & Replace(Me.txtName.Text, "'", "''") & "', " _
            & "phone='" & Replace(Me.txtPhone.Text, "'", "''") & "', " _
            & "testKit='" & Replace(Me.txtTestKit.Text, "'", "''") & "', " _
            & "calibrationDate='" & CalibrationDate & "', " _
            & "cccdi='" & Replace(Me.txtCCCDI.Text, "'", "''") & "' " _
            & "WHERE id = 1"

            ' Me.rtbManNotes.Text = strUpdate
            Dim cn As New SqlConnection(connectionString)
            Dim cmd As New SqlCommand(strUpdate, cn)

            Try
                cn.Open()
                cmd.ExecuteNonQuery()
            Catch ex As SqlException
                'A SqlException has occured - display details
                Dim i As Integer, msg As String
                For i = 0 To ex.Errors.Count - 1
                    Dim sqlErr As SqlError = ex.Errors(i)
                    msg = "Message = " & sqlErr.Message & ControlChars.CrLf
                    msg &= "Source = " & sqlErr.Source & ControlChars.CrLf
                Next
                ' MessageBox.Show(msg)
            Catch ex As Exception
                ' A generic exception has occured
                MessageBox.Show(ex.Message)
            Finally
                ' Close the connection
                cn.Close()
            End Try

        DialogResult() = DialogResult.OK

    End Sub
End Class