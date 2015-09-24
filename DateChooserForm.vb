Imports System.Data.SqlClient

Public Class DateChooserForm
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
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DateChooserForm))
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'dtpStart
        '
        Me.dtpStart.CalendarForeColor = System.Drawing.SystemColors.GrayText
        Me.dtpStart.CalendarTitleBackColor = System.Drawing.Color.LightSlateGray
        Me.dtpStart.Location = New System.Drawing.Point(96, 24)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(232, 20)
        Me.dtpStart.TabIndex = 0
        '
        'dtpEnd
        '
        Me.dtpEnd.CalendarForeColor = System.Drawing.SystemColors.GrayText
        Me.dtpEnd.CalendarTitleBackColor = System.Drawing.Color.LightSlateGray
        Me.dtpEnd.Location = New System.Drawing.Point(96, 72)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.Size = New System.Drawing.Size(232, 20)
        Me.dtpEnd.TabIndex = 1
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(128, 120)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 24)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Start Date:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 24)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Finish Date:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(200, 120)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 24)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'frmDateChsr
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 157)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dtpEnd)
        Me.Controls.Add(Me.dtpStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDateChsr"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Date Chooser"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected connectionString As String = Connection.SQL_CONNECTION_STRING

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim StartDate As Date = dtpStart.Value.Date
        Dim EndDate As Date = dtpEnd.Value.Date

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cmd.CommandText = "DELETE FROM Dates WHERE Record = 1"

        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cmd.CommandText = "INSERT INTO Dates (Record, StartDate, FinishDate, SingleDate) VALUES (1,'" & StartDate & "','" & EndDate & "','" & EndDate & "')"
            cmd.ExecuteNonQuery()
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

        Me.DialogResult = DialogResult.OK
    End Sub
End Class
