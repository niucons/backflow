Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class AssemblyDetailForm
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
    Friend WithEvents txtAssemNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents rtbAssemNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents cboSize As System.Windows.Forms.ComboBox
    Friend WithEvents cboManufacturer As System.Windows.Forms.ComboBox
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtInstallDate As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtModel As System.Windows.Forms.TextBox
    Friend WithEvents txtSerial As System.Windows.Forms.TextBox
    Friend WithEvents cboUsage As System.Windows.Forms.ComboBox
    Friend WithEvents txtTestHistory As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssemblyDetailForm))
        Me.txtAssemNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.lblNotes = New System.Windows.Forms.Label
        Me.rtbAssemNotes = New System.Windows.Forms.RichTextBox
        Me.cboSize = New System.Windows.Forms.ComboBox
        Me.cboManufacturer = New System.Windows.Forms.ComboBox
        Me.cboType = New System.Windows.Forms.ComboBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtInstallDate = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtModel = New System.Windows.Forms.TextBox
        Me.txtSerial = New System.Windows.Forms.TextBox
        Me.cboUsage = New System.Windows.Forms.ComboBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.txtTestHistory = New System.Windows.Forms.TextBox
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtAssemNo
        '
        Me.txtAssemNo.Location = New System.Drawing.Point(16, 392)
        Me.txtAssemNo.Name = "txtAssemNo"
        Me.txtAssemNo.Size = New System.Drawing.Size(88, 20)
        Me.txtAssemNo.TabIndex = 108
        Me.txtAssemNo.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 183)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 147
        Me.Label2.Text = "Location"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLocation
        '
        Me.txtLocation.HideSelection = False
        Me.txtLocation.Location = New System.Drawing.Point(120, 179)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(248, 20)
        Me.txtLocation.TabIndex = 8
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(12, 221)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(144, 16)
        Me.lblNotes.TabIndex = 145
        Me.lblNotes.Text = "Assembly Specific Notes"
        '
        'rtbAssemNotes
        '
        Me.rtbAssemNotes.Location = New System.Drawing.Point(12, 245)
        Me.rtbAssemNotes.Name = "rtbAssemNotes"
        Me.rtbAssemNotes.Size = New System.Drawing.Size(352, 128)
        Me.rtbAssemNotes.TabIndex = 9
        Me.rtbAssemNotes.Text = ""
        '
        'cboSize
        '
        Me.cboSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSize.Location = New System.Drawing.Point(120, 109)
        Me.cboSize.Name = "cboSize"
        Me.cboSize.Size = New System.Drawing.Size(152, 21)
        Me.cboSize.TabIndex = 5
        '
        'cboManufacturer
        '
        Me.cboManufacturer.Items.AddRange(New Object() {"Ames", "Conbraco", "Febco", "Hersey", "Hunter Ames", "Watts", "Wilkins"})
        Me.cboManufacturer.Location = New System.Drawing.Point(120, 62)
        Me.cboManufacturer.Name = "cboManufacturer"
        Me.cboManufacturer.Size = New System.Drawing.Size(152, 21)
        Me.cboManufacturer.TabIndex = 3
        '
        'cboType
        '
        Me.cboType.Items.AddRange(New Object() {"RP", "RPDA", "DC", "DCDA"})
        Me.cboType.Location = New System.Drawing.Point(120, 38)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(152, 21)
        Me.cboType.TabIndex = 2
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(308, 389)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 24)
        Me.btnOK.TabIndex = 10
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(244, 389)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 24)
        Me.btnCancel.TabIndex = 139
        Me.btnCancel.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 138
        Me.Label1.Text = "Install Date"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtInstallDate
        '
        Me.txtInstallDate.HideSelection = False
        Me.txtInstallDate.Location = New System.Drawing.Point(120, 156)
        Me.txtInstallDate.Name = "txtInstallDate"
        Me.txtInstallDate.Size = New System.Drawing.Size(152, 20)
        Me.txtInstallDate.TabIndex = 7
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 67)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 16)
        Me.Label10.TabIndex = 136
        Me.Label10.Text = "Manufacturer"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 137)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 135
        Me.Label7.Text = "Usage"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 114)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 134
        Me.Label6.Text = "Size"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 90)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 133
        Me.Label5.Text = "Model"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 132
        Me.Label4.Text = "Type"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 131
        Me.Label3.Text = "Serial"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtModel
        '
        Me.txtModel.HideSelection = False
        Me.txtModel.Location = New System.Drawing.Point(120, 86)
        Me.txtModel.Name = "txtModel"
        Me.txtModel.Size = New System.Drawing.Size(152, 20)
        Me.txtModel.TabIndex = 4
        '
        'txtSerial
        '
        Me.txtSerial.Location = New System.Drawing.Point(120, 15)
        Me.txtSerial.Name = "txtSerial"
        Me.txtSerial.Size = New System.Drawing.Size(152, 20)
        Me.txtSerial.TabIndex = 1
        '
        'cboUsage
        '
        Me.cboUsage.Items.AddRange(New Object() {"Boiler Feed", "Car Wash", "Filter System", "Fire Line", "Fire Line By-Pass", "Hose Bib", "Lawn Sprinkler", "Potable"})
        Me.cboUsage.Location = New System.Drawing.Point(120, 133)
        Me.cboUsage.Name = "cboUsage"
        Me.cboUsage.Size = New System.Drawing.Size(152, 21)
        Me.cboUsage.TabIndex = 6
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'txtTestHistory
        '
        Me.txtTestHistory.AcceptsReturn = True
        Me.txtTestHistory.AcceptsTab = True
        Me.txtTestHistory.BackColor = System.Drawing.SystemColors.Control
        Me.txtTestHistory.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtTestHistory.Location = New System.Drawing.Point(374, 14)
        Me.txtTestHistory.Multiline = True
        Me.txtTestHistory.Name = "txtTestHistory"
        Me.txtTestHistory.Size = New System.Drawing.Size(297, 399)
        Me.txtTestHistory.TabIndex = 148
        '
        'AssemblyDetailForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(679, 429)
        Me.Controls.Add(Me.txtTestHistory)
        Me.Controls.Add(Me.cboUsage)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtLocation)
        Me.Controls.Add(Me.txtInstallDate)
        Me.Controls.Add(Me.txtModel)
        Me.Controls.Add(Me.txtSerial)
        Me.Controls.Add(Me.txtAssemNo)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.rtbAssemNotes)
        Me.Controls.Add(Me.cboSize)
        Me.Controls.Add(Me.cboManufacturer)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AssemblyDetailForm"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Assembly Details"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Shared assemblyNumber As Integer
    Protected connectionString As String = Connection.SQL_CONNECTION_STRING
    Private strSerial As String
    Private Type As String
    Private Manufacturer As String
    Private strModel As String
    Private intSize As Integer
    Private strUsage As String
    Private strInstallDate As String
    Private strLocation As String
    Private strNotes As String


    ' Handle the Form load event.
    Private Sub AssemAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FillComboBoxes()
        FillAssemData()
    End Sub

    Private Sub FillComboBoxes()

        cboSize.ValueMember = "AssemSizeNo"
        cboSize.DisplayMember = "AssemSize"
        cboSize.DataSource = MainForm.ds.Tables("AssemSizes")

    End Sub

    Private Sub FillAssemData()
        Dim strHistory As String = "************************ Test History ************************" & ControlChars.CrLf & ControlChars.CrLf
        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(("SELECT Assemblies.*, Tests.* FROM Assemblies LEFT JOIN Tests ON Assemblies.assemNo = Tests.assemNo WHERE Assemblies.assemNo = " & Me.txtAssemNo.Text & " ORDER BY Tests.testDate DESC"), cn)
        ' Run the query; get the DataReader object.
        cn.Open()
        Dim dr As SqlDataReader = cmd.ExecuteReader()

        Do While dr.Read()
            strSerial = dr.Item("assemSerial").ToString() & ""
            Type = dr.Item("assemType").ToString() & ""
            Manufacturer = dr.Item("assemMan").ToString() & ""
            strModel = dr.Item("assemMod").ToString() & ""
            If Not IsDBNull(dr.Item("assemSizeNo")) Then
                intSize = CInt(dr.Item("assemSizeNo"))
            Else
                intSize = 1
            End If

            strUsage = dr.Item("assemUsage").ToString() & ""
            If IsDBNull(dr.Item("assemInstDt")) Then
                strInstallDate = ""
            ElseIf CDate(dr.Item("assemInstDt")) < CDate("1/2/1900") Then
                strInstallDate = ""
            Else
                strInstallDate = Format(CDate(dr.Item("assemInstDt")), "M/dd/yyyy")
            End If

            strLocation = dr.Item("assemLoc").ToString() & ""
            strNotes = dr.Item("assemNotes").ToString() & ""

            If Not IsDBNull(dr.Item("testPerformed")) Then
                strHistory &= "Test Date: " & CDate(dr.Item("testDate")) & "  Performed: " & dr.Item("testPerformed").ToString() & "  Passed: " & dr.Item("testPass").ToString() & ControlChars.CrLf
            End If

        Loop
        dr.Close()
        cn.Close()
        txtTestHistory.Text = strHistory
        Me.txtSerial.Text = strSerial
        Me.cboType.Text = Type
        Me.cboManufacturer.Text = Manufacturer
        Me.txtModel.Text = strModel
        Me.cboSize.SelectedValue = intSize
        Me.cboUsage.Text = strUsage
        Me.txtInstallDate.Text = strInstallDate
        Me.txtLocation.Text = strLocation
        Me.rtbAssemNotes.Text = strNotes

    End Sub 'FillData

    Public Sub UpdateAssembly()

        Dim UsagePriceNo As Integer
        UsagePriceNo = 1
        'If cboUsage.SelectedItem.ToString().IndexOf("Fire Line By-Pass") > -1 Then
        '    UsagePriceNo = 2
        'End If

        If cboUsage.Text = "Fire Line By-Pass" Then
            UsagePriceNo = 2
        End If

        ' Build the INSERT INTO string
        Dim strUpdate As String = "UPDATE Assemblies SET "

        strUpdate &= "assemSerial = '" & Replace(txtSerial.Text, "'", "''") & "', "
        strUpdate &= "assemType = '" & Replace(cboType.Text, "'", "''") & "', "
        strUpdate &= "assemMan = '" & Replace(cboManufacturer.Text, "'", "''") & "', "
        strUpdate &= "assemMod = '" & Replace(txtModel.Text, "'", "''") & "', "
        strUpdate &= "assemSizeNo = " & cboSize.SelectedValue.ToString() & ", "
        strUpdate &= "assemUsage = '" & Replace(cboUsage.Text, "'", "''") & "', "
        strUpdate &= "assemUsagePriceNo = " & UsagePriceNo & ", "
        strUpdate &= "assemInstDt = '" & txtInstallDate.Text & "', "
        strUpdate &= "assemLoc = '" & Replace(txtLocation.Text, "'", "''") & "', "
        strUpdate &= "assemNotes = '" & Replace(rtbAssemNotes.Text, "'", "''") & "' "
        strUpdate &= "WHERE assemNo = " & txtAssemNo.Text


        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(strUpdate, cn)
        cn.Open() ' Open the Connection
        Try
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
            DialogResult() = DialogResult.OK

        End Try

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If FormIsValid() Then
            UpdateAssembly()
        Else
            DialogResult() = DialogResult.None
        End If

    End Sub

    Private Function FormIsValid() As Boolean
        If SerialIsValid() And ModelIsValid() And UsageIsValid() And InstallDateIsValid() And TypeIsValid() And ManufacturerIsValid() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function TypeIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(cboType, "")
        If cboType.Text.Trim = "" Or cboType.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(cboType, "Type is a required field")
        End If
        Return Booly
    End Function

    Private Function ManufacturerIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(cboManufacturer, "")
        If cboManufacturer.Text.Trim = "" Or cboManufacturer.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(cboManufacturer, "Manufacturer is a required field")
        End If
        Return Booly
    End Function

    Private Function SerialIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtSerial, "")
        If txtSerial.Text.Trim = "" Or txtSerial.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtSerial, "Serial No. is a required field")
        End If
        Return Booly
    End Function

    Private Function ModelIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtModel, "")
        If txtModel.Text.Trim = "" Or txtModel.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtModel, "Model No. is a required field")
        End If
        Return Booly
    End Function

    Private Function UsageIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(cboUsage, "")
        If cboUsage.Text.Trim = "" Or cboUsage.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(cboUsage, "Usage is a required field.")
        End If
        Return Booly
    End Function

    Private Function InstallDateIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtInstallDate, "")
        If Not txtInstallDate.Text.Trim = "" And Not txtInstallDate.Text Is Nothing Then
            Dim InstallDateRegex As New Regex("(?n:^(?=\d)((?<month>(0?[13578])|1[02]|(0?[469]|11)(?!.31)|0" _
                & "?2(?(.29)(?=.29.((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][" _
                & "26])|(16|[2468][048]|[3579][26])00))|(?!.3[01])))(?<sep>[-./" _
                & "])(?<day>0?[1-9]|[12]\d|3[01])\k<sep>(?<year>(1[6-9]|[2-9]\d" _
                & ")\d{2})(?(?=\x20\d)\x20|$))?(?<time>((0?[1-9]|1[012])(:[0-5]" _
                & "\d){0,2}(?i:\x20[AP]M))|([01]\d|2[0-3])(:[0-5]\d){1,2})?$)")
            If Not InstallDateRegex.IsMatch(txtInstallDate.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtInstallDate, "Dates must be in the form of '2/12/2004'")
            End If
        End If
        Return Booly
    End Function
End Class
