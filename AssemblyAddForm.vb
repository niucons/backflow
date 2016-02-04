Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions

Public Class AssemblyAddForm
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
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboSize As System.Windows.Forms.ComboBox
    Friend WithEvents cboManufacturer As System.Windows.Forms.ComboBox
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents rtbAssemNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtInstallDate As System.Windows.Forms.TextBox
    Friend WithEvents txtModel As System.Windows.Forms.TextBox
    Friend WithEvents txtSerial As System.Windows.Forms.TextBox
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents cboUsage As System.Windows.Forms.ComboBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssemblyAddForm))
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
        Me.lblNotes = New System.Windows.Forms.Label
        Me.rtbAssemNotes = New System.Windows.Forms.RichTextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.cboUsage = New System.Windows.Forms.ComboBox
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboSize
        '
        Me.cboSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSize.Location = New System.Drawing.Point(120, 112)
        Me.cboSize.Name = "cboSize"
        Me.cboSize.Size = New System.Drawing.Size(152, 21)
        Me.cboSize.TabIndex = 4
        '
        'cboManufacturer
        '
        Me.cboManufacturer.Items.AddRange(New Object() {"Ames", "Conbraco", "Febco", "Hersey", "Hunter Ames", "Watts", "Wilkins"})
        Me.cboManufacturer.Location = New System.Drawing.Point(120, 65)
        Me.cboManufacturer.Name = "cboManufacturer"
        Me.cboManufacturer.Size = New System.Drawing.Size(152, 21)
        Me.cboManufacturer.TabIndex = 2
        '
        'cboType
        '
        Me.cboType.Items.AddRange(New Object() {"RP", "RPDA", "DC", "DCDA", "PVB"})
        Me.cboType.Location = New System.Drawing.Point(120, 41)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(152, 21)
        Me.cboType.TabIndex = 1
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(308, 392)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 24)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(244, 392)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 24)
        Me.btnCancel.TabIndex = 119
        Me.btnCancel.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 163)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 118
        Me.Label1.Text = "Install Date"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtInstallDate
        '
        Me.txtInstallDate.HideSelection = False
        Me.txtInstallDate.Location = New System.Drawing.Point(120, 159)
        Me.txtInstallDate.Name = "txtInstallDate"
        Me.txtInstallDate.Size = New System.Drawing.Size(152, 20)
        Me.txtInstallDate.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 70)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 16)
        Me.Label10.TabIndex = 116
        Me.Label10.Text = "Manufacturer"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 140)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 115
        Me.Label7.Text = "Usage"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 117)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 114
        Me.Label6.Text = "Size"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 93)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 113
        Me.Label5.Text = "Model"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 112
        Me.Label4.Text = "Type"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 111
        Me.Label3.Text = "Serial"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtModel
        '
        Me.txtModel.HideSelection = False
        Me.txtModel.Location = New System.Drawing.Point(120, 89)
        Me.txtModel.Name = "txtModel"
        Me.txtModel.Size = New System.Drawing.Size(152, 20)
        Me.txtModel.TabIndex = 3
        '
        'txtSerial
        '
        Me.txtSerial.Location = New System.Drawing.Point(120, 18)
        Me.txtSerial.Name = "txtSerial"
        Me.txtSerial.Size = New System.Drawing.Size(152, 20)
        Me.txtSerial.TabIndex = 0
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(12, 224)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(144, 16)
        Me.lblNotes.TabIndex = 125
        Me.lblNotes.Text = "Assembly Specific Notes"
        '
        'rtbAssemNotes
        '
        Me.rtbAssemNotes.Location = New System.Drawing.Point(12, 248)
        Me.rtbAssemNotes.Name = "rtbAssemNotes"
        Me.rtbAssemNotes.Size = New System.Drawing.Size(352, 128)
        Me.rtbAssemNotes.TabIndex = 8
        Me.rtbAssemNotes.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 186)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 127
        Me.Label2.Text = "Location"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLocation
        '
        Me.txtLocation.HideSelection = False
        Me.txtLocation.Location = New System.Drawing.Point(120, 182)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(248, 20)
        Me.txtLocation.TabIndex = 7
        '
        'cboUsage
        '
        Me.cboUsage.Items.AddRange(New Object() {"Boiler Feed", "Car Wash", "Filter System", "Fire Line", "Fire Line By-Pass", "Hose Bib", "Lawn Sprinkler", "Potable"})
        Me.cboUsage.Location = New System.Drawing.Point(120, 136)
        Me.cboUsage.Name = "cboUsage"
        Me.cboUsage.Size = New System.Drawing.Size(152, 21)
        Me.cboUsage.TabIndex = 5
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'AssemblyAddForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 429)
        Me.Controls.Add(Me.cboUsage)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtLocation)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.rtbAssemNotes)
        Me.Controls.Add(Me.cboSize)
        Me.Controls.Add(Me.cboManufacturer)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtInstallDate)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtModel)
        Me.Controls.Add(Me.txtSerial)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AssemblyAddForm"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Assembly"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Shared propNumber As Integer
    Protected connectionString As String = ConfigurationManager.ConnectionStrings("bftm.My.MySettings.qbConnectionString1").ConnectionString

    ' Handle the Form load event.
    Private Sub AssemAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FillComboBoxes()
    End Sub

    Private Sub FillComboBoxes()

        cboSize.ValueMember = "AssemSizeNo"
        cboSize.DisplayMember = "AssemSize"
        cboSize.DataSource = MainForm.ds.Tables("AssemSizes")

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If FormIsValid() Then
            AddAssembly()
        Else
            DialogResult = DialogResult.None
        End If
    End Sub

    Public Sub AddAssembly()

        Dim UsagePriceNo As Integer = 1
        If cboUsage.Text.Trim = "Fire Line By-Pass" Then
            UsagePriceNo = 2
        End If

        ' Build the INSERT INTO string
        Dim strInsert As String = "INSERT INTO Assemblies " _
            & "(propNo, assemSerial, assemType, assemMan, assemMod, assemSizeNo, assemUsage, " _
            & "assemUsagePriceNo, assemInstDt, assemLoc, assemNotes) " _
            & "VALUES ("

        strInsert &= propNumber & ", "
        strInsert &= "'" & Replace(Me.txtSerial.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.cboType.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.cboManufacturer.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtModel.Text, "'", "''") & "', "
        strInsert &= Me.cboSize.SelectedValue.ToString() & ", "
        strInsert &= "'" & Replace(Me.cboUsage.Text, "'", "''") & "', "
        strInsert &= CStr(UsagePriceNo) & ", "
        strInsert &= "'" & Me.txtInstallDate.Text & "', "
        strInsert &= "'" & Replace(Me.txtLocation.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.rtbAssemNotes.Text, "'", "''") & "')"


        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(strInsert, cn)
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
