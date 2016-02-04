Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class frmManAdd
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
    Friend WithEvents lblManNotes As System.Windows.Forms.Label
    Friend WithEvents rtbManNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtManOffice As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkManActive As System.Windows.Forms.CheckBox
    Friend WithEvents txtManFaxNo As System.Windows.Forms.TextBox
    Friend WithEvents txtManPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtManContact As System.Windows.Forms.TextBox
    Friend WithEvents txtManZipCode As System.Windows.Forms.TextBox
    Friend WithEvents txtManCity As System.Windows.Forms.TextBox
    Friend WithEvents txtManStrtAdd As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboState As System.Windows.Forms.ComboBox
    Friend WithEvents lblPriceList As System.Windows.Forms.Label
    Friend WithEvents grpPricing As System.Windows.Forms.GroupBox
    Friend WithEvents txtPriceList As System.Windows.Forms.TextBox
    Friend WithEvents lblPricingScheme As System.Windows.Forms.Label
    Friend WithEvents btnRight As System.Windows.Forms.Button
    Friend WithEvents btnLeft As System.Windows.Forms.Button
    Friend WithEvents txtManName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtManEmail As System.Windows.Forms.TextBox
    Friend WithEvents btnLogo As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblLogoPath As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmManAdd))
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblManNotes = New System.Windows.Forms.Label
        Me.rtbManNotes = New System.Windows.Forms.RichTextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtManOffice = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkManActive = New System.Windows.Forms.CheckBox
        Me.txtManFaxNo = New System.Windows.Forms.TextBox
        Me.txtManPhone = New System.Windows.Forms.TextBox
        Me.txtManContact = New System.Windows.Forms.TextBox
        Me.txtManZipCode = New System.Windows.Forms.TextBox
        Me.txtManCity = New System.Windows.Forms.TextBox
        Me.txtManStrtAdd = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.cboState = New System.Windows.Forms.ComboBox
        Me.lblPriceList = New System.Windows.Forms.Label
        Me.grpPricing = New System.Windows.Forms.GroupBox
        Me.txtPriceList = New System.Windows.Forms.TextBox
        Me.lblPricingScheme = New System.Windows.Forms.Label
        Me.btnRight = New System.Windows.Forms.Button
        Me.btnLeft = New System.Windows.Forms.Button
        Me.txtManName = New System.Windows.Forms.TextBox
        Me.txtManEmail = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.btnLogo = New System.Windows.Forms.Button
        Me.lblLogoPath = New System.Windows.Forms.Label
        Me.grpPricing.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(256, 448)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 24)
        Me.btnOK.TabIndex = 13
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(192, 448)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 24)
        Me.btnCancel.TabIndex = 14
        Me.btnCancel.Text = "Cancel"
        '
        'lblManNotes
        '
        Me.lblManNotes.Location = New System.Drawing.Point(13, 288)
        Me.lblManNotes.Name = "lblManNotes"
        Me.lblManNotes.Size = New System.Drawing.Size(144, 16)
        Me.lblManNotes.TabIndex = 64
        Me.lblManNotes.Text = "Manager Specific Notes"
        '
        'rtbManNotes
        '
        Me.rtbManNotes.Location = New System.Drawing.Point(8, 304)
        Me.rtbManNotes.Name = "rtbManNotes"
        Me.rtbManNotes.Size = New System.Drawing.Size(352, 128)
        Me.rtbManNotes.TabIndex = 11
        Me.rtbManNotes.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(40, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 16)
        Me.Label10.TabIndex = 63
        Me.Label10.Text = "Office No."
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtManOffice
        '
        Me.txtManOffice.HideSelection = False
        Me.txtManOffice.Location = New System.Drawing.Point(104, 112)
        Me.txtManOffice.Name = "txtManOffice"
        Me.txtManOffice.Size = New System.Drawing.Size(152, 20)
        Me.txtManOffice.TabIndex = 6
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(272, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 16)
        Me.Label8.TabIndex = 62
        Me.Label8.Text = "State"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 184)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 61
        Me.Label7.Text = "Fax No."
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 160)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 60
        Me.Label6.Text = "Phone No."
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 59
        Me.Label5.Text = "Contact Person"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 58
        Me.Label4.Text = "Postal Code"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 57
        Me.Label3.Text = "City"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 56
        Me.Label2.Text = "Street Address"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkManActive
        '
        Me.chkManActive.AccessibleName = "asdfasdf"
        Me.chkManActive.Location = New System.Drawing.Point(385, 424)
        Me.chkManActive.Name = "chkManActive"
        Me.chkManActive.Size = New System.Drawing.Size(104, 24)
        Me.chkManActive.TabIndex = 12
        Me.chkManActive.Text = "Active Account"
        '
        'txtManFaxNo
        '
        Me.txtManFaxNo.HideSelection = False
        Me.txtManFaxNo.Location = New System.Drawing.Point(104, 184)
        Me.txtManFaxNo.Name = "txtManFaxNo"
        Me.txtManFaxNo.Size = New System.Drawing.Size(152, 20)
        Me.txtManFaxNo.TabIndex = 9
        '
        'txtManPhone
        '
        Me.txtManPhone.HideSelection = False
        Me.txtManPhone.Location = New System.Drawing.Point(104, 160)
        Me.txtManPhone.Name = "txtManPhone"
        Me.txtManPhone.Size = New System.Drawing.Size(152, 20)
        Me.txtManPhone.TabIndex = 8
        '
        'txtManContact
        '
        Me.txtManContact.HideSelection = False
        Me.txtManContact.Location = New System.Drawing.Point(104, 136)
        Me.txtManContact.Name = "txtManContact"
        Me.txtManContact.Size = New System.Drawing.Size(152, 20)
        Me.txtManContact.TabIndex = 7
        '
        'txtManZipCode
        '
        Me.txtManZipCode.Location = New System.Drawing.Point(104, 88)
        Me.txtManZipCode.Name = "txtManZipCode"
        Me.txtManZipCode.Size = New System.Drawing.Size(152, 20)
        Me.txtManZipCode.TabIndex = 5
        '
        'txtManCity
        '
        Me.txtManCity.Location = New System.Drawing.Point(104, 64)
        Me.txtManCity.Name = "txtManCity"
        Me.txtManCity.Size = New System.Drawing.Size(152, 20)
        Me.txtManCity.TabIndex = 3
        '
        'txtManStrtAdd
        '
        Me.txtManStrtAdd.Location = New System.Drawing.Point(104, 40)
        Me.txtManStrtAdd.Name = "txtManStrtAdd"
        Me.txtManStrtAdd.Size = New System.Drawing.Size(248, 20)
        Me.txtManStrtAdd.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 16)
        Me.Label11.TabIndex = 55
        Me.Label11.Text = "Company Name"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboState
        '
        Me.cboState.Items.AddRange(New Object() {"IL", "IN", "IA", "WI", "MI"})
        Me.cboState.Location = New System.Drawing.Point(304, 64)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(48, 21)
        Me.cboState.TabIndex = 4
        '
        'lblPriceList
        '
        Me.lblPriceList.Location = New System.Drawing.Point(368, 32)
        Me.lblPriceList.Name = "lblPriceList"
        Me.lblPriceList.Size = New System.Drawing.Size(88, 200)
        Me.lblPriceList.TabIndex = 79
        Me.lblPriceList.Text = "Label9"
        '
        'grpPricing
        '
        Me.grpPricing.Controls.Add(Me.txtPriceList)
        Me.grpPricing.Controls.Add(Me.lblPricingScheme)
        Me.grpPricing.Controls.Add(Me.btnRight)
        Me.grpPricing.Controls.Add(Me.btnLeft)
        Me.grpPricing.Location = New System.Drawing.Point(372, 8)
        Me.grpPricing.Name = "grpPricing"
        Me.grpPricing.Size = New System.Drawing.Size(128, 384)
        Me.grpPricing.TabIndex = 67
        Me.grpPricing.TabStop = False
        Me.grpPricing.Text = "Price List"
        '
        'txtPriceList
        '
        Me.txtPriceList.BackColor = System.Drawing.SystemColors.Control
        Me.txtPriceList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtPriceList.Location = New System.Drawing.Point(21, 24)
        Me.txtPriceList.Multiline = True
        Me.txtPriceList.Name = "txtPriceList"
        Me.txtPriceList.Size = New System.Drawing.Size(96, 280)
        Me.txtPriceList.TabIndex = 81
        Me.txtPriceList.TabStop = False
        '
        'lblPricingScheme
        '
        Me.lblPricingScheme.Location = New System.Drawing.Point(56, 336)
        Me.lblPricingScheme.Name = "lblPricingScheme"
        Me.lblPricingScheme.Size = New System.Drawing.Size(24, 24)
        Me.lblPricingScheme.TabIndex = 82
        Me.lblPricingScheme.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnRight
        '
        Me.btnRight.Location = New System.Drawing.Point(80, 344)
        Me.btnRight.Name = "btnRight"
        Me.btnRight.Size = New System.Drawing.Size(30, 24)
        Me.btnRight.TabIndex = 81
        Me.btnRight.Text = ">"
        '
        'btnLeft
        '
        Me.btnLeft.Location = New System.Drawing.Point(24, 344)
        Me.btnLeft.Name = "btnLeft"
        Me.btnLeft.Size = New System.Drawing.Size(30, 24)
        Me.btnLeft.TabIndex = 80
        Me.btnLeft.Text = "<"
        '
        'txtManName
        '
        Me.txtManName.Location = New System.Drawing.Point(104, 16)
        Me.txtManName.Name = "txtManName"
        Me.txtManName.Size = New System.Drawing.Size(248, 20)
        Me.txtManName.TabIndex = 1
        '
        'txtManEmail
        '
        Me.txtManEmail.HideSelection = False
        Me.txtManEmail.Location = New System.Drawing.Point(104, 208)
        Me.txtManEmail.Name = "txtManEmail"
        Me.txtManEmail.Size = New System.Drawing.Size(152, 20)
        Me.txtManEmail.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 208)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 71
        Me.Label1.Text = "Email Address"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "PNG Files|*.png|BMP Files|*.bmp|JPEG Files|*.jpg|TIFF Files|*.tif"
        Me.OpenFileDialog1.InitialDirectory = "C:\Program Files\JDR\BackflowTestingManager\logos"
        '
        'btnLogo
        '
        Me.btnLogo.Location = New System.Drawing.Point(24, 248)
        Me.btnLogo.Name = "btnLogo"
        Me.btnLogo.Size = New System.Drawing.Size(75, 23)
        Me.btnLogo.TabIndex = 72
        Me.btnLogo.Text = "Select Logo"
        Me.btnLogo.UseVisualStyleBackColor = True
        '
        'lblLogoPath
        '
        Me.lblLogoPath.AutoSize = True
        Me.lblLogoPath.Location = New System.Drawing.Point(117, 251)
        Me.lblLogoPath.Name = "lblLogoPath"
        Me.lblLogoPath.Size = New System.Drawing.Size(62, 13)
        Me.lblLogoPath.TabIndex = 73
        Me.lblLogoPath.Text = "Default.png"
        '
        'frmManAdd
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 498)
        Me.Controls.Add(Me.lblLogoPath)
        Me.Controls.Add(Me.btnLogo)
        Me.Controls.Add(Me.txtManEmail)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtManName)
        Me.Controls.Add(Me.grpPricing)
        Me.Controls.Add(Me.txtManOffice)
        Me.Controls.Add(Me.txtManFaxNo)
        Me.Controls.Add(Me.txtManPhone)
        Me.Controls.Add(Me.txtManContact)
        Me.Controls.Add(Me.txtManZipCode)
        Me.Controls.Add(Me.txtManCity)
        Me.Controls.Add(Me.txtManStrtAdd)
        Me.Controls.Add(Me.cboState)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblManNotes)
        Me.Controls.Add(Me.rtbManNotes)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.chkManActive)
        Me.Controls.Add(Me.Label11)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmManAdd"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Manager"
        Me.grpPricing.ResumeLayout(False)
        Me.grpPricing.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private PricingScheme As Integer = 2
    Private MaxSchemeNo As Integer
    Private PricingSchemeNo As Integer = 1
    Private logoPath As String = "C:\Program Files\JDR\BackflowTestingManager\logos\Default.png"
    Protected connectionString As String = ConfigurationManager.ConnectionStrings("bftm.My.MySettings.qbConnectionString1").ConnectionString


    Private Sub frmManAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the number of Schemes that have been entered
        SetMaxSchemeNo()
        ' Update the Price List according to the pricing scheme of the default New Manager.
        UpdatePriceList()
        ' Becausse a newly added manager is most likely going to be an active account set it as checked.
        chkManActive.Checked = True
        ' Mark the list as read only to avoid highlighting
        txtPriceList.ReadOnly = True
    End Sub

    Private Sub SetMaxSchemeNo()
        Dim cn As New SqlConnection(connectionString)
        ' Query string to retrieve the highest scheme number.
        Dim strQuery As String = "SELECT MAX(manSchemeNo) FROM PricingSchemes"
        Dim cmd As New SqlCommand(strQuery, cn)
        Try
            cn.Open() ' Open Connection
            MaxSchemeNo = CInt(cmd.ExecuteScalar())
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

    Private Function FormIsValid() As Boolean
        If NameIsValid() And StreetIsValid() And CityIsValid() And ZipIsValid() And StateIsValid() And EmailIsValid() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function NameIsValid() As Boolean
        Dim Booly As Boolean = True
        Dim SimilarManagers As ArrayList = SimilarManagerList()
        ErrorProvider1.SetError(txtManName, "")
        If SimilarManagers.Count > 0 Then
            Booly = False
            Dim LineItem As String
            Dim FullText As String
            FullText = "Manager Already Exists: " & ControlChars.NewLine
            For Each LineItem In SimilarManagers
                FullText &= LineItem
            Next
            ErrorProvider1.SetError(txtManName, FullText)
        End If
        If txtManName.Text.Trim = "" Or txtManName.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtManName, "Required Field: e.g. 'Best Buy'")
        End If
        Return Booly
    End Function

    Private Function StreetIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtManStrtAdd, "")
        If txtManStrtAdd.Text.Trim = "" Or txtManStrtAdd.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtManStrtAdd, "Required Field: e.g. '1234 North Damen Avenue'")
        End If

        Dim individualChr As Char

        For Each individualChr In txtManStrtAdd.Text
            If individualChr.CompareTo(Chr(46)) = 0 Then
                Booly = False
                ErrorProvider1.SetError(txtManStrtAdd, "Street addresses may not contain periods." _
                    & "No abbreviations!")
            End If
        Next
        Return Booly
    End Function

    Private Function CityIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtManCity, "")
        Dim CityRegex As New Regex("([A-Z][a-z]+[ ]?)+")
        If Not CityRegex.IsMatch(txtManCity.Text) Then
            Booly = False
            ErrorProvider1.SetError(txtManCity, "Please correct City name (Capitalization, No Numbers, etc.)")
        End If
        If txtManCity.Text.Trim = "" Or txtManCity.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtManCity, "Required Field: e.g. 'Orland Park'")
        End If
        Return Booly
    End Function

    Private Function ZipIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtManZipCode, "")
        Dim ZipRegex As New Regex("(^\d{5}$)|(^\d{5}-\d{4}$)")
        If Not ZipRegex.IsMatch(txtManZipCode.Text) Then
            Booly = False
            ErrorProvider1.SetError(txtManZipCode, "Zip code must be in the following forms: " _
                & ControlChars.NewLine & "of '60462' or '60462-1499'")
        End If
        If txtManZipCode.Text.Trim = "" Or txtManZipCode.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtManZipCode, "Required Field: e.g. '60462' or '60462-1499'")
        End If
        Return Booly
    End Function

    Private Function StateIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(cboState, "")
        Dim StateRegex As New Regex("^(?-i:A[LKSZRAP]|C[AOT]|D[EC]|F[LM]|G[AU]|HI|I[ADLN]|K[SY]|LA" _
            & "|M[ADEHINOPST]|N[CDEHJMVY]|O[HKR]|P[ARW]|RI|S[CD]|T[NX]|UT|" _
            & "V[AIT]|W[AIVY])$")
        If Not StateRegex.IsMatch(cboState.Text) Then
            Booly = False
            ErrorProvider1.SetError(cboState, "State abbreviation must be in the following form: 'IL'")
        End If
        If cboState.Text.Trim = "" Or cboState.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(cboState, "Required Field: e.g. 'IL'")
        End If
        Return Booly
    End Function

    Private Function PhoneIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtManPhone, "")
        If Not txtManPhone.Text.Trim = "" And Not txtManPhone.Text Is Nothing Then
            Dim PhoneRegex As New Regex("^\([1-9]\d{2}\)\s?\d{3}\-\d{4}( Ext.[0-9]+)?$")
            If Not PhoneRegex.IsMatch(txtManPhone.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtManPhone, "Phone numbers must be in the following form: " _
                & ControlChars.NewLine & "'(708)123-4567' or '(708)123-4567 Ext.456'")
            End If
        End If
        Return Booly
    End Function

    Private Function FaxIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtManFaxNo, "")
        If Not txtManFaxNo.Text.Trim = "" And Not txtManFaxNo.Text Is Nothing Then
            Dim FaxRegex As New Regex("^\([1-9]\d{2}\)\s?\d{3}\-\d{4}$")
            If Not FaxRegex.IsMatch(txtManFaxNo.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtManFaxNo, "Fax numbers must be in the following form: " _
                & ControlChars.NewLine & "'(708)123-4567'")
            End If
        End If
        Return Booly
    End Function

    Private Function EmailIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtManEmail, "")
        If Not txtManEmail.Text.Trim = "" And Not txtManEmail.Text Is Nothing Then
            Dim EmailRegex As New Regex("^([a-zA-Z0-9_\-])([a-zA-Z0-9_\-\.]*)@(\[((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}|((([a-zA-Z0-9\-]+)\.)+))([a-zA-Z]{2,}|(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\])$")
            If Not EmailRegex.IsMatch(txtManEmail.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtManEmail, "Email addresses must be in the following form: " _
                & ControlChars.NewLine & "'johnq@hotmail.com'")
            End If
        End If
        Return Booly
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If FormIsValid() Then

            AddManager()
        Else
            DialogResult = DialogResult.None
        End If
    End Sub

    Private Function SimilarManagerList() As ArrayList

        Dim resultList As New ArrayList
        Dim cn As New SqlConnection(connectionString)
        Dim strQuery As String = "SELECT manName, manStrtAdd FROM Managers WHERE (manName = '" _
            & Replace(txtManName.Text, "'", "''") & "' AND manStrtAdd = '" & Replace(txtManStrtAdd.Text, "'", "''") & "') AND manDeleted = 0"

        Dim cmd As New SqlCommand(strQuery, cn)
        Try
            cn.Open() ' Open Connection
            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            Do While dr.Read()
                resultList.Add(dr.Item("manName").ToString() & " at " & dr.Item("manStrtAdd").ToString() & ControlChars.NewLine)
            Loop
            dr.Close()
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

        Return resultList

    End Function

    Public Sub AddManager()
        ' Build the INSERT INTO string
        Dim strInsert As String = "INSERT INTO Managers " _
            & "(manName, manStrtAdd, manCity, manState, manZip, manSuite, " _
            & "manCntct, manPhone, manFax, manEmail, manNotes, manCurAcct, manSchemeNo, manLogoPath) " _
            & "VALUES ("

        strInsert &= "'" & Replace(Me.txtManName.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtManStrtAdd.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtManCity.Text, "'", "''") & "', "
        strInsert &= "'" & Me.cboState.Text & "', "
        strInsert &= "'" & Me.txtManZipCode.Text & "', "
        strInsert &= "'" & Replace(Me.txtManOffice.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtManContact.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtManPhone.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtManFaxNo.Text, "'", "''") & "', "
        strInsert &= "'" & Me.txtManEmail.Text & "', "
        strInsert &= "'" & Replace(Me.rtbManNotes.Text, "'", "''") & "', "
        strInsert &= CStr(CInt(Me.chkManActive.Checked)) & ", "
        strInsert &= CInt(Me.lblPricingScheme.Text) + 1 & ", "
        strInsert &= "'" & logoPath & "')"

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(strInsert, cn)
        Try
            cn.Open() ' Open the Connection
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

    Private Sub btnRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRight.Click
        btnLeft.Enabled = True
        If PricingScheme < MaxSchemeNo Then
            PricingScheme = PricingScheme + 1
            UpdatePriceList()
            If PricingScheme = MaxSchemeNo Then
                btnRight.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLeft.Click
        btnRight.Enabled = True
        If PricingScheme > 2 Then
            PricingScheme = PricingScheme - 1
            UpdatePriceList()
            If PricingScheme = 2 Then
                btnLeft.Enabled = False
            End If
        End If

    End Sub

    Sub UpdatePriceList()
        Dim BackflowByPass As Boolean = False
        Dim str As String = "Backflow Preventer" & ControlChars.CrLf & ControlChars.CrLf
        Dim cn As New SqlConnection(connectionString)
        Dim strSQL As String = "SELECT DISTINCT PricingSchemes.manSchemeNo, AssemblySizes.AssemSizeNo, " _
            & "AssemblySizes.AssemSize, Price.price, AssemblyUsage.assemUsagePriceNo FROM Price " _
            & "LEFT JOIN AssemblySizes ON AssemblySizes.AssemSizeNo = Price.assemSizeNo " _
            & "LEFT JOIN PricingSchemes ON PricingSchemes.manSchemeNo = Price.manSchemeNo " _
            & "LEFT JOIN AssemblyUsagePrice ON AssemblyUsagePrice.assemUsagePriceNo = Price.assemUsagePriceNo " _
            & "LEFT JOIN AssemblyUsage ON AssemblyUsage.assemUsagePriceNo = AssemblyUsagePrice.assemUsagePriceNo " _
            & "WHERE PricingSchemes.manSchemeNo = " & PricingScheme _
            & "ORDER BY  AssemblyUsage.assemUsagePriceNo, AssemblySizes.AssemSizeNo"

        Dim cmd As New SqlCommand(strSQL, cn)
        Try
            cn.Open() ' Open Connection
            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            Do While dr.Read()
                If CInt(dr.Item("assemUsagePriceNo")) = 2 And BackflowByPass = False Then
                    str &= ControlChars.CrLf & "Fire Line By-Pass" & ControlChars.CrLf & ControlChars.CrLf
                    BackflowByPass = True
                End If
                str &= dr.Item("AssemSize").ToString() & ControlChars.Tab
                str &= CDbl(dr("price")).ToString("c") & ControlChars.CrLf
            Loop
            dr.Close()
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

        lblPricingScheme.Text = (PricingScheme - 1).ToString()
        txtPriceList.Text = str

    End Sub

    Private Sub btnLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogo.Click
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            logoPath = OpenFileDialog1.FileName
            lblLogoPath.Text = logoPath.Substring(logoPath.LastIndexOf("\") + 1)
        End If
    End Sub
End Class
