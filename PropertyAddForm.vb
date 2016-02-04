Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class PropertyAddForm
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
    Friend WithEvents rtbPropNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtPropStore As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPropFaxNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPropPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtPropContact As System.Windows.Forms.TextBox
    Friend WithEvents txtPropZipCode As System.Windows.Forms.TextBox
    Friend WithEvents txtPropCity As System.Windows.Forms.TextBox
    Friend WithEvents txtPropStrtAdd As System.Windows.Forms.TextBox
    Friend WithEvents txtPropName As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboMunicipality As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboState As System.Windows.Forms.ComboBox
    Friend WithEvents btnManToProp As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PropertyAddForm))
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblManNotes = New System.Windows.Forms.Label
        Me.rtbPropNotes = New System.Windows.Forms.RichTextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtPropStore = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtPropFaxNo = New System.Windows.Forms.TextBox
        Me.txtPropPhone = New System.Windows.Forms.TextBox
        Me.txtPropContact = New System.Windows.Forms.TextBox
        Me.txtPropZipCode = New System.Windows.Forms.TextBox
        Me.txtPropCity = New System.Windows.Forms.TextBox
        Me.txtPropStrtAdd = New System.Windows.Forms.TextBox
        Me.txtPropName = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.cboMunicipality = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboState = New System.Windows.Forms.ComboBox
        Me.btnManToProp = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(308, 426)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 24)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(244, 426)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 24)
        Me.btnCancel.TabIndex = 78
        Me.btnCancel.Text = "Cancel"
        '
        'lblManNotes
        '
        Me.lblManNotes.Location = New System.Drawing.Point(12, 258)
        Me.lblManNotes.Name = "lblManNotes"
        Me.lblManNotes.Size = New System.Drawing.Size(144, 16)
        Me.lblManNotes.TabIndex = 89
        Me.lblManNotes.Text = "Property Specific Notes"
        '
        'rtbPropNotes
        '
        Me.rtbPropNotes.Location = New System.Drawing.Point(12, 282)
        Me.rtbPropNotes.Name = "rtbPropNotes"
        Me.rtbPropNotes.Size = New System.Drawing.Size(352, 128)
        Me.rtbPropNotes.TabIndex = 10
        Me.rtbPropNotes.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(32, 120)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 16)
        Me.Label10.TabIndex = 88
        Me.Label10.Text = "Store No."
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPropStore
        '
        Me.txtPropStore.HideSelection = False
        Me.txtPropStore.Location = New System.Drawing.Point(96, 120)
        Me.txtPropStore.Name = "txtPropStore"
        Me.txtPropStore.Size = New System.Drawing.Size(152, 20)
        Me.txtPropStore.TabIndex = 5
        Me.txtPropStore.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(264, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 24)
        Me.Label8.TabIndex = 87
        Me.Label8.Text = "State"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(0, 192)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 86
        Me.Label7.Text = "Fax No."
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(0, 168)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 85
        Me.Label6.Text = "Phone No."
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(0, 144)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 84
        Me.Label5.Text = "Contact Person"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(0, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 83
        Me.Label4.Text = "Postal Code"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(0, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 82
        Me.Label3.Text = "City"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 81
        Me.Label2.Text = "Street Address"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPropFaxNo
        '
        Me.txtPropFaxNo.HideSelection = False
        Me.txtPropFaxNo.Location = New System.Drawing.Point(96, 192)
        Me.txtPropFaxNo.Name = "txtPropFaxNo"
        Me.txtPropFaxNo.Size = New System.Drawing.Size(152, 20)
        Me.txtPropFaxNo.TabIndex = 8
        Me.txtPropFaxNo.Text = ""
        '
        'txtPropPhone
        '
        Me.txtPropPhone.HideSelection = False
        Me.txtPropPhone.Location = New System.Drawing.Point(96, 168)
        Me.txtPropPhone.Name = "txtPropPhone"
        Me.txtPropPhone.Size = New System.Drawing.Size(152, 20)
        Me.txtPropPhone.TabIndex = 7
        Me.txtPropPhone.Text = ""
        '
        'txtPropContact
        '
        Me.txtPropContact.HideSelection = False
        Me.txtPropContact.Location = New System.Drawing.Point(96, 144)
        Me.txtPropContact.Name = "txtPropContact"
        Me.txtPropContact.Size = New System.Drawing.Size(152, 20)
        Me.txtPropContact.TabIndex = 6
        Me.txtPropContact.Text = ""
        '
        'txtPropZipCode
        '
        Me.txtPropZipCode.Location = New System.Drawing.Point(96, 96)
        Me.txtPropZipCode.Name = "txtPropZipCode"
        Me.txtPropZipCode.Size = New System.Drawing.Size(152, 20)
        Me.txtPropZipCode.TabIndex = 4
        Me.txtPropZipCode.Text = ""
        '
        'txtPropCity
        '
        Me.txtPropCity.Location = New System.Drawing.Point(96, 72)
        Me.txtPropCity.Name = "txtPropCity"
        Me.txtPropCity.Size = New System.Drawing.Size(152, 20)
        Me.txtPropCity.TabIndex = 2
        Me.txtPropCity.Text = ""
        '
        'txtPropStrtAdd
        '
        Me.txtPropStrtAdd.Location = New System.Drawing.Point(96, 48)
        Me.txtPropStrtAdd.Name = "txtPropStrtAdd"
        Me.txtPropStrtAdd.Size = New System.Drawing.Size(256, 20)
        Me.txtPropStrtAdd.TabIndex = 1
        Me.txtPropStrtAdd.Text = ""
        '
        'txtPropName
        '
        Me.txtPropName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPropName.Location = New System.Drawing.Point(96, 24)
        Me.txtPropName.Name = "txtPropName"
        Me.txtPropName.Size = New System.Drawing.Size(256, 20)
        Me.txtPropName.TabIndex = 0
        Me.txtPropName.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(0, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 16)
        Me.Label11.TabIndex = 80
        Me.Label11.Text = "Property Name"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboMunicipality
        '
        Me.cboMunicipality.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMunicipality.Location = New System.Drawing.Point(96, 216)
        Me.cboMunicipality.Name = "cboMunicipality"
        Me.cboMunicipality.Size = New System.Drawing.Size(152, 21)
        Me.cboMunicipality.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 216)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 90
        Me.Label1.Text = "Municipality"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboState
        '
        Me.cboState.Items.AddRange(New Object() {"IL", "IN", "IA", "WI", "MI"})
        Me.cboState.Location = New System.Drawing.Point(304, 72)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(48, 21)
        Me.cboState.TabIndex = 3
        '
        'btnManToProp
        '
        Me.btnManToProp.Location = New System.Drawing.Point(12, 426)
        Me.btnManToProp.Name = "btnManToProp"
        Me.btnManToProp.Size = New System.Drawing.Size(154, 24)
        Me.btnManToProp.TabIndex = 116
        Me.btnManToProp.Text = "Manager Info to Property"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmPropAdd
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 461)
        Me.Controls.Add(Me.btnManToProp)
        Me.Controls.Add(Me.cboState)
        Me.Controls.Add(Me.cboMunicipality)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblManNotes)
        Me.Controls.Add(Me.rtbPropNotes)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txtPropStore)
        Me.Controls.Add(Me.txtPropFaxNo)
        Me.Controls.Add(Me.txtPropPhone)
        Me.Controls.Add(Me.txtPropContact)
        Me.Controls.Add(Me.txtPropZipCode)
        Me.Controls.Add(Me.txtPropCity)
        Me.Controls.Add(Me.txtPropStrtAdd)
        Me.Controls.Add(Me.txtPropName)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label11)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPropAdd"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Property"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Shared manNumber As Integer ' Set through MainForm
    Protected connectionString As String = ConfigurationManager.ConnectionStrings("bftm.My.MySettings.qbConnectionString1").ConnectionString


    ' Handle the Form load event.
    Private Sub frmPropAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FillComboBoxes()
    End Sub

    Private Sub FillComboBoxes()
        cboMunicipality.ValueMember = "munNo"
        cboMunicipality.DisplayMember = "munName"
        cboMunicipality.DataSource = MainForm.ds.Tables("Municip")
        cboMunicipality.SelectedValue = 294
    End Sub

    Public Sub AddProperty()
        ' Build the INSERT INTO string
        Dim strInsert As String = "INSERT INTO Properties " _
            & "(munNo, manNo, propName, propStrt, propCity, propState, propZip, storeNo, " _
            & "propCon, propPhone, propFax, propAdded, propNotes) " _
            & "VALUES ("

        strInsert &= cboMunicipality.SelectedValue.ToString() & ", "
        strInsert &= manNumber & ", "
        strInsert &= "'" & Replace(txtPropName.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(txtPropStrtAdd.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(txtPropCity.Text, "'", "''") & "', "
        strInsert &= "'" & Me.cboState.Text & "', "
        strInsert &= "'" & Me.txtPropZipCode.Text & "', "
        strInsert &= "'" & Replace(txtPropStore.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(txtPropContact.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtPropPhone.Text, "'", "''") & "', "
        strInsert &= "'" & Replace(Me.txtPropFaxNo.Text, "'", "''") & "', "
        strInsert &= "'" & Date.Today.ToShortDateString & "', "
        strInsert &= "'" & Replace(rtbPropNotes.Text, "'", "''") & "')"

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

    Private Sub btnManToProp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManToProp.Click
        GetManData()
    End Sub

    Private Sub GetManData()

        Dim PropName As String
        Dim Street As String
        Dim City As String
        Dim State As String
        Dim Zip As String
        Dim Store As String
        Dim Contact As String
        Dim Phone As String
        Dim Fax As String

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(("SELECT * FROM Managers WHERE manNo = " & manNumber), cn)
        ' Run the query; get the DataReader object.

        Try
            cn.Open()
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            Do While dr.Read()
                PropName = dr.Item("manName").ToString() + ""
                Street = dr.Item("manStrtAdd").ToString() + ""
                City = dr.Item("manCity").ToString() + ""
                State = dr.Item("manState").ToString() + ""
                Zip = dr.Item("manZip").ToString() + ""
                Store = dr.Item("manSuite").ToString() + ""
                Contact = dr.Item("manCntct").ToString() + ""
                Phone = dr.Item("manPhone").ToString() + ""
                Fax = dr.Item("manFax").ToString() + ""
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

        Me.txtPropName.Text = PropName
        Me.txtPropStrtAdd.Text = Street
        Me.txtPropCity.Text = City
        Me.cboState.Text = State
        Me.txtPropZipCode.Text = Zip
        Me.txtPropStore.Text = Store
        Me.txtPropContact.Text = Contact
        Me.txtPropPhone.Text = Phone
        Me.txtPropFaxNo.Text = Fax

        Me.cboMunicipality.SelectedValue = 294

    End Sub 'GetManData

    Private Function FormIsValid() As Boolean
        If NameIsValid() And StreetIsValid() And CityIsValid() And ZipIsValid() And StateIsValid() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function NameIsValid() As Boolean
        Dim Booly As Boolean = True
        If 1 = 2 Then
            Dim SimilarProperties As ArrayList = SimilarPropertyList()
            ErrorProvider1.SetError(txtPropName, "")
            If SimilarProperties.Count > 0 Then
                Booly = False
                Dim LineItem As String
                Dim FullText As String
                FullText = "Property Already Exists: " & ControlChars.NewLine
                For Each LineItem In SimilarProperties
                    FullText &= LineItem
                Next
                ErrorProvider1.SetError(txtPropName, FullText)
            End If
        End If
        If txtPropName.Text.Trim = "" Or txtPropName.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtPropName, "Required Field: e.g. 'Best Buy'")
        End If
        Return Booly
    End Function

    Private Function StreetIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtPropStrtAdd, "")
        If txtPropStrtAdd.Text.Trim = "" Or txtPropStrtAdd.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtPropStrtAdd, "Required Field: e.g. '1234 North Damen Avenue'")
        End If

        Dim individualChr As Char

        For Each individualChr In txtPropStrtAdd.Text
            If individualChr.CompareTo(Chr(46)) = 0 Then
                Booly = False
                ErrorProvider1.SetError(txtPropStrtAdd, "Street addresses may not contain periods." _
                    & "No abbreviations!")
            End If
        Next
        Return Booly
    End Function

    Private Function CityIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtPropCity, "")
        Dim CityRegex As New Regex("([A-Z][a-z]+[ ]?)+")
        If Not CityRegex.IsMatch(txtPropCity.Text) Then
            Booly = False
            ErrorProvider1.SetError(txtPropCity, "Please correct City name (Capitalization, No Numbers, etc.)")
        End If
        If txtPropCity.Text.Trim = "" Or txtPropCity.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtPropCity, "Required Field: e.g. 'Orland Park'")
        End If
        Return Booly
    End Function

    Private Function ZipIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtPropZipCode, "")
        Dim ZipRegex As New Regex("(^\d{5}$)|(^\d{5}-\d{4}$)")
        If Not ZipRegex.IsMatch(txtPropZipCode.Text) Then
            Booly = False
            ErrorProvider1.SetError(txtPropZipCode, "Zip code must be in the following forms: " _
                & ControlChars.NewLine & "of '60462' or '60462-1499'")
        End If
        If txtPropZipCode.Text.Trim = "" Or txtPropCity.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtPropZipCode, "Required Field: e.g. '60462' or '60462-1499'")
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
        ErrorProvider1.SetError(txtPropPhone, "")
        If Not txtPropPhone.Text.Trim = "" And Not txtPropPhone.Text Is Nothing Then
            Dim PhoneRegex As New Regex("^\([1-9]\d{2}\)\s?\d{3}\-\d{4}( Ext.[0-9]+)?$")
            If Not PhoneRegex.IsMatch(txtPropPhone.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtPropPhone, "Phone numbers must be in the following form: " _
                & ControlChars.NewLine & "'(708)123-4567' or '(708)123-4567 Ext.456'")
            End If
        End If
        Return Booly
    End Function

    Private Function FaxIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtPropFaxNo, "")
        If Not txtPropFaxNo.Text.Trim = "" And Not txtPropFaxNo.Text Is Nothing Then
            Dim FaxRegex As New Regex("^\([1-9]\d{2}\)\s?\d{3}\-\d{4}$")
            If Not FaxRegex.IsMatch(txtPropFaxNo.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtPropFaxNo, "Fax numbers must be in the following form: " _
                & ControlChars.NewLine & "'(708)123-4567'")
            End If
        End If
        Return Booly
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If FormIsValid() Then
            AddProperty()
        Else
            DialogResult = DialogResult.None
        End If
    End Sub

    Private Function SimilarPropertyList() As ArrayList

        Dim resultList As New ArrayList
        Dim cn As New SqlConnection(connectionString)
        Dim strQuery As String = "SELECT propName, propStrt FROM Properties WHERE (propName = '" _
            & Replace(txtPropName.Text, "'", "''") & "' AND propStrt = '" & Replace(txtPropStrtAdd.Text, "'", "''") & "') AND propDeleted = 0"

        Dim cmd As New SqlCommand(strQuery, cn)
        Try
            cn.Open() ' Open Connection
            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            Do While dr.Read()
                resultList.Add(dr.Item("propName").ToString() & " at " & dr.Item("propStrt").ToString() & ControlChars.NewLine)
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

End Class
