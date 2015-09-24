Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class MunicipalityForm
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
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtDept As System.Windows.Forms.TextBox
    Friend WithEvents txtStreet As System.Windows.Forms.TextBox
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtZipCode As System.Windows.Forms.TextBox
    Friend WithEvents cboPreface As System.Windows.Forms.ComboBox
    Friend WithEvents cboState As System.Windows.Forms.ComboBox
    Friend WithEvents txtContact As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtFax As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents lvMunicipalities As System.Windows.Forms.ListView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MunicipalityForm))
        Me.cboPreface = New System.Windows.Forms.ComboBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtDept = New System.Windows.Forms.TextBox
        Me.txtStreet = New System.Windows.Forms.TextBox
        Me.txtCity = New System.Windows.Forms.TextBox
        Me.cboState = New System.Windows.Forms.ComboBox
        Me.txtZipCode = New System.Windows.Forms.TextBox
        Me.txtContact = New System.Windows.Forms.TextBox
        Me.txtPhone = New System.Windows.Forms.TextBox
        Me.txtFax = New System.Windows.Forms.TextBox
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.lvMunicipalities = New System.Windows.Forms.ListView
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider
        Me.SuspendLayout()
        '
        'cboPreface
        '
        Me.cboPreface.Items.AddRange(New Object() {"Village Of", "City Of"})
        Me.cboPreface.Location = New System.Drawing.Point(96, 150)
        Me.cboPreface.Name = "cboPreface"
        Me.cboPreface.Size = New System.Drawing.Size(152, 21)
        Me.cboPreface.TabIndex = 1
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(96, 174)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(280, 20)
        Me.txtName.TabIndex = 2
        Me.txtName.Text = ""
        '
        'txtDept
        '
        Me.txtDept.Location = New System.Drawing.Point(96, 198)
        Me.txtDept.Name = "txtDept"
        Me.txtDept.Size = New System.Drawing.Size(280, 20)
        Me.txtDept.TabIndex = 3
        Me.txtDept.Text = ""
        '
        'txtStreet
        '
        Me.txtStreet.Location = New System.Drawing.Point(96, 222)
        Me.txtStreet.Name = "txtStreet"
        Me.txtStreet.Size = New System.Drawing.Size(280, 20)
        Me.txtStreet.TabIndex = 4
        Me.txtStreet.Text = ""
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(96, 246)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(152, 20)
        Me.txtCity.TabIndex = 5
        Me.txtCity.Text = ""
        '
        'cboState
        '
        Me.cboState.Items.AddRange(New Object() {"IL", "IN", "IA", "MI"})
        Me.cboState.Location = New System.Drawing.Point(320, 246)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(56, 21)
        Me.cboState.TabIndex = 6
        '
        'txtZipCode
        '
        Me.txtZipCode.Location = New System.Drawing.Point(96, 270)
        Me.txtZipCode.Name = "txtZipCode"
        Me.txtZipCode.Size = New System.Drawing.Size(152, 20)
        Me.txtZipCode.TabIndex = 7
        Me.txtZipCode.Text = ""
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(96, 294)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Size = New System.Drawing.Size(280, 20)
        Me.txtContact.TabIndex = 8
        Me.txtContact.Text = ""
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(96, 318)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(152, 20)
        Me.txtPhone.TabIndex = 9
        Me.txtPhone.Text = ""
        '
        'txtFax
        '
        Me.txtFax.Location = New System.Drawing.Point(96, 342)
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(152, 20)
        Me.txtFax.TabIndex = 10
        Me.txtFax.Text = ""
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(96, 366)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(152, 20)
        Me.txtEmail.TabIndex = 11
        Me.txtEmail.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(280, 254)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 16)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "State"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(400, 190)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(72, 32)
        Me.btnAdd.TabIndex = 23
        Me.btnAdd.Text = "New"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(400, 238)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(72, 32)
        Me.btnSave.TabIndex = 24
        Me.btnSave.Text = "Save"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(400, 286)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(72, 32)
        Me.btnDelete.TabIndex = 25
        Me.btnDelete.Text = "Delete"
        '
        'lvMunicipalities
        '
        Me.lvMunicipalities.FullRowSelect = True
        Me.lvMunicipalities.Location = New System.Drawing.Point(8, 8)
        Me.lvMunicipalities.MultiSelect = False
        Me.lvMunicipalities.Name = "lvMunicipalities"
        Me.lvMunicipalities.Size = New System.Drawing.Size(368, 128)
        Me.lvMunicipalities.TabIndex = 26
        Me.lvMunicipalities.View = System.Windows.Forms.View.Details
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 222)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "Street Address"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 198)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Department"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 174)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "Name"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 150)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Preface"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 246)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 16)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "City"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 294)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(80, 16)
        Me.Label13.TabIndex = 33
        Me.Label13.Text = "Contact"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 318)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(80, 16)
        Me.Label14.TabIndex = 34
        Me.Label14.Text = "Phone"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 342)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 16)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "Fax"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(8, 366)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(80, 16)
        Me.Label16.TabIndex = 36
        Me.Label16.Text = "Email"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 270)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(80, 16)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Zip Code"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmMun
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 397)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.lvMunicipalities)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtFax)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.txtContact)
        Me.Controls.Add(Me.txtZipCode)
        Me.Controls.Add(Me.cboState)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtStreet)
        Me.Controls.Add(Me.txtDept)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.cboPreface)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(512, 424)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(440, 424)
        Me.Name = "frmMun"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Municipalities"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private municipalityNo As Integer
    Dim mode As String = "update"

    Protected connectionString As String = Connection.SQL_CONNECTION_STRING
    ' Handle the Form load event.
    Private Sub frmMun_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PopulateMunicipalityList()
        lvMunicipalities.Items(0).Selected = True
    End Sub

    Private Sub PopulateMunicipalityList()
        Dim cn As New SqlConnection(connectionString)
        Dim sql As String = "SELECT munName, munNo FROM Municipalities WHERE munDeleted = 0 ORDER BY munName"
        Dim cmd As New SqlCommand(sql, cn)
        ' lvMunicipalities.Items.Clear()
        lvMunicipalities.Clear()

        ' Set to details view
        Dim columnheader As ColumnHeader    ' Used for creating column headers.
        Dim listviewitem As ListViewItem    ' Used for creating ListView items.

        ' Make sure that the view is set to show details.
        lvMunicipalities.View = View.Details

        Try
            cn.Open()
            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            Do While dr.Read()

                ' Create some ListView items consisting of first and last names.
                listviewitem = New myListViewItem((dr.Item("munName").ToString() & ""), "Municipality", CInt(dr.Item("munNo")), 1)
                Me.lvMunicipalities.Items.Add(listviewitem)

            Loop
            ' lstMunicipalities.SetSelected(0, True)
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
        ' Create some column headers for the data.
        columnheader = New ColumnHeader
        columnheader.Text = "Municipality"
        Me.lvMunicipalities.Columns.Add(columnheader)

        ' Loop through and size each column header to fit the column header text.
        For Each columnheader In Me.lvMunicipalities.Columns
            columnheader.Width = -2
        Next

    End Sub

    Private Sub lvMunicipalities_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvMunicipalities.SelectedIndexChanged
        ClearErrors()
        If Me.lvMunicipalities.SelectedIndices.Count <= 0 Then
            Return
        End If
        Dim selNdx As Integer = Me.lvMunicipalities.SelectedIndices(0)
        If CInt(selNdx) >= 0 Then
            municipalityNo = CType(lvMunicipalities.Items(selNdx), myListViewItem).intID
            PopulateFields(municipalityNo)
            mode = "update"
        End If
        btnDelete.Enabled = True
    End Sub

    Private Sub ClearErrors()
        ErrorProvider1.SetError(txtName, "")
        ErrorProvider1.SetError(txtStreet, "")
        ErrorProvider1.SetError(txtCity, "")
        ErrorProvider1.SetError(txtZipCode, "")
        ErrorProvider1.SetError(cboState, "")
        ErrorProvider1.SetError(txtPhone, "")
        ErrorProvider1.SetError(txtFax, "")
        ErrorProvider1.SetError(txtEmail, "")
    End Sub

    Private Sub PopulateFields(ByVal munNumber As Integer)
        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(("SELECT * FROM Municipalities WHERE munNo = " & munNumber), cn)
        Try
            cn.Open() ' Open Connection
            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            If dr.Read() Then
                txtName.Text = dr.Item("munName").ToString() & ""
                txtDept.Text = dr.Item("munDept").ToString() & ""
                txtStreet.Text = dr.Item("munStrtAddr").ToString() & ""
                txtCity.Text = dr.Item("munCity").ToString() & ""
                txtZipCode.Text = dr.Item("munZip").ToString() & ""
                txtContact.Text = dr.Item("munCntct").ToString() & ""
                txtPhone.Text = dr.Item("munPhone").ToString() & ""
                txtFax.Text = dr.Item("munFax").ToString() & ""
                txtEmail.Text = dr.Item("munEmail").ToString() & ""
                cboPreface.Text = dr.Item("munPref").ToString() & ""
                cboState.Text = dr.Item("munState").ToString() & ""

            End If
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

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ClearForm()
        mode = "add"
        btnDelete.Enabled = False
    End Sub

    Private Sub ClearForm()
        txtName.Text = ""
        txtDept.Text = ""
        txtStreet.Text = ""
        txtCity.Text = ""
        txtZipCode.Text = ""
        txtContact.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        cboPreface.Text = ""
        cboState.Text = ""
    End Sub

    Private Sub AddOrUpdate()
        If mode = "add" Then
            AddMunicipality()
        Else
            UpdateMunicipality(municipalityNo)
        End If
    End Sub

    Private Sub AddMunicipality()
        ' Build the INSERT INTO string
        Dim strInsert As String = "INSERT INTO Municipalities " _
            & "(munName, munCity, munFax, munPhone, munState, munStrtAddr, munCntct, " _
            & "munPref, munZip, munEmail, munDept) " _
            & "VALUES ("

        strInsert &= "'" & txtName.Text & "', "
        strInsert &= "'" & txtCity.Text & "', "
        strInsert &= "'" & txtFax.Text & "', "
        strInsert &= "'" & txtPhone.Text & "', "
        strInsert &= "'" & cboState.Text & "', "
        strInsert &= "'" & txtStreet.Text & "', "
        strInsert &= "'" & txtContact.Text & "', "
        strInsert &= "'" & cboPreface.Text & "', "
        strInsert &= "'" & txtZipCode.Text & "', "
        strInsert &= "'" & txtEmail.Text & "', "
        strInsert &= "'" & txtDept.Text & "')"

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
        End Try
        ClearForm()
        RefillMunicipalityDT()
    End Sub

    Private Sub UpdateMunicipality(ByVal MunicipalityNumber As Integer)
        ' Build the Update String
        Dim strUpdate As String = "UPDATE Municipalities SET "
        strUpdate &= "munName = '" & txtName.Text & "', "
        strUpdate &= "munCity = '" & txtCity.Text & "', "
        strUpdate &= "munFax = '" & txtFax.Text & "', "
        strUpdate &= "munPhone = '" & txtPhone.Text & "', "
        strUpdate &= "munState = '" & cboState.Text & "', "
        strUpdate &= "munStrtAddr = '" & txtStreet.Text & "', "
        strUpdate &= "munCntct = '" & txtContact.Text & "', "
        strUpdate &= "munPref = '" & cboPreface.Text & "', "
        strUpdate &= "munZip = '" & txtZipCode.Text & "', "
        strUpdate &= "munEmail = '" & txtEmail.Text & "', "
        strUpdate &= "munDept = '" & txtDept.Text & "' "
        strUpdate &= "WHERE munNo = " & MunicipalityNumber

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
        End Try
        ClearForm()
        RefillMunicipalityDT()

    End Sub

    Private Sub RefillMunicipalityDT()
        Dim cn As New SqlConnection(connectionString)
        Dim sqlStatement As String = "SELECT munNo, munName FROM Municipalities ORDER BY munName"
        Dim cmd As New SqlCommand(sqlStatement, cn)
        Dim sda As New SqlDataAdapter(cmd)
        MainForm.ds.Tables("Municip").Clear()

        sda.Fill(MainForm.ds, "Municip") ' Fill the Municipalities datatable
    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If FormIsValid() Then
            AddOrUpdate()
            PopulateMunicipalityList()
            lvMunicipalities.Items(0).Selected = True
        Else
            DialogResult = DialogResult.None
        End If
    End Sub

    Private Sub frmMun_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        DialogResult() = DialogResult.OK
    End Sub

    Private Function FormIsValid() As Boolean
        If NameIsValid() And StreetIsValid() And CityIsValid() And ZipIsValid() And StateIsValid() And PhoneIsValid() And FaxIsValid() And EmailIsValid() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function NameIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtName, "")
        If txtName.Text.Trim = "" Or txtName.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtName, "Name is a required field")
        End If
        Return Booly
    End Function

    Private Function StreetIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtStreet, "")
        If txtStreet.Text.Trim = "" Or txtStreet.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtStreet, "Required Field: e.g. '1234 North Damen Avenue'")
        End If

        Dim individualChr As Char

        For Each individualChr In txtStreet.Text
            If individualChr.CompareTo(Chr(46)) = 0 Then
                Booly = False
                ErrorProvider1.SetError(txtStreet, "Street addresses may not contain periods." _
                    & "No abbreviations!")
            End If
        Next
        Return Booly
    End Function

    Private Function CityIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtCity, "")
        Dim CityRegex As New Regex("([A-Z][a-z]+[ ]?)+")
        If Not CityRegex.IsMatch(txtCity.Text) Then
            Booly = False
            ErrorProvider1.SetError(txtCity, "Please correct City name (Capitalization, No Numbers, etc.)")
        End If
        If txtCity.Text.Trim = "" Or txtCity.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtCity, "Required Field: e.g. 'Orland Park'")
        End If
        Return Booly
    End Function

    Private Function ZipIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtZipCode, "")
        Dim ZipRegex As New Regex("(^\d{5}$)|(^\d{5}-\d{4}$)")
        If Not ZipRegex.IsMatch(txtZipCode.Text) Then
            Booly = False
            ErrorProvider1.SetError(txtZipCode, "Zip code must be in the following forms: " _
                & ControlChars.NewLine & "of '60462' or '60462-1499'")
        End If
        If txtZipCode.Text.Trim = "" Or txtZipCode.Text Is Nothing Then
            Booly = False
            ErrorProvider1.SetError(txtZipCode, "Required Field: e.g. '60462' or '60462-1499'")
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
        ErrorProvider1.SetError(txtPhone, "")
        If Not txtPhone.Text.Trim = "" And Not txtPhone.Text Is Nothing Then
            Dim PhoneRegex As New Regex("^\([1-9]\d{2}\)\s?\d{3}\-\d{4}( Ext.[0-9]+)?$")
            If Not PhoneRegex.IsMatch(txtPhone.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtPhone, "Phone numbers must be in the following form: " _
                & ControlChars.NewLine & "'(708)123-4567' or '(708)123-4567 Ext.456'")
            End If
        End If
        Return Booly
    End Function

    Private Function FaxIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtFax, "")
        If Not txtFax.Text.Trim = "" And Not txtFax.Text Is Nothing Then
            Dim FaxRegex As New Regex("^\([1-9]\d{2}\)\s?\d{3}\-\d{4}$")
            If Not FaxRegex.IsMatch(txtFax.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtFax, "Fax numbers must be in the following form: " _
                & ControlChars.NewLine & "'(708)123-4567'")
            End If
        End If
        Return Booly
    End Function

    Private Function EmailIsValid() As Boolean
        Dim Booly As Boolean = True
        ErrorProvider1.SetError(txtEmail, "")
        If Not txtEmail.Text.Trim = "" And Not txtEmail.Text Is Nothing Then
            Dim EmailRegex As New Regex("^([a-zA-Z0-9_\-])([a-zA-Z0-9_\-\.]*)@(\[((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}|((([a-zA-Z0-9\-]+)\.)+))([a-zA-Z]{2,}|(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\])$")
            If Not EmailRegex.IsMatch(txtEmail.Text) Then
                Booly = False
                ErrorProvider1.SetError(txtEmail, "Email addresses must be in the following form: " _
                & ControlChars.NewLine & "'johnq@hotmail.com'")
            End If
        End If
        Return Booly
    End Function

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If MessageBox.Show("Are you sure you want to delete the Municipality?", "Confirm Delete", MessageBoxButtons.YesNo, _
                    MessageBoxIcon.Question) = DialogResult.Yes Then

            Dim selNdx As Integer = Me.lvMunicipalities.SelectedIndices(0)
            If selNdx >= 0 Then
                Dim municipality As myListViewItem = CType(lvMunicipalities.Items(selNdx), myListViewItem)
                Dim strUpdate As String = "UPDATE Properties SET munNo = 294 WHERE " _
                    & "munNo = " & municipality.intID
                Dim cn As New SqlConnection(connectionString)
                Dim cmd As New SqlCommand(strUpdate, cn)
                Try
                    cn.Open() ' Open the Connection
                    If Not municipality.intID = 294 Then
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "DELETE FROM Municipalities WHERE munNo = " & municipality.intID
                        cmd.ExecuteNonQuery()
                    End If
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

                municipality.Remove()
                ClearErrors()
                ClearForm()
                mode = "add"
                btnDelete.Enabled = False
            End If
        End If

    End Sub

End Class
