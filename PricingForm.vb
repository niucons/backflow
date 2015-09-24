Imports System.Data
Imports System.Data.SqlClient
Public Class PricingForm
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
    Friend WithEvents btnLeft As System.Windows.Forms.Button
    Friend WithEvents btnRight As System.Windows.Forms.Button
    Friend WithEvents grdScheme As System.Windows.Forms.DataGrid
    Friend WithEvents btnSave As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PricingForm))
        Me.grdScheme = New System.Windows.Forms.DataGrid
        Me.btnLeft = New System.Windows.Forms.Button
        Me.btnRight = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        CType(Me.grdScheme, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdScheme
        '
        Me.grdScheme.BackgroundColor = System.Drawing.SystemColors.AppWorkspace
        Me.grdScheme.DataMember = ""
        Me.grdScheme.GridLineColor = System.Drawing.SystemColors.Control
        Me.grdScheme.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdScheme.Location = New System.Drawing.Point(8, 8)
        Me.grdScheme.Name = "grdScheme"
        Me.grdScheme.Size = New System.Drawing.Size(176, 288)
        Me.grdScheme.TabIndex = 0
        '
        'btnLeft
        '
        Me.btnLeft.Location = New System.Drawing.Point(192, 256)
        Me.btnLeft.Name = "btnLeft"
        Me.btnLeft.Size = New System.Drawing.Size(30, 24)
        Me.btnLeft.TabIndex = 2
        Me.btnLeft.Text = "<"
        '
        'btnRight
        '
        Me.btnRight.Location = New System.Drawing.Point(232, 256)
        Me.btnRight.Name = "btnRight"
        Me.btnRight.Size = New System.Drawing.Size(30, 24)
        Me.btnRight.TabIndex = 3
        Me.btnRight.Text = ">"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(192, 208)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(64, 32)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        '
        'frmPricing
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(264, 301)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnRight)
        Me.Controls.Add(Me.btnLeft)
        Me.Controls.Add(Me.grdScheme)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPricing"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Pricing Schemes"
        CType(Me.grdScheme, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private PricingScheme As Integer = 1
    Private MaxSchemeNo As Integer
    Private dtSchemes As DataTable
    Private dvSchemes As DataView
    Private sda As SqlDataAdapter

    Protected connectionString As String = Connection.SQL_CONNECTION_STRING

    ' Handle the Form load event.
    Private Sub frmPricing_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FillPricingDataTable()
        SetMaxSchemeNo()
        btnLeft.Enabled = False
    End Sub

    Private Sub SetMaxSchemeNo()
        Dim cn As New SqlConnection(connectionString)
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

    Private Sub FillPricingDataTable()
        Dim cn As New SqlConnection(connectionString)
        Dim sqlStatement As String = "SELECT DISTINCT Price.priceNo, PricingSchemes.manSchemeNo, AssemblySizes.AssemSizeNo, " _
            & "AssemblySizes.AssemSize, Price.price, AssemblyUsage.assemUsagePriceNo FROM Price " _
            & "LEFT JOIN AssemblySizes ON AssemblySizes.AssemSizeNo = Price.assemSizeNo " _
            & "LEFT JOIN PricingSchemes ON PricingSchemes.manSchemeNo = Price.manSchemeNo " _
            & "LEFT JOIN AssemblyUsagePrice ON AssemblyUsagePrice.assemUsagePriceNo = Price.assemUsagePriceNo " _
            & "LEFT JOIN AssemblyUsage ON AssemblyUsage.assemUsagePriceNo = AssemblyUsagePrice.assemUsagePriceNo " _
            & "ORDER BY  AssemblyUsage.assemUsagePriceNo, AssemblySizes.AssemSizeNo"
        Dim cmd As New SqlCommand(sqlStatement, cn)
        sda = New SqlDataAdapter(cmd)


        If Not MainForm.ds.Tables("PricingSchemes") Is Nothing Then
            MainForm.ds.Tables("PricingSchemes").Clear()
        End If

        sda.Fill(MainForm.ds, "PricingSchemes")

        dtSchemes = MainForm.ds.Tables("PricingSchemes")

        dvSchemes = dtSchemes.DefaultView
        dvSchemes.AllowDelete = False
        dvSchemes.AllowEdit = True
        dvSchemes.AllowNew = False

        BindSchemesGrid()
    End Sub

    Sub BindSchemesGrid()

        ' Filter the OrderDetails data based on the currently selected OrderID.
        dvSchemes.RowFilter = "manSchemeNo = " & PricingScheme

        With grdScheme
            .CaptionText = "Scheme No. " & PricingScheme
            .DataSource = dvSchemes
        End With

        ' You must clear out the TableStyles collection before 
        grdScheme.TableStyles.Clear()

        Dim grdTableStyle1 As New DataGridTableStyle
        With grdTableStyle1
            ' You must always set the MappingName, even with a DataView that has
            ' only one Table. If not, you will get no errors but the formatting
            ' will not appear. To avoid mistakes, instead of typing the name of 
            ' the table used when creating the DataSet, you can access the 
            ' TableName property.
            .MappingName = dvSchemes.Table.TableName
        End With

        ' The use of column styles overrides the automatic generation of columns 
        ' for every column in the DataTable. When column style objects are used,
        ' every column you want to display has to have an associate column style 
        ' object.
        Dim grdColStyle1 As New DataGridTextBoxColumn
        With grdColStyle1
            .MappingName = "AssemSize"
            .HeaderText = "Size"
            .Width = 65
        End With

        Dim grdColStyle2 As New DataGridTextBoxColumn
        With grdColStyle2
            .MappingName = "price"
            .HeaderText = "Price"
            .Width = 70
            .Format = "c"
        End With

        ' Add the column style objects to the table style's collection of 
        ' column styles. Without this the styles do not take effect.        
        grdTableStyle1.GridColumnStyles.AddRange _
            (New DataGridColumnStyle() {grdColStyle1, grdColStyle2})

        grdScheme.TableStyles.Add(grdTableStyle1)

        Dim bn As Binding
        For Each bn In grdScheme.DataBindings
            Trace.Write(bn.Control.ToString & ", " & bn.PropertyName)
            Trace.Write(vbCrLf)
        Next
    End Sub

    Private Sub NextScheme()
        PricingScheme += 1
        BindSchemesGrid()
    End Sub

    Private Sub PreviousScheme()
        PricingScheme -= 1
        BindSchemesGrid()
    End Sub

    Private Sub btnRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRight.Click
        btnLeft.Enabled = True
        If PricingScheme < MaxSchemeNo Then
            NextScheme()
            If PricingScheme = MaxSchemeNo Then
                btnRight.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLeft.Click
        btnRight.Enabled = True
        If PricingScheme >= 2 Then
            PreviousScheme()
            If PricingScheme = 1 Then
                btnLeft.Enabled = False
            End If
        End If
    End Sub



    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim index As Integer
        Dim lookUp As Integer
        Dim Price As Double

        Dim cn As New SqlConnection(connectionString)
        Dim strQuery As String = "SELECT MAX(manSchemeNo) FROM PricingSchemes"
        Dim cmd As New SqlCommand
        cmd.Connection = cn

        Try
            cn.Open() ' Open Connection

            For index = 0 To grdScheme.VisibleRowCount - 1
                lookUp = CInt(dvSchemes.Item(index).Item("priceNo"))
                Price = CDbl(grdScheme.Item(index, 1))
                cmd.CommandText = "UPDATE Price SET price = " & Price & " WHERE priceNo = " & lookUp
                cmd.ExecuteNonQuery()
            Next


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

End Class
