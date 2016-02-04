Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports BackflowTestingManager.CSharp
Imports log4net





Public Class MainForm
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
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBar1 As System.Windows.Forms.ToolBar
    Friend WithEvents grpDetails As System.Windows.Forms.GroupBox
    Friend WithEvents btnMap As System.Windows.Forms.Button
    Friend WithEvents lblImage As System.Windows.Forms.Label
    Friend WithEvents tbbAddManager As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbAddProp As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbFind As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbEdit As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbDelete As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbCut As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbPaste As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbAddAssembly As System.Windows.Forms.ToolBarButton
    Friend WithEvents cmTree As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuAddManager As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddProperty As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddAssembly As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFind As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMunicipalities As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPropToMan As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents grdTests As System.Windows.Forms.DataGrid
    Friend WithEvents llblEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents btnPhoneOrTest As System.Windows.Forms.Button
    Friend WithEvents lblRight As System.Windows.Forms.Label
    Friend WithEvents lblLeft As System.Windows.Forms.Label
    Friend WithEvents mnuPricing As System.Windows.Forms.MenuItem
    Friend WithEvents lblBottom As System.Windows.Forms.Label
    Friend WithEvents cmTests As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDeleteTest As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuManList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMunList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRetestList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRetestLetters As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents TreeView2 As System.Windows.Forms.TreeView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents mnuTestReports As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTestsByManager As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTestsByMunicipality As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTopManagers As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTestComparisonBTWYears As System.Windows.Forms.MenuItem
    Friend WithEvents btnPostToAccounting As System.Windows.Forms.Button
    Friend WithEvents mnuOptions As System.Windows.Forms.MenuItem
    Friend WithEvents CompanyInfo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuManagerTest As System.Windows.Forms.MenuItem
    Friend WithEvents tbbRefresh As System.Windows.Forms.ToolBarButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.cmTree = New System.Windows.Forms.ContextMenu()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuAddManager = New System.Windows.Forms.MenuItem()
        Me.mnuAddProperty = New System.Windows.Forms.MenuItem()
        Me.mnuAddAssembly = New System.Windows.Forms.MenuItem()
        Me.CompanyInfo = New System.Windows.Forms.MenuItem()
        Me.mnuOptions = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.mnuCut = New System.Windows.Forms.MenuItem()
        Me.mnuPaste = New System.Windows.Forms.MenuItem()
        Me.mnuEdit = New System.Windows.Forms.MenuItem()
        Me.mnuDelete = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.mnuFind = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.mnuDeleteTest = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuPricing = New System.Windows.Forms.MenuItem()
        Me.mnuMunicipalities = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.mnuPropToMan = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.mnuManList = New System.Windows.Forms.MenuItem()
        Me.mnuMunList = New System.Windows.Forms.MenuItem()
        Me.mnuRetestList = New System.Windows.Forms.MenuItem()
        Me.mnuRetestLetters = New System.Windows.Forms.MenuItem()
        Me.mnuTestReports = New System.Windows.Forms.MenuItem()
        Me.mnuTestsByManager = New System.Windows.Forms.MenuItem()
        Me.mnuManagerTest = New System.Windows.Forms.MenuItem()
        Me.mnuTestsByMunicipality = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.mnuTopManagers = New System.Windows.Forms.MenuItem()
        Me.mnuTestComparisonBTWYears = New System.Windows.Forms.MenuItem()
        Me.ToolBar1 = New System.Windows.Forms.ToolBar()
        Me.tbbAddManager = New System.Windows.Forms.ToolBarButton()
        Me.tbbAddProp = New System.Windows.Forms.ToolBarButton()
        Me.tbbAddAssembly = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton()
        Me.tbbEdit = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton()
        Me.tbbDelete = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton()
        Me.tbbFind = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton4 = New System.Windows.Forms.ToolBarButton()
        Me.tbbCut = New System.Windows.Forms.ToolBarButton()
        Me.tbbPaste = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton()
        Me.tbbRefresh = New System.Windows.Forms.ToolBarButton()
        Me.grpDetails = New System.Windows.Forms.GroupBox()
        Me.lblBottom = New System.Windows.Forms.Label()
        Me.llblEmail = New System.Windows.Forms.LinkLabel()
        Me.btnPhoneOrTest = New System.Windows.Forms.Button()
        Me.lblRight = New System.Windows.Forms.Label()
        Me.btnMap = New System.Windows.Forms.Button()
        Me.lblImage = New System.Windows.Forms.Label()
        Me.lblLeft = New System.Windows.Forms.Label()
        Me.grdTests = New System.Windows.Forms.DataGrid()
        Me.cmTests = New System.Windows.Forms.ContextMenu()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.TreeView2 = New System.Windows.Forms.TreeView()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.btnPostToAccounting = New System.Windows.Forms.Button()
        Me.grpDetails.SuspendLayout()
        CType(Me.grdTests, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TreeView1
        '
        Me.TreeView1.AllowDrop = True
        Me.TreeView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeView1.ContextMenu = Me.cmTree
        Me.TreeView1.Cursor = System.Windows.Forms.Cursors.Default
        Me.TreeView1.HideSelection = False
        Me.TreeView1.ImageIndex = 0
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.ItemHeight = 18
        Me.TreeView1.Location = New System.Drawing.Point(0, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = 0
        Me.TreeView1.Size = New System.Drawing.Size(280, 320)
        Me.TreeView1.TabIndex = 0
        '
        'cmTree
        '
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        Me.ImageList1.Images.SetKeyName(4, "")
        Me.ImageList1.Images.SetKeyName(5, "")
        Me.ImageList1.Images.SetKeyName(6, "")
        Me.ImageList1.Images.SetKeyName(7, "")
        Me.ImageList1.Images.SetKeyName(8, "")
        Me.ImageList1.Images.SetKeyName(9, "")
        Me.ImageList1.Images.SetKeyName(10, "")
        Me.ImageList1.Images.SetKeyName(11, "")
        Me.ImageList1.Images.SetKeyName(12, "")
        Me.ImageList1.Images.SetKeyName(13, "")
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem7, Me.MenuItem3, Me.MenuItem6, Me.MenuItem5})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.CompanyInfo, Me.mnuOptions, Me.mnuExit})
        Me.MenuItem1.Text = "F&ile"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddManager, Me.mnuAddProperty, Me.mnuAddAssembly})
        Me.MenuItem2.Text = "&New"
        '
        'mnuAddManager
        '
        Me.mnuAddManager.Index = 0
        Me.mnuAddManager.Text = "New Manager"
        '
        'mnuAddProperty
        '
        Me.mnuAddProperty.Index = 1
        Me.mnuAddProperty.Text = "New Property"
        '
        'mnuAddAssembly
        '
        Me.mnuAddAssembly.Index = 2
        Me.mnuAddAssembly.Text = "New Assembly"
        '
        'CompanyInfo
        '
        Me.CompanyInfo.Index = 1
        Me.CompanyInfo.Text = "Company Info"
        '
        'mnuOptions
        '
        Me.mnuOptions.Index = 2
        Me.mnuOptions.Text = "Options"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 3
        Me.mnuExit.Text = "E&xit"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCut, Me.mnuPaste, Me.mnuEdit, Me.mnuDelete, Me.MenuItem12, Me.mnuFind, Me.MenuItem4, Me.mnuDeleteTest})
        Me.MenuItem7.Text = "&Edit"
        '
        'mnuCut
        '
        Me.mnuCut.Index = 0
        Me.mnuCut.Text = "Cu&t"
        '
        'mnuPaste
        '
        Me.mnuPaste.Index = 1
        Me.mnuPaste.Text = "&Paste"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 2
        Me.mnuEdit.Text = "Edi&t"
        '
        'mnuDelete
        '
        Me.mnuDelete.Index = 3
        Me.mnuDelete.Text = "&Delete"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 4
        Me.MenuItem12.Text = "-"
        '
        'mnuFind
        '
        Me.mnuFind.Index = 5
        Me.mnuFind.Text = "&Find"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 6
        Me.MenuItem4.Text = "-"
        '
        'mnuDeleteTest
        '
        Me.mnuDeleteTest.Index = 7
        Me.mnuDeleteTest.Text = "Delete Test"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPricing, Me.mnuMunicipalities})
        Me.MenuItem3.Text = "Lists"
        '
        'mnuPricing
        '
        Me.mnuPricing.Index = 0
        Me.mnuPricing.Text = "Pricing Schemes"
        '
        'mnuMunicipalities
        '
        Me.mnuMunicipalities.Index = 1
        Me.mnuMunicipalities.Text = "Municipalities"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPropToMan})
        Me.MenuItem6.Text = "Functions"
        '
        'mnuPropToMan
        '
        Me.mnuPropToMan.Index = 0
        Me.mnuPropToMan.Text = "Property to Manager"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 4
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuManList, Me.mnuMunList, Me.mnuRetestList, Me.mnuRetestLetters, Me.mnuTestReports, Me.mnuTestsByManager, Me.mnuManagerTest, Me.mnuTestsByMunicipality, Me.MenuItem10, Me.mnuTopManagers, Me.mnuTestComparisonBTWYears})
        Me.MenuItem5.Text = "Reports"
        '
        'mnuManList
        '
        Me.mnuManList.Index = 0
        Me.mnuManList.Text = "Manager List"
        '
        'mnuMunList
        '
        Me.mnuMunList.Index = 1
        Me.mnuMunList.Text = "Municipality List"
        '
        'mnuRetestList
        '
        Me.mnuRetestList.Index = 2
        Me.mnuRetestList.Text = "Retest List"
        '
        'mnuRetestLetters
        '
        Me.mnuRetestLetters.Index = 3
        Me.mnuRetestLetters.Text = "Retest Letters"
        '
        'mnuTestReports
        '
        Me.mnuTestReports.Index = 4
        Me.mnuTestReports.Text = "Test Reports"
        '
        'mnuTestsByManager
        '
        Me.mnuTestsByManager.Index = 5
        Me.mnuTestsByManager.Text = "Tests By Manager"
        '
        'mnuManagerTest
        '
        Me.mnuManagerTest.Index = 6
        Me.mnuManagerTest.Text = "Manager Test"
        '
        'mnuTestsByMunicipality
        '
        Me.mnuTestsByMunicipality.Index = 7
        Me.mnuTestsByMunicipality.Text = "Tests By Municipality"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 8
        Me.MenuItem10.Text = "-"
        '
        'mnuTopManagers
        '
        Me.mnuTopManagers.Index = 9
        Me.mnuTopManagers.Text = "Top Managers In Last Year"
        '
        'mnuTestComparisonBTWYears
        '
        Me.mnuTestComparisonBTWYears.Index = 10
        Me.mnuTestComparisonBTWYears.Text = "Test Comparison Between Last Two Years"
        '
        'ToolBar1
        '
        Me.ToolBar1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ToolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tbbAddManager, Me.tbbAddProp, Me.tbbAddAssembly, Me.ToolBarButton1, Me.tbbEdit, Me.ToolBarButton2, Me.tbbDelete, Me.ToolBarButton3, Me.tbbFind, Me.ToolBarButton4, Me.tbbCut, Me.tbbPaste, Me.ToolBarButton5, Me.tbbRefresh})
        Me.ToolBar1.DropDownArrows = True
        Me.ToolBar1.ImageList = Me.ImageList1
        Me.ToolBar1.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar1.Name = "ToolBar1"
        Me.ToolBar1.ShowToolTips = True
        Me.ToolBar1.Size = New System.Drawing.Size(1012, 29)
        Me.ToolBar1.TabIndex = 7
        '
        'tbbAddManager
        '
        Me.tbbAddManager.ImageIndex = 0
        Me.tbbAddManager.Name = "tbbAddManager"
        Me.tbbAddManager.Tag = "tester"
        Me.tbbAddManager.ToolTipText = "New Manager"
        '
        'tbbAddProp
        '
        Me.tbbAddProp.ImageIndex = 1
        Me.tbbAddProp.Name = "tbbAddProp"
        Me.tbbAddProp.ToolTipText = "New Property"
        '
        'tbbAddAssembly
        '
        Me.tbbAddAssembly.ImageIndex = 2
        Me.tbbAddAssembly.Name = "tbbAddAssembly"
        Me.tbbAddAssembly.ToolTipText = "New Assembly"
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.Name = "ToolBarButton1"
        Me.ToolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbEdit
        '
        Me.tbbEdit.ImageIndex = 5
        Me.tbbEdit.Name = "tbbEdit"
        Me.tbbEdit.ToolTipText = "Edit"
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.Name = "ToolBarButton2"
        Me.ToolBarButton2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbDelete
        '
        Me.tbbDelete.ImageIndex = 4
        Me.tbbDelete.Name = "tbbDelete"
        Me.tbbDelete.ToolTipText = "Delete"
        '
        'ToolBarButton3
        '
        Me.ToolBarButton3.Name = "ToolBarButton3"
        Me.ToolBarButton3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbFind
        '
        Me.tbbFind.ImageIndex = 3
        Me.tbbFind.Name = "tbbFind"
        Me.tbbFind.ToolTipText = "Find"
        '
        'ToolBarButton4
        '
        Me.ToolBarButton4.Name = "ToolBarButton4"
        Me.ToolBarButton4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbCut
        '
        Me.tbbCut.ImageIndex = 11
        Me.tbbCut.Name = "tbbCut"
        Me.tbbCut.ToolTipText = "Cut"
        '
        'tbbPaste
        '
        Me.tbbPaste.ImageIndex = 12
        Me.tbbPaste.Name = "tbbPaste"
        Me.tbbPaste.ToolTipText = "Paste"
        '
        'ToolBarButton5
        '
        Me.ToolBarButton5.Name = "ToolBarButton5"
        Me.ToolBarButton5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbbRefresh
        '
        Me.tbbRefresh.ImageIndex = 13
        Me.tbbRefresh.Name = "tbbRefresh"
        Me.tbbRefresh.ToolTipText = "Refresh TreeView"
        '
        'grpDetails
        '
        Me.grpDetails.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpDetails.Controls.Add(Me.lblBottom)
        Me.grpDetails.Controls.Add(Me.llblEmail)
        Me.grpDetails.Controls.Add(Me.btnPhoneOrTest)
        Me.grpDetails.Controls.Add(Me.lblRight)
        Me.grpDetails.Controls.Add(Me.btnMap)
        Me.grpDetails.Controls.Add(Me.lblImage)
        Me.grpDetails.Controls.Add(Me.lblLeft)
        Me.grpDetails.Location = New System.Drawing.Point(287, 32)
        Me.grpDetails.Name = "grpDetails"
        Me.grpDetails.Size = New System.Drawing.Size(724, 104)
        Me.grpDetails.TabIndex = 11
        Me.grpDetails.TabStop = False
        '
        'lblBottom
        '
        Me.lblBottom.Location = New System.Drawing.Point(32, 80)
        Me.lblBottom.Name = "lblBottom"
        Me.lblBottom.Size = New System.Drawing.Size(608, 16)
        Me.lblBottom.TabIndex = 22
        '
        'llblEmail
        '
        Me.llblEmail.Location = New System.Drawing.Point(376, 64)
        Me.llblEmail.Name = "llblEmail"
        Me.llblEmail.Size = New System.Drawing.Size(248, 16)
        Me.llblEmail.TabIndex = 21
        '
        'btnPhoneOrTest
        '
        Me.btnPhoneOrTest.Location = New System.Drawing.Point(568, 29)
        Me.btnPhoneOrTest.Name = "btnPhoneOrTest"
        Me.btnPhoneOrTest.Size = New System.Drawing.Size(72, 32)
        Me.btnPhoneOrTest.TabIndex = 17
        Me.btnPhoneOrTest.Text = "Phonebook"
        '
        'lblRight
        '
        Me.lblRight.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblRight.Location = New System.Drawing.Point(360, 16)
        Me.lblRight.Name = "lblRight"
        Me.lblRight.Size = New System.Drawing.Size(200, 48)
        Me.lblRight.TabIndex = 16
        '
        'btnMap
        '
        Me.btnMap.Location = New System.Drawing.Point(240, 29)
        Me.btnMap.Name = "btnMap"
        Me.btnMap.Size = New System.Drawing.Size(64, 32)
        Me.btnMap.TabIndex = 15
        Me.btnMap.Text = "Map"
        '
        'lblImage
        '
        Me.lblImage.ImageAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblImage.ImageList = Me.ImageList1
        Me.lblImage.Location = New System.Drawing.Point(8, 16)
        Me.lblImage.Name = "lblImage"
        Me.lblImage.Size = New System.Drawing.Size(24, 56)
        Me.lblImage.TabIndex = 14
        '
        'lblLeft
        '
        Me.lblLeft.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblLeft.Location = New System.Drawing.Point(32, 16)
        Me.lblLeft.Name = "lblLeft"
        Me.lblLeft.Size = New System.Drawing.Size(200, 56)
        Me.lblLeft.TabIndex = 13
        '
        'grdTests
        '
        Me.grdTests.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdTests.CaptionBackColor = System.Drawing.Color.LightSlateGray
        Me.grdTests.CaptionForeColor = System.Drawing.Color.White
        Me.grdTests.ContextMenu = Me.cmTests
        Me.grdTests.DataMember = ""
        Me.grdTests.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdTests.Location = New System.Drawing.Point(288, 144)
        Me.grdTests.Name = "grdTests"
        Me.grdTests.Size = New System.Drawing.Size(724, 360)
        Me.grdTests.TabIndex = 12
        '
        'cmTests
        '
        '
        'txtSearch
        '
        Me.txtSearch.Location = New System.Drawing.Point(10, 10)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(184, 20)
        Me.txtSearch.TabIndex = 13
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(202, 4)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 31)
        Me.btnSearch.TabIndex = 14
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'TreeView2
        '
        Me.TreeView2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TreeView2.ImageIndex = 0
        Me.TreeView2.ImageList = Me.ImageList1
        Me.TreeView2.Location = New System.Drawing.Point(0, 40)
        Me.TreeView2.Name = "TreeView2"
        Me.TreeView2.SelectedImageIndex = 0
        Me.TreeView2.Size = New System.Drawing.Size(280, 272)
        Me.TreeView2.TabIndex = 15
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Left
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 29)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TreeView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtSearch)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TreeView2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSearch)
        Me.SplitContainer1.Size = New System.Drawing.Size(280, 640)
        Me.SplitContainer1.SplitterDistance = 320
        Me.SplitContainer1.TabIndex = 16
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MonthCalendar1.CalendarDimensions = New System.Drawing.Size(3, 1)
        Me.MonthCalendar1.ForeColor = System.Drawing.SystemColors.GrayText
        Me.MonthCalendar1.Location = New System.Drawing.Point(289, 509)
        Me.MonthCalendar1.MaxSelectionCount = 1
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.ShowTodayCircle = False
        Me.MonthCalendar1.TabIndex = 6
        Me.MonthCalendar1.TitleBackColor = System.Drawing.Color.LightSlateGray
        Me.MonthCalendar1.TitleForeColor = System.Drawing.Color.White
        '
        'btnPostToAccounting
        '
        Me.btnPostToAccounting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPostToAccounting.Location = New System.Drawing.Point(855, 560)
        Me.btnPostToAccounting.Name = "btnPostToAccounting"
        Me.btnPostToAccounting.Size = New System.Drawing.Size(124, 23)
        Me.btnPostToAccounting.TabIndex = 17
        Me.btnPostToAccounting.Text = "Post to Accounting"
        Me.btnPostToAccounting.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1012, 669)
        Me.ContextMenu = Me.cmTests
        Me.Controls.Add(Me.grdTests)
        Me.Controls.Add(Me.btnPostToAccounting)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.grpDetails)
        Me.Controls.Add(Me.ToolBar1)
        Me.Controls.Add(Me.MonthCalendar1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "MainForm"
        Me.Text = "Backflow Testing Manager"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.grpDetails.ResumeLayout(False)
        CType(Me.grdTests, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Public Shared logger As ILog = Nothing

    Protected connectionString As String

    Public Shared ds As DataSet
    Private dtTests As DataTable
    Private dvTests As DataView
    Private strLabelLeft As String
    Private strLabelRight As String
    Private strLabelBottom As String
    Private TheCutNode As myTreeNode = Nothing
    Private intPrevMan As Integer
    Private intPrevProp As Integer

    Private ComName As String
    Private Street As String
    Private City As String
    Private State As String
    Private Zip As String
    Private Contact As String
    Private Phone As String
    Private Fax As String
    Private Email As String

    Private treeViewRef As TreeView = TreeView1


    ' Handle the Form load event.
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try


            connectionString = ConfigurationManager.ConnectionStrings("bftm.My.MySettings.qbConnectionString1").ConnectionString
            logger = LogManager.GetLogger("GeneralLogger")
            logger.Info("frmMain_Load - Init")

            ' Fill the DataTables for ComboBoxes
            FillDataTables()
            ' Fill the TreeView
            FillTreeView()
            RefreshTestList()
            Me.WindowState() = FormWindowState.Maximized
            If dvTests.Count > 0 Then
                mnuDeleteTest.Enabled = True
            Else
                mnuDeleteTest.Enabled = False
            End If
            mnuPaste.Enabled = False
            tbbPaste.Enabled = False

            TreeView1.Select()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)


        End Try
    End Sub

    Private Sub FillDataTables()
        ' Instantiate the dataset 
        ds = New DataSet ' Create a new dataset
        ' Establish connection object
        Try
            Dim cn As New SqlConnection(connectionString)
            Dim cmd As New SqlCommand
            cmd.Connection = cn
            Dim sda As New SqlDataAdapter(cmd)

            ' Fill the tables for the comboboxes

            ' Reset the CommandText to get the Assembly Sizes
            cmd.CommandText = "SELECT AssemSize, AssemSizeNo FROM AssemblySizes"
            sda.Fill(ds, "AssemSizes") ' Fill Assembly Sizes datatable
            ' Reset the CommandText to get the Municipality
            cmd.CommandText = "SELECT munNo, munName FROM Municipalities ORDER BY munName"
            sda.Fill(ds, "Municip") ' Fill the Municipalities datatable
            cmd.CommandText = "SELECT tstrName, tstrNo FROM Testers"
            sda.Fill(ds, "Testers") ' Fill the Testers datatable.
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillTreeView()
        ' Create a New SQL Connection
        logger.Info("FillTreeView - Enter")
        Dim cn As New SqlConnection(connectionString)
        ' Query which returns a resultset like the following
        ' Managers No., Manager's Name, Property No., Property's Name, Assembly No., Assembly's Serial No.

        Dim cmd As New SqlCommand("BFTM_TreeView_SP", cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@searchString", ""))

        Try
            cn.Open() ' Open connection

            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            ' Display a wait cursor while the TreeNodes are being created.
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            ' Suppress repainting the TreeView until all the objects have been created.
            TreeView1.BeginUpdate()

            ' Clear the TreeView each time the method is called.
            TreeView1.Nodes.Clear()

            ' Variables needed to iterate through the resultset
            Dim currentManager As Integer
            Dim lastManager As Integer = 0
            Dim currentProperty As Integer
            Dim lastProperty As Integer
            Dim currentAssembly As Integer
            Dim lastAssembly As Integer
            Dim propDeleted As Boolean
            Dim assemDeleted As Boolean
            Dim currentPropertyName As String
            Dim PropertyName As String
            Dim storeNo As String

            ' Ititial position values for populating the TreeView
            Dim j As Integer = -1
            Dim k As Integer = -1

            Do While dr.Read()
                currentManager = CInt(dr.Item("manNo"))
                If currentManager <> lastManager Then
                    TreeView1.Nodes.Add(New myTreeNode(dr.Item("manName").ToString(), "Manager", currentManager, True, 0, 0))
                    lastManager = currentManager
                    j += 1 ' Increment the Manager Index
                    k = -1 ' Reset the Property Index once a new manager is added
                End If
                If Not IsDBNull(dr.Item("propNo")) Then
                    currentProperty = CInt(dr.Item("propNo"))
                    If currentProperty <> lastProperty And Boolean.Parse(dr.Item("propDeleted").ToString()) = False Then
                        If IsDBNull(dr.Item("storeNo")) Then
                            storeNo = ""
                        ElseIf dr.Item("storeNo").ToString() = "" Then
                            storeNo = ""
                        Else
                            storeNo = " - " & dr.Item("storeNo").ToString()
                        End If
                        ' This is around where I have to look into
                        TreeView1.Nodes(j).Nodes.Add( _
                        New myTreeNode(dr.Item("propName").ToString() & storeNo, "Property", currentProperty, True, 1, 1))
                        lastProperty = currentProperty
                        k += 1 ' Increment the Property Index
                    End If
                End If
                If Not IsDBNull(dr.Item("assemNo")) Then
                    currentAssembly = CInt(dr.Item("assemNo"))
                    If currentAssembly <> lastAssembly And Boolean.Parse(dr.Item("assemDeleted").ToString()) = False Then
                        TreeView1.Nodes(j).Nodes(k).Nodes.Add( _
                            New myTreeNode(dr.Item("assemSerial").ToString() & "", "Assembly", currentAssembly, True, 2, 2))
                    End If
                End If
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
            logger.Error(ex.ToString(), ex)
            MessageBox.Show(msg)

        Catch ex As Exception
            ' A generic exception has occured
            logger.Error(ex.ToString(), ex)
            MessageBox.Show(ex.Message)
        Finally
            ' Close the connection
            cn.Close()
        End Try

        ' Reset the cursor to the default for all controls.
        Cursor.Current = System.Windows.Forms.Cursors.Default

        ' Begin repainting the TreeView.
        TreeView1.EndUpdate()

        logger.Info("FillTreeView - Exit")
    End Sub 'FillTreeView


    ' Toolbar buttons are enabled or disabled based on the Node selected in the Treeview
    ' in addition must update the labels after after a node is selected by calling FillLabels
    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        treeViewRef = TreeView1
        EnableMenuItems()
    End Sub

    Private Sub EnableMenuItems()
        FillLabels(CType(treeViewRef.SelectedNode, myTreeNode))
        If Not treeViewRef.SelectedNode Is Nothing Then
            Dim N As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)
            If N.Tag.ToString() = "Manager" Then
                mnuAddProperty.Enabled = True
                mnuAddAssembly.Enabled = False
                mnuCut.Enabled = False
                If Not TheCutNode Is Nothing Then
                    If TheCutNode.Tag.ToString() = "Property" Then
                        mnuPaste.Enabled = True
                        tbbPaste.Enabled = True
                    Else
                        mnuPaste.Enabled = False
                        tbbPaste.Enabled = False
                    End If
                End If
                mnuPropToMan.Enabled = False
                tbbAddProp.Enabled = True
                tbbAddAssembly.Enabled = False
                tbbCut.Enabled = False
            End If
            If N.Tag.ToString() = "Property" Then
                mnuAddProperty.Enabled = False
                mnuAddAssembly.Enabled = True
                mnuPropToMan.Enabled = True

                If Not TheCutNode Is Nothing Then
                    If TheCutNode.Tag.ToString() = "Assembly" Then
                        mnuPaste.Enabled = True
                        tbbPaste.Enabled = True
                    Else
                        mnuPaste.Enabled = False
                        tbbPaste.Enabled = False
                    End If
                    mnuCut.Enabled = False
                    tbbCut.Enabled = False
                Else
                    mnuCut.Enabled = True
                    tbbCut.Enabled = True
                End If
                tbbAddProp.Enabled = False
                tbbAddAssembly.Enabled = True
            End If
            If N.Tag.ToString() = "Assembly" Then
                mnuAddProperty.Enabled = False
                mnuAddAssembly.Enabled = False
                mnuPropToMan.Enabled = False

                mnuCut.Enabled = False
                tbbCut.Enabled = False

                mnuPaste.Enabled = False
                tbbAddProp.Enabled = False
                tbbAddAssembly.Enabled = False
                tbbPaste.Enabled = False
            End If
        End If
    End Sub

    Private Sub FillSearchTree(ByVal searchString As String)
        TreeView2.Nodes.Clear()
        ' Create a New SQL Connection
        Dim cn As New SqlConnection(connectionString)
        ' Query which returns a resultset like the following
        ' Managers No., Manager's Name, Property No., Property's Name, Assembly No., Assembly's Serial No.

        Dim cmd As New SqlCommand("BFTM_TreeView_SP", cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@searchString", searchString))

        Try
            cn.Open() ' Open connection

            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            ' Display a wait cursor while the TreeNodes are being created.
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            ' Suppress repainting the TreeView until all the objects have been created.
            TreeView2.BeginUpdate()

            ' Clear the TreeView each time the method is called.
            TreeView2.Nodes.Clear()

            ' Variables needed to iterate through the resultset
            Dim currentManager As Integer
            Dim lastManager As Integer = 0
            Dim currentProperty As Integer
            Dim lastProperty As Integer
            Dim currentAssembly As Integer
            Dim lastAssembly As Integer
            Dim propDeleted As Boolean
            Dim assemDeleted As Boolean
            Dim currentPropertyName As String
            Dim PropertyName As String
            Dim storeNo As String

            ' Ititial position values for populating the TreeView
            Dim j As Integer = -1
            Dim k As Integer = -1

            Do While dr.Read()
                currentManager = CInt(dr.Item("manNo"))
                If currentManager <> lastManager Then
                    TreeView2.Nodes.Add(New myTreeNode(dr.Item("manName").ToString(), "Manager", currentManager, True, 0, 0))
                    lastManager = currentManager
                    j += 1 ' Increment the Manager Index
                    k = -1 ' Reset the Property Index once a new manager is added
                End If
                If Not IsDBNull(dr.Item("propNo")) Then
                    currentProperty = CInt(dr.Item("propNo"))
                    If currentProperty <> lastProperty And Boolean.Parse(dr.Item("propDeleted").ToString()) = False Then
                        If IsDBNull(dr.Item("storeNo")) Then
                            storeNo = ""
                        ElseIf dr.Item("storeNo").ToString() = "" Then
                            storeNo = ""
                        Else
                            storeNo = " - " & dr.Item("storeNo").ToString()
                        End If
                        ' This is around where I have to look into
                        TreeView2.Nodes(j).Nodes.Add( _
                        New myTreeNode(dr.Item("propName").ToString() & storeNo, "Property", currentProperty, True, 1, 1))
                        lastProperty = currentProperty
                        k += 1 ' Increment the Property Index
                    End If
                End If
                If Not IsDBNull(dr.Item("assemNo")) Then
                    currentAssembly = CInt(dr.Item("assemNo"))
                    If currentAssembly <> lastAssembly And Boolean.Parse(dr.Item("assemDeleted").ToString()) = False Then
                        TreeView2.Nodes(j).Nodes(k).Nodes.Add( _
                            New myTreeNode(dr.Item("assemSerial").ToString() & "", "Assembly", currentAssembly, True, 2, 2))
                    End If
                End If
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

        ' Reset the cursor to the default for all controls.
        Cursor.Current = System.Windows.Forms.Cursors.Default


        ' Begin repainting the TreeView.
        TreeView2.EndUpdate()
    End Sub


    ' Fill lblAddress and lblContact based on the node selected in the TreeView
    Private Sub FillLabels(ByRef N As myTreeNode)

        ' If node is a Manager display the Manager's information
        If N.Tag.ToString() = "Manager" Then
            DisplayManagerInfo(N)
            ' Else if Property display the Property's information
        ElseIf N.Tag.ToString() = "Property" Then
            DisplayPropertyInfo(N)
            ' Else if Assembly display the Assembly's information
        ElseIf N.Tag.ToString() = "Assembly" Then
            DisplayAssemblyInfo(N)
        End If

        lblLeft.Text = strLabelLeft
        lblRight.Text = strLabelRight
        lblBottom.Text = strLabelBottom
        llblEmail.Text = Email
        grpDetails.Text = N.Tag.ToString()
    End Sub

    Private Sub DisplayPropertyInfo(ByRef N As myTreeNode)

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn

        ' Displays the image next to the label information
        lblImage.ImageIndex = 1

        ' Displays the buttons in the label area
        btnMap.Visible = True
        btnPhoneOrTest.Text = "Phonebook"

        Try
            cn.Open() ' Open connection
            ' Get the Property Information based on the Type and ID provided in the Selected Node (myTreeNode)
            cmd.CommandText = "SELECT propName, propStrt, propCity, propState, propZip, propCon, propPhone, " _
                & "propFax, propEmail FROM Properties Where propNo = " & N.ID
            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            If dr.Read() Then
                ' Null + Empty = Empty
                ComName = dr.Item("propName").ToString() & ""
                Street = dr.Item("propStrt").ToString() & ""
                City = dr.Item("propCity").ToString() & ""
                State = dr.Item("propState").ToString() & ""
                Zip = dr.Item("propZip").ToString() & ""
                Contact = dr.Item("propCon").ToString() & ""
                Phone = dr.Item("propPhone").ToString() & ""
                Fax = dr.Item("propFax").ToString() & ""
                Email = dr.Item("propEmail").ToString() & ""
            End If

            dr.Close() ' Close DataReader

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
            cn.Close() ' Close the connection
        End Try

        ' Set the Property's information to the following strings
        strLabelLeft = ComName & ControlChars.CrLf & Street & ControlChars.CrLf & City & ", " _
            & State & " " & Zip
        strLabelRight = "Contact: " & Contact & ControlChars.CrLf & "Phone: " & Phone _
            & ControlChars.CrLf & "Fax: " & Fax
        strLabelBottom = ""

    End Sub

    Private Sub DisplayAssemblyInfo(ByRef N As myTreeNode)

        Dim Model As String
        Dim Serial As String
        Dim Usage As String
        Dim InstallDate As String
        Dim Location As String
        Dim Type As String
        Dim Manufacturer As String
        Dim Size As String
        Dim LastTest As String

        ' If the Selected Node is an Assembly then get the Assembly's info
        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn

        lblImage.ImageIndex = 2
        btnMap.Visible = False
        btnPhoneOrTest.Text = "Add Test"

        Try
            cn.Open() ' Open connection
            cmd.CommandText = "SELECT Assemblies.assemMod, Assemblies.assemSerial, Assemblies.assemUsage, " _
                & "Assemblies.assemInstDt, Assemblies.assemLoc, Assemblies.assemType, " _
                & "Assemblies.AssemMan, AssemblySizes.AssemSize " _
                & "FROM Assemblies " _
                & "LEFT JOIN AssemblySizes ON AssemblySizes.AssemSizeNo = Assemblies.assemSizeNo " _
                & "LEFT JOIN Tests ON Tests.assemNo = Assemblies.assemNo " _
                & "WHERE Assemblies.assemNo = " & N.ID

            Dim dr As SqlDataReader = cmd.ExecuteReader()

            If dr.Read() Then
                ' Null + Empty = Empty
                Model = dr.Item("assemMod").ToString() & ""
                Serial = dr.Item("assemSerial").ToString() & ""
                Usage = dr.Item("assemUsage").ToString() & ""
                If Not IsDBNull(dr.Item("assemInstDt")) Then
                    InstallDate = Format(CDate(dr.Item("assemInstDt")).ToString(), "DD/MM/YYYY")
                Else
                    InstallDate = ""
                End If
                Location = dr.Item("assemLoc").ToString() & ""
                Type = dr.Item("assemType").ToString() & ""
                Manufacturer = dr.Item("AssemMan").ToString() & ""
                Size = dr.Item("AssemSize").ToString() & ""
            End If
            Email = ""
            dr.Close() ' Close the DataReader
            cmd.CommandText = "SELECT MAX(Tests.testDate) FROM Tests " _
                & "LEFT JOIN Assemblies ON Assemblies.assemNo = Tests.assemNo " _
                & "WHERE Assemblies.assemNo = " & N.ID

            If IsDBNull(cmd.ExecuteScalar()) Then
                LastTest = Nothing
            Else
                LastTest = Format(CDate(cmd.ExecuteScalar()), "M/dd/yyyy")
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

        ' Set the Assembly's information to the following strings
        strLabelLeft = "Serial: " & Serial & ControlChars.CrLf & "Type: " & Type _
            & ControlChars.CrLf & "Manufacturer: " & Manufacturer & ControlChars.CrLf & "Model: " & Model
        strLabelRight = "Size: " & Size & ControlChars.CrLf & "Usage: " & Usage _
            & ControlChars.CrLf & "Last Test Date: " & LastTest & ControlChars.CrLf
        strLabelBottom = "Location: " & Location

    End Sub
    Private Sub DisplayManagerInfo(ByRef N As myTreeNode)

        ' Create SQL Connection
        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn

        ' Displays the image next to the label information
        lblImage.ImageIndex = 0
        ' Displays the buttons in the label area
        btnMap.Visible = True
        btnPhoneOrTest.Text = "Phonebook"

        Try
            ' Open the DataConnection
            cn.Open()
            ' Select the Manager information where the Manager No. is equal to the Manager No.
            ' stored in the myTreeNode of the Selected Node.
            cmd.CommandText = "SELECT manName, manStrtAdd, manCity, manState, manZip, " _
            & "manCntct, manPhone, manFax, manEmail FROM Managers Where manNo = " _
            & CType(treeViewRef.SelectedNode, myTreeNode).ID

            ' Run the query; get the DataReader object.
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            Dim str As String

            If dr.Read() Then
                ' Null + Empty = Empty
                ComName = dr.Item("manName").ToString() & ""
                Street = dr.Item("manStrtAdd").ToString() & ""
                City = dr.Item("manCity").ToString() & ""
                State = dr.Item("manState").ToString() & ""
                Zip = dr.Item("manZip").ToString() & ""
                Contact = dr.Item("manCntct").ToString() & ""
                Phone = dr.Item("manPhone").ToString() & ""
                Fax = dr.Item("manFax").ToString() & ""
                Email = dr.Item("manEmail").ToString() & ""
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

        ' Set the Managers info to the following strings
        strLabelLeft = ComName & ControlChars.CrLf & Street & ControlChars.CrLf & City _
            & ", " & State & " " & Zip
        strLabelRight = "Contact: " & Contact & ControlChars.CrLf & "Phone: " & Phone _
            & ControlChars.CrLf & "Fax: " & Fax
        strLabelBottom = ""
    End Sub


    Private Sub btnMap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMap.Click
        ' Get info to build mapquest string
        Street = Street.Trim
        City = City.Trim
        State = State.Trim
        ' Replace blank spaces in the strings with + for the Query String
        Street = Street.Replace(" ", "+")
        City = City.Replace(" ", "+")
        ' Build mapquest string
        Dim MapQuest As String = "http://www.mapquest.com/maps/map.adp?country=US&countryid=250" & _
            "&searchtype=address&address=" & Street & "&city=" & City & "&state=" & State.ToUpper & _
            "&zipcode=" & Zip & vbCrLf

        Dim psi As New ProcessStartInfo
        psi.UseShellExecute = True
        psi.FileName = MapQuest
        Try
            ' This throws an exception because the final webpage may differ from the one specified
            Process.Start(psi)
        Catch ex As Exception
            ' This may suffice
        End Try

    End Sub

    Private Sub btnPhoneORTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPhoneOrTest.Click
        If Not treeViewRef.SelectedNode Is Nothing Then
            Dim N As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)
            If N.Tag.ToString() = "Manager" Or N.Tag.ToString() = "Property" Then
                LookUpInPhonebook()
            ElseIf N.Tag.ToString() = "Assembly" Then
                AddTest(N.ID)
            End If
        End If
    End Sub

    Private Sub AddTest(ByVal AssemblyID As Integer)
        UpdateTests()
        Dim TestDate As Date = MonthCalendar1.SelectionStart()

        Dim cn As New SqlConnection(connectionString)
        Dim sqlStatement As String = "INSERT INTO Tests (assemNo, testDate, Notes, PONo, permitNo) VALUES " _
            & "('" & AssemblyID & "','" & TestDate & "','', '', '')"

        Dim cmd As New SqlCommand(sqlStatement, cn)

        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
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

        RefreshTestList()

    End Sub

    Private Sub RefreshTestList()
        Console.WriteLine("RefreshTestList Called")
        Dim chosenDate As Date = MonthCalendar1.SelectionStart()

        If Not ds.Tables("Tests") Is Nothing Then
            ds.Tables("Tests").Clear()
        End If

        Dim cn As New SqlConnection(connectionString)
        Dim sqlStatement As String = "SELECT Managers.manName, Properties.propName + ', ' + Properties.propStrt + ', ' + propCity As [PropertyAll], " _
            & "Properties.propStrt, Properties.propCity, Assemblies.assemSerial, " _
            & "Tests.testdate, Tests.testPass, Tests.tstrName, Tests.Notes, " _
            & "Tests.testNo, Tests.testPerformed, Tests.testHours, Tests.PONo, Tests.post, Tests.test_posted, Tests.permitNo " _
            & "FROM Managers " _
            & "INNER JOIN Properties ON Properties.manNo = Managers.manNo " _
            & "INNER JOIN Assemblies ON Assemblies.propNo = Properties.propNo " _
            & "INNER JOIN Tests ON Tests.assemNo = Assemblies.assemNo " _
            & "WHERE Tests.testdate = '" & chosenDate & "'"
        Dim cmd As New SqlCommand(sqlStatement, cn)
        Dim sda As New SqlDataAdapter(cmd)
        sda.Fill(ds, "Tests") ' Fill the Tests datatable

        dtTests = ds.Tables("Tests")
        dvTests = dtTests.DefaultView
        dvTests.AllowNew = False
        BindTests(chosenDate)

    End Sub

    Private Sub LookUpInPhonebook()
        ' Get info to build mapquest string
        ComName = ComName.Trim
        City = City.Trim
        State = State.Trim.ToUpper
        ComName = ComName.Replace(" ", "+")
        City = City.Replace(" ", "+")
        ' Build whitePages string

        Dim WhitePages As String = "http://www.smartpages.com/co/wp/index.jhtml?From=wp&QueryString=" & _
            ComName & "&QueryType=2&CityZipAC=" & City & "&State=" & State & "&ClearLevel=Cloud9"

        Dim psi As New ProcessStartInfo
        psi.UseShellExecute = True
        psi.FileName = WhitePages
        Try
            ' This throws an exception because the final webpage may differ from the one specified
            Process.Start(psi)
        Catch ex As Exception
            ' This may suffice
        End Try

    End Sub

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick

        ' Evaluate the Button property to determine which button was clicked.
        Select Case ToolBar1.Buttons.IndexOf(e.Button)
            Case 0
                NewManager()
            Case 1
                NewProperty()
            Case 2
                NewAssembly()
            Case 3
                ' Spacer
            Case 4
                Edit()
            Case 5
                ' Spacer
            Case 6
                DeleteNode()
            Case 7
                ' Spacer
            Case 8
                LookUp()
            Case 9
                ' Spacer
            Case 10
                CutNode()
            Case 11
                PasteNode()
            Case 12
                ' Spacer
            Case 13
                FillTreeView()
                If Not txtSearch.Text.Equals("") Then
                    FillSearchTree(txtSearch.Text.Trim())
                Else
                    TreeView2.Nodes.Clear()
                End If

        End Select

    End Sub

    Private Sub Edit()
        If Not treeViewRef.SelectedNode Is Nothing Then
            Dim N As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)
            If N.Tag.ToString() = "Manager" Then
                Dim manKeyHolder As New ManagerKeyHolder(Integer.Parse(N.ID.ToString()))
                Dim manager As New Manager(manKeyHolder)
                Dim manAdapt As New ManagerAdapter()
                manAdapt.Fill(manager)
                Dim frm As New ManagerForm(manager)
                frm.Text = "Manager Detail"

                If frm.ShowDialog() = DialogResult.OK Then
                    Dim qbCustAdapt As New QuickbooksCustomerAdapter()
                    If qbCustAdapt.Update(manager) = True Then
                        manAdapt.Update(manager)
                        treeViewRef.SelectedNode.Text = manager.Name
                        FillLabels(N)
                    End If
                End If
            End If
            If N.Tag.ToString() = "Property" Then
                Dim frm As New PropertyDetailsForm
                frm.txtPropNo.Text = N.ID.ToString() ' use txtPropNo to pass the propNo to the form
                If frm.ShowDialog() = DialogResult.OK Then
                    If frm.txtPropStore.Text Is Nothing Or frm.txtPropStore.Text.Trim = "" Then
                        N.Text = frm.txtPropName.Text
                    Else
                        N.Text = frm.txtPropName.Text & " - " & frm.txtPropStore.Text
                    End If

                    FillLabels(N)
                End If
            End If
            If N.Tag.ToString() = "Assembly" Then
                Dim frm As New AssemblyDetailForm
                frm.txtAssemNo.Text = N.ID.ToString() ' use txtAssemNo to pass the assemNo to the form
                If frm.ShowDialog() = DialogResult.OK Then
                    treeViewRef.SelectedNode.Text = frm.txtSerial.Text
                    FillLabels(N)
                End If
            End If
        End If
        'If treeViewRef.Equals(TreeView1) Then
        '    If Not txtSearch.Text.Equals("") Then
        '        FillSearchTree(txtSearch.Text.Trim())
        '    Else
        '        TreeView2.Nodes.Clear()
        '    End If
        'ElseIf treeViewRef.Equals(TreeView2) Then
        '    FillTreeView()
        'End If

    End Sub

    Private Sub NewManager()

        Dim manKeyHolder As New ManagerKeyHolder(0)
        Dim manager As New Manager(manKeyHolder)
        ' Set new managers to the first pricing scheme
        manager.PricingScheme = 1
        Dim frm As New ManagerForm(manager)
        frm.Text = "Manager Detail"
        If frm.ShowDialog() = DialogResult.OK Then
            Dim qbCustAdapt As New QuickbooksCustomerAdapter()
            If qbCustAdapt.Update(manager) = True Then
                Dim manAdapt As New ManagerAdapter()
                manAdapt.Update(manager)
                TreeView1.Nodes.Add(New myTreeNode(manager.Name, "Manager", manager.PrimaryKey, True, 0, 0))
            End If


        End If



    End Sub

    Private Sub NewProperty()

        Dim treeViewLabel As String
        Dim frm As New PropertyAddForm
        Dim strTitle As String = "New Property for " & CType(treeViewRef.SelectedNode, myTreeNode).Text
        Dim manNo As Integer = CInt(CType(treeViewRef.SelectedNode, myTreeNode).ID)
        frm.Text = strTitle
        frm.manNumber = manNo
        If frm.ShowDialog() = DialogResult.OK Then
            Dim cn As New SqlConnection(connectionString)
            Dim sql As String = "SELECT MAX(propNo) FROM Properties"
            Dim cmd As New SqlCommand(sql, cn)
            Try
                cn.Open()
                Dim propNo As Integer = CInt(cmd.ExecuteScalar())
                If frm.txtPropStore.Text Is Nothing Or frm.txtPropStore.Text.Trim = "" Then
                    treeViewLabel = frm.txtPropName.Text
                Else
                    treeViewLabel = frm.txtPropName.Text & " - " & frm.txtPropStore.Text
                End If
                treeViewRef.SelectedNode.Nodes.Add(New myTreeNode(treeViewLabel, "Property", propNo, True, 1, 1))
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
        End If

    End Sub

    Private Sub NewAssembly()
        Dim strTitle As String
        strTitle = "New Assembly for " & treeViewRef.SelectedNode.Text
        Dim propNo As Integer = CInt(CType(treeViewRef.SelectedNode, myTreeNode).ID)
        Dim frm As New AssemblyAddForm
        frm.Text = strTitle
        frm.propNumber = propNo
        If frm.ShowDialog() = DialogResult.OK Then
            Dim cn As New SqlConnection(connectionString)
            Dim sql As String = "SELECT MAX(assemNo) FROM Assemblies"
            Dim cmd As New SqlCommand(sql, cn)
            Try
                cn.Open()
                Dim assemNo As Integer = CInt(cmd.ExecuteScalar())
                treeViewRef.SelectedNode.Nodes.Add(New myTreeNode(frm.txtSerial.Text, "Assembly", assemNo, True, 2, 2))
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
        End If
    End Sub

    ' Deleted records remain in the database for reporting and history
    Private Sub DeleteNode()
        Dim N As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)
        If (Not (N Is Nothing)) And MessageBox.Show("Are you sure you want to delete " _
            & ControlChars.CrLf & "the " & N.Tag.ToString() & " " & N.Text & "?", "Confirm Delete", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question) = DialogResult.Yes Then

            ' Establish the connection
            Dim cn As New SqlConnection(connectionString)
            Dim cmd As New SqlCommand
            cmd.Connection = cn
            Try
                cn.Open() ' Open the connection
                ' If Manager is "deleted" then set the Properties and Assemblies to "deleted" as well
                If N.Tag.ToString() = "Manager" Then
                    ' Set the Manager's manDeleted to True
                    cmd.CommandText() = "UPDATE Managers SET manDeleted = 1 WHERE manNo = " _
                        & N.ID
                    cmd.ExecuteNonQuery()
                    cmd.CommandText() = "UPDATE Properties SET propDeleted = 1 WHERE manNo = " _
                        & N.ID
                    cmd.ExecuteNonQuery()
                    cmd.CommandText() = "UPDATE Assemblies SET assemDeleted = 1 " _
                        & "FROM Managers, Properties, Assemblies " _
                        & "WHERE Managers.manNo = Properties.manNo And " _
                        & "Properties.propNo = Assemblies.propNo And Managers.manNo = " _
                        & N.ID
                    cmd.ExecuteNonQuery()
                End If
                ' If the Property is "deleted" then set the Assemlies to "deleted" as well
                If N.Tag.ToString() = "Property" Then
                    cmd.CommandText() = "UPDATE Properties SET propDeleted = 1 WHERE propNo = " _
                        & CType(treeViewRef.SelectedNode, myTreeNode).ID
                    cmd.ExecuteNonQuery()
                    cmd.CommandText() = "UPDATE Assemblies SET assemDeleted = 1 WHERE propNo = " _
                        & CType(treeViewRef.SelectedNode, myTreeNode).ID
                    cmd.ExecuteNonQuery()
                    If Not TheCutNode Is Nothing Then
                        If N.ID = TheCutNode.ID Then
                            TheCutNode = Nothing
                        End If
                    End If
                End If
                ' If the Assembly is "deleted" just mark it as "deleted"
                If N.Tag.ToString() = "Assembly" Then
                    cmd.CommandText() = "UPDATE Assemblies SET assemDeleted = 1 WHERE assemNo = " _
                        & CType(treeViewRef.SelectedNode, myTreeNode).ID
                    cmd.ExecuteNonQuery()
                    If Not TheCutNode Is Nothing Then
                        If N.Tag.ToString() = TheCutNode.Tag.ToString() And N.ID = TheCutNode.ID Then
                            TheCutNode = Nothing
                        End If
                    End If
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

                '     Catch ex As Exception
                ' A generic exception has occured
                '        MessageBox.Show(ex.Message)
            Finally
                ' Close the connection
                cn.Close()
            End Try

            Dim tn As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode) ' reference the selected node
            treeViewRef.Nodes.Remove(tn) ' remove the selected node
            If treeViewRef.Equals(TreeView1) And txtSearch.Text.Trim.Length > 0 Then
                FillSearchTree(txtSearch.Text)
            ElseIf treeViewRef.Equals(TreeView2) Then
                FillTreeView()
            End If
        End If

    End Sub

    Private Sub LookUp()
        txtSearch.Focus()
    End Sub

    Private Sub CutNode()
        ' Save the currently selected node to the holder, CutNode.
        TheCutNode = CType(treeViewRef.SelectedNode, myTreeNode)
        mnuCut.Enabled = False
        tbbCut.Enabled = False
    End Sub

    Private Sub PasteNode()
        Dim N As myTreeNode
        N = CType(treeViewRef.SelectedNode, myTreeNode)
        If N.Tag.ToString() = "Manager" And TheCutNode.Tag.ToString() = "Property" Then
            Dim cn As New SqlConnection(connectionString)
            Try
                Dim sql As String = "SELECT manNo FROM Properties WHERE propNo = " & TheCutNode.ID
                Dim cmd As New SqlCommand(sql, cn)
                cn.Open()
                intPrevMan = CInt(cmd.ExecuteScalar())
                cmd.CommandText = "UPDATE Properties SET manNo = " & N.ID & " WHERE propNo = " & _
                    CType(TheCutNode, myTreeNode).ID
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Properties SET propPrevManNo = " & intPrevMan & " WHERE propNo = " & _
                    TheCutNode.ID
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
                ' Remove Node from previous manager
                TheCutNode.Remove()
                ' Add Node to the new manager
                TheCutNode.ForeColor = Color.Black
                N.Nodes.Add(TheCutNode)
                TheCutNode = Nothing
                tbbPaste.Enabled = False
                mnuPaste.Enabled = False
            End Try
        ElseIf N.Tag.ToString() = "Property" And TheCutNode.Tag.ToString() = "Assembly" Then
            Dim cn As New SqlConnection(connectionString)
            Try
                Dim sql As String = "SELECT propNo FROM Assemblies WHERE assemNo = " & TheCutNode.ID
                Dim cmd As New SqlCommand(sql, cn)
                cn.Open()
                intPrevProp = CInt(cmd.ExecuteScalar())
                cmd.CommandText = "UPDATE Assemblies SET propNo = " & N.ID & " WHERE assemNo = " & _
                    CType(TheCutNode, myTreeNode).ID
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
                ' Remove Node from previous manager
                TheCutNode.Remove()
                ' Add Node to the new property
                TheCutNode.ForeColor = Color.Black
                N.Nodes.Add(TheCutNode)
                TheCutNode = Nothing
                tbbPaste.Enabled = False
                mnuPaste.Enabled = False
            End Try
            If treeViewRef.Equals(TreeView1) And txtSearch.Text.Trim.Length > 0 Then
                FillSearchTree(txtSearch.Text)
            ElseIf treeViewRef.Equals(TreeView2) Then
                FillTreeView()
            End If
        End If
    End Sub

    Private Sub cmTree_Popup(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmTree.Popup
        Dim N As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)


        cmTree.MenuItems.Clear()
        If N.Tag.ToString() = "Manager" Then
            cmTree.MenuItems.Add("Edit", AddressOf mnuEdit_Click).Enabled = mnuEdit.Enabled
            cmTree.MenuItems.Add("Add Property", AddressOf mnuAddProperty_Click).Enabled = mnuAddProperty.Enabled
            cmTree.MenuItems.Add("Delete", AddressOf mnuDelete_Click).Enabled = mnuDelete.Enabled
            cmTree.MenuItems.Add("Paste", AddressOf mnuPaste_Click).Enabled = mnuPaste.Enabled
        ElseIf N.Tag.ToString() = "Property" Then
            cmTree.MenuItems.Add("Edit", AddressOf mnuEdit_Click).Enabled = mnuEdit.Enabled
            cmTree.MenuItems.Add("Add Assembly", AddressOf mnuAddAssembly_Click).Enabled = mnuAddAssembly.Enabled
            cmTree.MenuItems.Add("Delete", AddressOf mnuDelete_Click).Enabled = mnuDelete.Enabled
            cmTree.MenuItems.Add("Cut", AddressOf mnuCut_Click).Enabled = mnuCut.Enabled
            cmTree.MenuItems.Add("Paste", AddressOf mnuPaste_Click).Enabled = mnuPaste.Enabled
        ElseIf N.Tag.ToString() = "Assembly" Then
            cmTree.MenuItems.Add("Edit", AddressOf mnuEdit_Click).Enabled = mnuEdit.Enabled
            cmTree.MenuItems.Add("Delete", AddressOf mnuDelete_Click).Enabled = mnuDelete.Enabled
            cmTree.MenuItems.Add("Cut", AddressOf mnuCut_Click).Enabled = mnuCut.Enabled
        End If



    End Sub

    Private Sub cmTests_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmTests.Popup
        cmTests.MenuItems.Clear()

        cmTests.MenuItems.Add("Delete Test", AddressOf mnuDeleteTest_Click).Enabled = mnuDeleteTest.Enabled

    End Sub


    Private Sub mnuAddManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddManager.Click
        NewManager()
    End Sub

    Private Sub mnuAddProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddProperty.Click
        NewProperty()
    End Sub

    Private Sub mnuAddAssembly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddAssembly.Click
        NewAssembly()
    End Sub

    Private Sub mnuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEdit.Click
        Edit()
    End Sub

    Private Sub mnuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelete.Click
        DeleteNode()
    End Sub

    Private Sub mnuFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFind.Click
        LookUp()
    End Sub

    Private Sub mnuCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCut.Click
        CutNode()
    End Sub

    Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
        PasteNode()
    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        UpdateTests()
        RefreshTestList()
        If dvTests.Count > 0 Then
            mnuDeleteTest.Enabled = True
        Else
            mnuDeleteTest.Enabled = False
        End If
    End Sub

    Private Sub BindTests(ByVal testDate As Date)
        With grdTests
            .CaptionText = "Total of " & dvTests.Count & " tests on " & testDate
            .DataSource = dvTests
        End With

        ' You must clear out the TableStyles collection before 
        grdTests.TableStyles.Clear()

        Dim grdTableStyle1 As New DataGridTableStyle
        With grdTableStyle1
            ' You must always set the MappingName, even with a DataView that has
            ' only one Table. If not, you will get no errors but the formatting
            ' will not appear. To avoid mistakes, instead of typing the name of 
            ' the table used when creating the DataSet, you can access the 
            ' TableName property.
            .MappingName = dvTests.Table.TableName
        End With

        ' The use of column styles overrides the automatic generation of columns 
        ' for every column in the DataTable. When column style objects are used,
        ' every column you want to display has to have an associate column style 
        ' object.

        Dim grdColStyle1 As New DataGridTextBoxColumn
        With grdColStyle1
            .MappingName = "manName"
            .HeaderText = "Manager"
            .ReadOnly = True
            .Width = 175
        End With

        Dim grdColStyle2 As New DataGridTextBoxColumn
        With grdColStyle2
            .MappingName = "PropertyAll"
            .HeaderText = "Property"
            .ReadOnly = True
            .Width = 300

        End With

        'Dim grdColStyle3 As New DataGridTextBoxColumn
        'With grdColStyle3
        '    .MappingName = "propStrt"
        '    .HeaderText = "Street Address"
        '    .Width = 175
        'End With

        'Dim grdColStyle4 As New DataGridTextBoxColumn
        'With grdColStyle4
        '    .MappingName = "propCity"
        '    .HeaderText = "City"
        '    .Width = 125
        'End With

        Dim grdColStyle5 As New DataGridTextBoxColumn
        With grdColStyle5
            .MappingName = "assemSerial"
            .HeaderText = "Serial No."
            .Width = 70
        End With

        '   Dim grdColStyle6 As New DataGridTextBoxColumn
        '   With grdColStyle6
        '   .MappingName = "testDate"
        '   .HeaderText = "Test Date"
        '   .Width = 70
        '   End With

        Dim grdColStyle7 As New DataGridBoolColumn
        With grdColStyle7
            .MappingName = "testPerformed"
            .HeaderText = "Performed"
            .AllowNull = False
            .Width = 60
        End With

        Dim grdColStyle8 As New DataGridBoolColumn
        With grdColStyle8
            .MappingName = "testPass"
            .HeaderText = "Passed"
            .AllowNull = False
            .Width = 60
        End With

        '   Dim grdColStyle9 As New DataGridTextBoxColumn
        '   With grdColStyle9
        '   .MappingName = "testHours"
        '   .HeaderText = "Hours"
        '   .Width = 40
        '   End With

        Dim grdColStyle10 As New DataGridComboBoxColumn
        With grdColStyle10
            .MappingName = "tstrName"
            .HeaderText = "Tester"
            .Width = 100
            .ColumnComboBox.DataSource = ds.Tables("Testers").DefaultView
            .ColumnComboBox.DisplayMember = "tstrName"
            .ColumnComboBox.ValueMember = "tstrName"
        End With

        Dim grdColStyle11 As New DataGridTextBoxColumn
        With grdColStyle11
            .MappingName = "PONo"
            .HeaderText = "P.O. No."
            .Width = 80
        End With

        Dim grdColStyle12 As New DataGridTextBoxColumn
        With grdColStyle12
            .MappingName = "Notes"
            .HeaderText = "Notes"
            .Width = 400
        End With

        Dim grdColStyle13 As New DataGridBoolColumn
        With grdColStyle13
            .MappingName = "post"
            .HeaderText = "Post?"
            .AllowNull = False
            .Width = 60
        End With

        Dim grdColStyle14 As New DataGridBoolColumn
        With grdColStyle14
            .MappingName = "test_posted"
            .HeaderText = "Was Posted"
            .AllowNull = False
            .ReadOnly = True
            .Width = 80
        End With

        Dim grdColStyle15 As New DataGridTextBoxColumn
        With grdColStyle15
            .MappingName = "permitNo"
            .HeaderText = "Permit No"
            .Width = 80
        End With

        ' Add the column style objects to the table style's collection of 
        ' column styles. Without this the styles do not take effect.        
        grdTableStyle1.GridColumnStyles.AddRange _
            (New DataGridColumnStyle() {grdColStyle1, grdColStyle2, _
                grdColStyle5, _
                grdColStyle7, grdColStyle8, grdColStyle13, grdColStyle14, grdColStyle10, _
                grdColStyle11, grdColStyle15, grdColStyle12})

        grdTests.TableStyles.Add(grdTableStyle1)

        Dim bn As Binding
        For Each bn In grdTests.DataBindings
            Trace.Write(bn.Control.ToString & ", " & bn.PropertyName)
            Trace.Write(vbCrLf)
        Next
    End Sub

    Private Sub RefillMunicipalityDT()
        Dim cn As New SqlConnection(connectionString)
        Dim sqlStatement As String = "SELECT munNo, munName FROM Municipalities ORDER BY munName"
        Dim cmd As New SqlCommand(sqlStatement, cn)
        Dim sda As New SqlDataAdapter(cmd)
        MainForm.ds.Tables("Municip").Clear()

        ' Display a wait cursor while the TreeNodes are being created.
        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        sda.Fill(ds, "Municip") ' Fill the Municipalities datatable

        ' Reset the cursor to the default for all controls.
        Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub mnuMunicipalities_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMunicipalities.Click
        Dim frm As New MunicipalityForm
        If frm.ShowDialog() = DialogResult.OK Then
            RefillMunicipalityDT()
        End If
    End Sub

    ' This code makes certain that the proper context menu displays if the user selects the node via right-click.
    Private Sub TreeView1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeView1.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim N As myTreeNode
            N = CType(TreeView1.GetNodeAt(e.X, e.Y), myTreeNode)
            TreeView1.SelectedNode = N
        End If
    End Sub

    ' This code makes certain that the proper context menu displays if the user selects the node via right-click.
    Private Sub TreeView2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeView1.MouseDown, TreeView2.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim N As myTreeNode
            N = CType(TreeView2.GetNodeAt(e.X, e.Y), myTreeNode)
            TreeView2.SelectedNode = N
        End If
    End Sub

    Private Sub PropertyToManager()
        Dim PrevManNo As Integer = -1
        Dim PropNo As Integer = -1
        Dim Name As String = ""
        Dim Street As String = ""
        Dim City As String = ""
        Dim State As String = ""
        Dim Zip As String = ""
        Dim Contact As String = ""
        Dim Phone As String = ""
        Dim Fax As String = ""
        Dim Email As String = ""

        If treeViewRef.SelectedNode.Tag.ToString() = "Property" Then

            Dim cn As New SqlConnection(connectionString)
            Try
                ' Get the data for the new Manager'
                Dim sqlStatement As String = "SELECT * FROM Properties WHERE propNo = " & CType(treeViewRef.SelectedNode, myTreeNode).ID
                Dim cmd As New SqlCommand(sqlStatement, cn)
                cn.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                If dr.Read() Then
                    ' Null + Empty = Empty
                    PropNo = CInt(dr.Item("propNo"))
                    Name = dr.Item("propName").ToString() & ""
                    Street = dr.Item("propStrt").ToString() & ""
                    City = dr.Item("propCity").ToString() & ""
                    State = dr.Item("propState").ToString() & ""
                    Zip = dr.Item("propZip").ToString() & ""
                    Contact = dr.Item("propCon").ToString() & ""
                    Phone = dr.Item("propPhone").ToString() & ""
                    Fax = dr.Item("propFax").ToString() & ""
                    Email = dr.Item("propEmail").ToString() & ""
                    PrevManNo = CInt(dr.Item("manNo"))

                End If
                dr.Close()
                cn.Open()

                Dim manKeyHolder As New ManagerKeyHolder(0)
                Dim manager As New Manager(manKeyHolder)
                ' Set new managers to the first pricing scheme
                manager.PricingScheme = 1
                manager.Name = Name
                manager.StreetAddress = Street
                manager.City = City
                manager.State = State
                manager.ZipCode = Zip
                manager.Suite = ""
                manager.Notes = ""
                manager.CurrentAccount = True
                manager.Contact1Name = Contact
                manager.Contact1Phone = Phone
                manager.Contact1Fax = Fax
                manager.Contact1EmailAddress = Email

                Dim qbCustAdapt As New QuickbooksCustomerAdapter()
                If qbCustAdapt.Update(manager) = True Then
                    Dim manAdapt As New ManagerAdapter()
                    manAdapt.Update(manager)

                    cmd.CommandText = "UPDATE Properties SET propPrevManNo = " & PrevManNo & " WHERE propNo = " & PropNo
                    cmd.ExecuteNonQuery()

                    cmd.CommandText() = "UPDATE Properties SET manNo = " & manager.PrimaryKey _
                        & " WHERE propNo = " & PropNo
                    cmd.ExecuteNonQuery()

                    Dim M As myTreeNode = New myTreeNode(manager.Name, "Manager", manager.PrimaryKey, True, 0, 0)
                    Dim P As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)
                    P.Remove()
                    M.Nodes.Add(P)
                    treeViewRef.Nodes.Add(M)
                End If

            Catch ex As SqlException
                'A SqlException has occured - display details
                Dim i As Integer, msg As String
                For i = 0 To ex.Errors.Count - 1
                    Dim sqlErr As SqlError = ex.Errors(i)
                    msg = "Message = " & sqlErr.Message & ControlChars.CrLf
                    msg &= "Source = " & sqlErr.Source & ControlChars.CrLf
                Next
                MessageBox.Show(msg & "this is it")

            Catch ex As Exception
                ' A generic exception has occured
                MessageBox.Show(ex.Message)
            Finally
                ' Close the connection
                cn.Close()
            End Try
        End If
    End Sub


    Private Sub mnuPropToMan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPropToMan.Click
        If Not ManagerExists() Then
            PropertyToManager()
        End If

    End Sub

    Private Function ManagerExists() As Boolean
        Dim ManagerName As String
        Dim StreetAddress As String
        Dim Matches As Integer
        Dim Booly As Boolean = False

        Dim N As myTreeNode = CType(treeViewRef.SelectedNode, myTreeNode)

        Dim cn As New SqlConnection(connectionString)
        Dim sqlStatement As String = "SELECT propName, propStrt FROM Properties WHERE propNo = " & N.ID
        Dim cmd As New SqlCommand(sqlStatement, cn)
        Try
            cn.Open()
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            If dr.Read() Then
                ManagerName = Replace(dr.Item("propName").ToString(), "'", "''")
                StreetAddress = dr.Item("propStrt").ToString()
            End If
            dr.Close()
            cmd.CommandText = "SELECT COUNT(*) FROM Managers WHERE manName = '" & ManagerName & "' AND " _
                & "manStrtAdd = '" & StreetAddress & "' AND manDeleted = 0"
            Matches = CInt(cmd.ExecuteScalar)
            If Matches > 0 Then
                Booly = True
                MessageBox.Show("The manager " & Replace(ManagerName, "''", "'") & " at " & StreetAddress _
                    & " already exists.", "Manager Already Exists", MessageBoxButtons.OK, _
                    MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
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
        Return Booly
    End Function

    Private Sub llblEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblEmail.LinkClicked
        If Not llblEmail.Text Is Nothing And llblEmail.Text <> "" Then
            Dim strEmail As String
            strEmail = "mailto:" & llblEmail.Text
            strEmail = strEmail.Trim()

            Dim psi As New ProcessStartInfo
            psi.UseShellExecute = True
            psi.FileName = strEmail
            Try
                ' This throws an exception because the final webpage may differ from the one specified
                Process.Start(psi)
            Catch ex As Exception
                ' This may suffice
            End Try
        End If
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    Private Sub mnuPricing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPricing.Click
        Dim frm As New PricingForm
        If frm.ShowDialog() = DialogResult.OK Then

        End If
    End Sub

    Private Sub UpdateTests()

        logger.Debug("UpdateTests - Entered")

        Dim index As Integer
        Dim lookUp As Integer
        Dim Performed As Boolean
        Dim Passed As Boolean
        Dim Hours As Double
        Dim Tester As String
        Dim PermitNo As String
        Dim PO As String
        Dim Notes As String
        Dim Post As Boolean
        Dim Max As Integer = dtTests.Rows.Count - 1
        Dim msg As String

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn

        Try
            cn.Open() ' Open Connection
            ' this snippet sets the currently edited cells to end edit and therefore saving them
            With grdTests
                .EndEdit(.TableStyles(0).GridColumnStyles(.CurrentCell.ColumnNumber), _
                         .CurrentCell.RowNumber, False)

                Dim cm As CurrencyManager = CType(.BindingContext(.DataSource), CurrencyManager)
                cm.EndCurrentEdit()
                cm.Refresh()
            End With

            For index = 0 To Max
                Console.WriteLine("1: " & grdTests.Item(index, 1).ToString() & "2: " & grdTests.Item(index, 2).ToString() & "3: " & grdTests.Item(index, 3).ToString() & "4: " & grdTests.Item(index, 4).ToString() & "5: " & grdTests.Item(index, 5).ToString() & "6: " & grdTests.Item(index, 6).ToString())
                lookUp = CInt(dvTests.Item(index).Item("testNo"))
                Performed = Boolean.Parse(grdTests.Item(index, 3).ToString())
                Passed = Boolean.Parse(grdTests.Item(index, 4).ToString())
                ' Hours = grdTests.Item(index, 8)
                Tester = grdTests.Item(index, 7).ToString()
                PO = grdTests.Item(index, 8).ToString()
                PermitNo = grdTests.Item(index, 9).ToString()
                Notes = grdTests.Item(index, 10).ToString()
                Post = Boolean.Parse(grdTests.Item(index, 5).ToString())
                ' WasPosted = Boolean.Parse(grdTests.Item(index, 11).ToString())

                cmd.CommandText &= "UPDATE Tests SET testPerformed = " & CInt(Performed) & ", " _
                    & "testPass = " & CInt(Passed) & ", " _
                    & "testHours = " & Hours & ", " _
                    & "tstrName = '" & Tester & "', " _
                    & "PONo = '" & PO & "', " _
                    & "permitNo = '" & PermitNo & "', " _
                    & "Notes = '" & Notes & "', " _
                    & "post = " & CInt(Post) _
                    & " WHERE testNo = " & lookUp
                cmd.ExecuteNonQuery()
            Next
        Catch ex As SqlException
            'A SqlException has occured - display details
            Dim i As Integer
            For i = 0 To ex.Errors.Count - 1
                Dim sqlErr As SqlError = ex.Errors(i)
                msg = "Message = " & sqlErr.Message & ControlChars.CrLf
                msg &= "Source = " & sqlErr.Source & ControlChars.CrLf
            Next
            logger.Error(msg, ex)
            ' MessageBox.Show(msg)

        Catch ex As Exception
            ' A generic exception has occured
            MessageBox.Show(ex.Message)
            logger.Error(ex.Message,ex)
        Finally
            ' Close the connection
            cn.Close()
        End Try
        logger.Debug("UpdateTests - Exit")
    End Sub



    Private Sub mnuDeleteTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteTest.Click
        If dvTests.Count > 0 And MessageBox.Show("Are you sure you want to delete " _
            & ControlChars.CrLf & "the selected test?", "Confirm Delete", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim TestNo As Integer = CInt(dvTests.Item(grdTests.CurrentRowIndex).Item("testNo"))
            DeleteTest(TestNo)
        End If

    End Sub

    Private Sub DeleteTest(ByVal TestNo As Integer)

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn

        Try
            cn.Open() ' Open Connection
            cmd.CommandText = "DELETE FROM Tests WHERE testNo = " & TestNo
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
            RefreshTestList()
        End Try



    End Sub



    Private Sub mnuRetestLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRetestLetters.Click
        Dim frm As New DateChooserForm
        If frm.ShowDialog() = DialogResult.OK Then
            Dim frmRprt As New ReportsForm
            frmRprt.ReportType = "RetestLetters"
            frmRprt.ShowDialog()
        End If
    End Sub

    Private Sub mnuRetestList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRetestList.Click
        Dim frm As New DateChooserForm
        If frm.ShowDialog() = DialogResult.OK Then
            Dim frmRprt As New ReportsForm
            frmRprt.ReportType = "RetestList"
            frmRprt.ShowDialog()
        End If
    End Sub

    Private Sub mnuManList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuManList.Click
        Dim frm As New ReportsForm
        frm.ReportType = "ManagerList"
        frm.ShowDialog()
    End Sub

    Private Sub mnuMunList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMunList.Click
        Dim frm As New ReportsForm
        frm.ReportType = "MunicipalityList"
        frm.ShowDialog()
    End Sub


    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        UpdateTests()
        RefreshTestList()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        FillSearchTree(txtSearch.Text)
    End Sub

    Private Sub mnuTestReports_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestReports.Click
        Dim frm As New DateChooserForm
        frm.Text = "Test Reports"
        If frm.ShowDialog() = DialogResult.OK Then
            Dim frmRprt As New ReportsForm
            frmRprt.ReportType = "TestReports"
            frmRprt.ShowDialog()
        End If
    End Sub


    Private Sub mnuTestsByManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestsByManager.Click
        Dim frm As New ManagerDateChooserForm
        If frm.ShowDialog() = DialogResult.OK Then
            Dim frmRprt As New ReportsForm
            frmRprt.ReportType = "TestsByManager"
            frmRprt.ShowDialog()
        End If
    End Sub

    Private Sub mnuTestsByMunicipality_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestsByMunicipality.Click
        Dim frm As New MunicipalityDateChooserForm
        If frm.ShowDialog() = DialogResult.OK Then
            Dim frmRprt As New ReportsForm
            frmRprt.ReportType = "TestsByMunicipality"
            frmRprt.ShowDialog()
        End If
    End Sub

    Private Sub mnuTopManagers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTopManagers.Click
        Dim frm As New ReportsForm
        frm.ReportType = "TopManagers"
        frm.ShowDialog()
    End Sub

    Private Sub mnuTestComparisonBTWYears_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTestComparisonBTWYears.Click
        Dim frm As New ReportsForm
        frm.ReportType = "TestComparisonBetweenYears"
        frm.ShowDialog()
    End Sub

    Private Sub TreeView2_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView2.AfterSelect
        treeViewRef = TreeView2
        EnableMenuItems()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FillTreeView()
    End Sub

    Private Sub mnuOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptions.Click
        Dim frm As New OptionsForm()
        frm.ShowDialog()
    End Sub

    Private Sub txtSearch_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            FillSearchTree(txtSearch.Text)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub btnPostToAccounting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPostToAccounting.Click
        UpdateTests()
        Dim selectedDate As Date
        selectedDate = Convert.ToDateTime(MonthCalendar1.SelectionStart())
        Dim managers As New Managers()
        Dim testAdapter As New TestAdapter()
        testAdapter.Fill(managers, selectedDate, True, True)
        Dim qbInvoiceAdapt As New QuickbooksInvoiceAdapter()
        If qbInvoiceAdapt.AddInvoices(managers) = True Then
            MarkTestsAsPosted(selectedDate)
        End If

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New AddToQuickbooksForm()
        frm.Show()
    End Sub

    Private Sub MarkTestsAsPosted(ByVal aDate As Date)
        Dim cn As New SqlConnection(connectionString)
        Dim sql As String = "UPDATE Tests SET post = 0, test_posted = 1 WHERE testDate = '" & aDate & "' AND post = 1"
        Dim cmd As New SqlCommand(sql, cn)
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
            MessageBox.Show(msg)
        Catch ex As Exception
            ' A generic exception has occured
            MessageBox.Show(ex.Message)
        Finally
            ' Close the connection
            cn.Close()
            RefreshTestList()
        End Try
    End Sub


    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged

    End Sub

    Private Sub CompanyInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyInfo.Click
        Dim frm As New CompanyInfoForm()
        frm.ShowDialog()
    End Sub

    Private Sub MenuItem8_Click(sender As Object, e As EventArgs) Handles mnuManagerTest.Click
        Dim frm As New ReportsForm
        frm.ReportType = "ManagerTest"
        frm.ShowDialog()
    End Sub
End Class
