using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;
using System.Data.SqlClient;


namespace bftm
{
	public class MainForm : System.Windows.Forms.Form
	{
		'asdfa
		
		#region " Windows Form Designer generated code "
		
		public MainForm()
		{
			
			//This call is required by the Windows Form Designer.
			InitializeComponent();
			
			//Add any initialization after the InitializeComponent() call
			
		}
		
		//Form overrides dispose to clean up the component list.
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (!(components == null))
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		//Required by the Windows Form Designer
		private System.ComponentModel.Container components = null;
		
		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.
		//Do not modify it using the code editor.
		internal System.Windows.Forms.TreeView TreeView1;
		internal System.Windows.Forms.ImageList ImageList1;
		internal System.Windows.Forms.MonthCalendar MonthCalendar1;
		internal System.Windows.Forms.MainMenu MainMenu1;
		internal System.Windows.Forms.MenuItem MenuItem1;
		internal System.Windows.Forms.MenuItem MenuItem2;
		internal System.Windows.Forms.MenuItem MenuItem7;
		internal System.Windows.Forms.MenuItem MenuItem12;
		internal System.Windows.Forms.ToolBar ToolBar1;
		internal System.Windows.Forms.GroupBox grpDetails;
		internal System.Windows.Forms.Button btnMap;
		internal System.Windows.Forms.Label lblImage;
		internal System.Windows.Forms.ToolBarButton tbbAddManager;
		internal System.Windows.Forms.ToolBarButton tbbAddProp;
		internal System.Windows.Forms.ToolBarButton tbbFind;
		internal System.Windows.Forms.ToolBarButton tbbEdit;
		internal System.Windows.Forms.ToolBarButton tbbDelete;
		internal System.Windows.Forms.ToolBarButton tbbCut;
		internal System.Windows.Forms.ToolBarButton tbbPaste;
		internal System.Windows.Forms.ToolBarButton tbbAddAssembly;
		internal System.Windows.Forms.ContextMenu cmTree;
		internal System.Windows.Forms.MenuItem mnuAddManager;
		internal System.Windows.Forms.MenuItem mnuAddProperty;
		internal System.Windows.Forms.MenuItem mnuAddAssembly;
		internal System.Windows.Forms.MenuItem mnuExit;
		internal System.Windows.Forms.MenuItem mnuCut;
		internal System.Windows.Forms.MenuItem mnuPaste;
		internal System.Windows.Forms.MenuItem mnuEdit;
		internal System.Windows.Forms.MenuItem mnuDelete;
		internal System.Windows.Forms.MenuItem mnuFind;
		internal System.Windows.Forms.MenuItem MenuItem3;
		internal System.Windows.Forms.MenuItem mnuMunicipalities;
		internal System.Windows.Forms.MenuItem MenuItem6;
		internal System.Windows.Forms.MenuItem mnuPropToMan;
		internal System.Windows.Forms.ToolBarButton ToolBarButton1;
		internal System.Windows.Forms.ToolBarButton ToolBarButton2;
		internal System.Windows.Forms.ToolBarButton ToolBarButton3;
		internal System.Windows.Forms.ToolBarButton ToolBarButton4;
		internal System.Windows.Forms.DataGrid grdTests;
		internal System.Windows.Forms.LinkLabel llblEmail;
		internal System.Windows.Forms.Button btnPhoneOrTest;
		internal System.Windows.Forms.Label lblRight;
		internal System.Windows.Forms.Label lblLeft;
		internal System.Windows.Forms.MenuItem mnuPricing;
		internal System.Windows.Forms.Label lblBottom;
		internal System.Windows.Forms.ContextMenu cmTests;
		internal System.Windows.Forms.MenuItem MenuItem4;
		internal System.Windows.Forms.MenuItem mnuDeleteTest;
		internal System.Windows.Forms.MenuItem MenuItem5;
		internal System.Windows.Forms.MenuItem mnuManList;
		internal System.Windows.Forms.MenuItem mnuMunList;
		internal System.Windows.Forms.MenuItem mnuRetestList;
		internal System.Windows.Forms.MenuItem mnuRetestLetters;
		internal System.Windows.Forms.ToolBarButton ToolBarButton5;
		internal System.Windows.Forms.ToolBarButton tbbRefresh;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			base.Load += new System.EventHandler(frmMain_Load);
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.TreeView1 = new System.Windows.Forms.TreeView();
			this.TreeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(TreeView1_AfterSelect);
			this.TreeView1.MouseDown += new System.Windows.Forms.MouseEventHandler(TreeView1_MouseDown);
			this.cmTree = new System.Windows.Forms.ContextMenu();
			this.cmTree.Popup += new System.EventHandler(cmTree_Popup);
			this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
			this.MonthCalendar1 = new System.Windows.Forms.MonthCalendar();
			this.MonthCalendar1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(MonthCalendar1_DateChanged);
			this.MainMenu1 = new System.Windows.Forms.MainMenu();
			this.MenuItem1 = new System.Windows.Forms.MenuItem();
			this.MenuItem2 = new System.Windows.Forms.MenuItem();
			this.mnuAddManager = new System.Windows.Forms.MenuItem();
			this.mnuAddManager.Click += new System.EventHandler(mnuAddManager_Click);
			this.mnuAddProperty = new System.Windows.Forms.MenuItem();
			this.mnuAddProperty.Click += new System.EventHandler(mnuAddProperty_Click);
			this.mnuAddAssembly = new System.Windows.Forms.MenuItem();
			this.mnuAddAssembly.Click += new System.EventHandler(mnuAddAssembly_Click);
			this.mnuExit = new System.Windows.Forms.MenuItem();
			this.mnuExit.Click += new System.EventHandler(mnuExit_Click);
			this.MenuItem7 = new System.Windows.Forms.MenuItem();
			this.mnuCut = new System.Windows.Forms.MenuItem();
			this.mnuCut.Click += new System.EventHandler(mnuCut_Click);
			this.mnuPaste = new System.Windows.Forms.MenuItem();
			this.mnuPaste.Click += new System.EventHandler(mnuPaste_Click);
			this.mnuEdit = new System.Windows.Forms.MenuItem();
			this.mnuEdit.Click += new System.EventHandler(mnuEdit_Click);
			this.mnuDelete = new System.Windows.Forms.MenuItem();
			this.mnuDelete.Click += new System.EventHandler(mnuDelete_Click);
			this.MenuItem12 = new System.Windows.Forms.MenuItem();
			this.mnuFind = new System.Windows.Forms.MenuItem();
			this.mnuFind.Click += new System.EventHandler(mnuFind_Click);
			this.MenuItem4 = new System.Windows.Forms.MenuItem();
			this.mnuDeleteTest = new System.Windows.Forms.MenuItem();
			this.mnuDeleteTest.Click += new System.EventHandler(mnuDeleteTest_Click);
			this.MenuItem3 = new System.Windows.Forms.MenuItem();
			this.mnuPricing = new System.Windows.Forms.MenuItem();
			this.mnuPricing.Click += new System.EventHandler(mnuPricing_Click);
			this.mnuMunicipalities = new System.Windows.Forms.MenuItem();
			this.mnuMunicipalities.Click += new System.EventHandler(mnuMunicipalities_Click);
			this.MenuItem6 = new System.Windows.Forms.MenuItem();
			this.mnuPropToMan = new System.Windows.Forms.MenuItem();
			this.mnuPropToMan.Click += new System.EventHandler(mnuPropToMan_Click);
			this.MenuItem5 = new System.Windows.Forms.MenuItem();
			this.mnuManList = new System.Windows.Forms.MenuItem();
			this.mnuManList.Click += new System.EventHandler(mnuManList_Click);
			this.mnuMunList = new System.Windows.Forms.MenuItem();
			this.mnuMunList.Click += new System.EventHandler(mnuMunList_Click);
			this.mnuRetestList = new System.Windows.Forms.MenuItem();
			this.mnuRetestList.Click += new System.EventHandler(mnuRetestList_Click);
			this.mnuRetestLetters = new System.Windows.Forms.MenuItem();
			this.mnuRetestLetters.Click += new System.EventHandler(mnuRetestLetters_Click);
			this.ToolBar1 = new System.Windows.Forms.ToolBar();
			this.ToolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(ToolBar1_ButtonClick);
			this.tbbAddManager = new System.Windows.Forms.ToolBarButton();
			this.tbbAddProp = new System.Windows.Forms.ToolBarButton();
			this.tbbAddAssembly = new System.Windows.Forms.ToolBarButton();
			this.ToolBarButton1 = new System.Windows.Forms.ToolBarButton();
			this.tbbEdit = new System.Windows.Forms.ToolBarButton();
			this.ToolBarButton2 = new System.Windows.Forms.ToolBarButton();
			this.tbbDelete = new System.Windows.Forms.ToolBarButton();
			this.ToolBarButton3 = new System.Windows.Forms.ToolBarButton();
			this.tbbFind = new System.Windows.Forms.ToolBarButton();
			this.ToolBarButton4 = new System.Windows.Forms.ToolBarButton();
			this.tbbCut = new System.Windows.Forms.ToolBarButton();
			this.tbbPaste = new System.Windows.Forms.ToolBarButton();
			this.ToolBarButton5 = new System.Windows.Forms.ToolBarButton();
			this.tbbRefresh = new System.Windows.Forms.ToolBarButton();
			this.grpDetails = new System.Windows.Forms.GroupBox();
			this.lblBottom = new System.Windows.Forms.Label();
			this.llblEmail = new System.Windows.Forms.LinkLabel();
			this.llblEmail.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(llblEmail_LinkClicked);
			this.btnPhoneOrTest = new System.Windows.Forms.Button();
			this.btnPhoneOrTest.Click += new System.EventHandler(btnPhoneORTest_Click);
			this.lblRight = new System.Windows.Forms.Label();
			this.btnMap = new System.Windows.Forms.Button();
			this.btnMap.Click += new System.EventHandler(btnMap_Click);
			this.lblImage = new System.Windows.Forms.Label();
			this.lblLeft = new System.Windows.Forms.Label();
			this.grdTests = new System.Windows.Forms.DataGrid();
			this.cmTests = new System.Windows.Forms.ContextMenu();
			this.cmTests.Popup += new System.EventHandler(cmTests_Popup);
			this.grpDetails.SuspendLayout();
			((System.ComponentModel.ISupportInitialize) this.grdTests).BeginInit();
			this.SuspendLayout();
			//
			//TreeView1
			//
			this.TreeView1.AllowDrop = true;
			this.TreeView1.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left);
			this.TreeView1.ContextMenu = this.cmTree;
			this.TreeView1.Cursor = System.Windows.Forms.Cursors.Default;
			this.TreeView1.HideSelection = false;
			this.TreeView1.ImageList = this.ImageList1;
			this.TreeView1.ItemHeight = 18;
			this.TreeView1.Location = new System.Drawing.Point(0, 33);
			this.TreeView1.Name = "TreeView1";
			this.TreeView1.Size = new System.Drawing.Size(280, 632);
			this.TreeView1.TabIndex = 0;
			//
			//cmTree
			//
			//
			//ImageList1
			//
			this.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit;
			this.ImageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.ImageList1.ImageStream = (System.Windows.Forms.ImageListStreamer) resources.GetObject("ImageList1.ImageStream");
			this.ImageList1.TransparentColor = System.Drawing.Color.Transparent;
			//
			//MonthCalendar1
			//
			this.MonthCalendar1.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.MonthCalendar1.CalendarDimensions = new System.Drawing.Size(3, 1);
			this.MonthCalendar1.ForeColor = System.Drawing.SystemColors.GrayText;
			this.MonthCalendar1.Location = new System.Drawing.Point(360, 509);
			this.MonthCalendar1.MaxSelectionCount = 1;
			this.MonthCalendar1.Name = "MonthCalendar1";
			this.MonthCalendar1.ShowTodayCircle = false;
			this.MonthCalendar1.TabIndex = 6;
			this.MonthCalendar1.TitleBackColor = System.Drawing.Color.LightSlateGray;
			this.MonthCalendar1.TitleForeColor = System.Drawing.Color.White;
			//
			//MainMenu1
			//
			this.MainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { this.MenuItem1, this.MenuItem7, this.MenuItem3, this.MenuItem6, this.MenuItem5 });
			//
			//MenuItem1
			//
			this.MenuItem1.Index = 0;
			this.MenuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { this.MenuItem2, this.mnuExit });
			this.MenuItem1.Text = "F&ile";
			//
			//MenuItem2
			//
			this.MenuItem2.Index = 0;
			this.MenuItem2.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { this.mnuAddManager, this.mnuAddProperty, this.mnuAddAssembly });
			this.MenuItem2.Text = "&New";
			//
			//mnuAddManager
			//
			this.mnuAddManager.Index = 0;
			this.mnuAddManager.Text = "New Manager";
			//
			//mnuAddProperty
			//
			this.mnuAddProperty.Index = 1;
			this.mnuAddProperty.Text = "New Property";
			//
			//mnuAddAssembly
			//
			this.mnuAddAssembly.Index = 2;
			this.mnuAddAssembly.Text = "New Assembly";
			//
			//mnuExit
			//
			this.mnuExit.Index = 1;
			this.mnuExit.Text = "E&xit";
			//
			//MenuItem7
			//
			this.MenuItem7.Index = 1;
			this.MenuItem7.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { this.mnuCut, this.mnuPaste, this.mnuEdit, this.mnuDelete, this.MenuItem12, this.mnuFind, this.MenuItem4, this.mnuDeleteTest });
			this.MenuItem7.Text = "&Edit";
			//
			//mnuCut
			//
			this.mnuCut.Index = 0;
			this.mnuCut.Text = "Cu&t";
			//
			//mnuPaste
			//
			this.mnuPaste.Index = 1;
			this.mnuPaste.Text = "&Paste";
			//
			//mnuEdit
			//
			this.mnuEdit.Index = 2;
			this.mnuEdit.Text = "Edi&t";
			//
			//mnuDelete
			//
			this.mnuDelete.Index = 3;
			this.mnuDelete.Text = "&Delete";
			//
			//MenuItem12
			//
			this.MenuItem12.Index = 4;
			this.MenuItem12.Text = "-";
			//
			//mnuFind
			//
			this.mnuFind.Index = 5;
			this.mnuFind.Text = "&Find";
			//
			//MenuItem4
			//
			this.MenuItem4.Index = 6;
			this.MenuItem4.Text = "-";
			//
			//mnuDeleteTest
			//
			this.mnuDeleteTest.Index = 7;
			this.mnuDeleteTest.Text = "Delete Test";
			//
			//MenuItem3
			//
			this.MenuItem3.Index = 2;
			this.MenuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { this.mnuPricing, this.mnuMunicipalities });
			this.MenuItem3.Text = "Lists";
			//
			//mnuPricing
			//
			this.mnuPricing.Index = 0;
			this.mnuPricing.Text = "Pricing Schemes";
			//
			//mnuMunicipalities
			//
			this.mnuMunicipalities.Index = 1;
			this.mnuMunicipalities.Text = "Municipalities";
			//
			//MenuItem6
			//
			this.MenuItem6.Index = 3;
			this.MenuItem6.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {this.mnuPropToMan});
			this.MenuItem6.Text = "Functions";
			//
			//mnuPropToMan
			//
			this.mnuPropToMan.Index = 0;
			this.mnuPropToMan.Text = "Property to Manager";
			//
			//MenuItem5
			//
			this.MenuItem5.Index = 4;
			this.MenuItem5.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { this.mnuManList, this.mnuMunList, this.mnuRetestList, this.mnuRetestLetters });
			this.MenuItem5.Text = "Reports";
			//
			//mnuManList
			//
			this.mnuManList.Index = 0;
			this.mnuManList.Text = "Manager List";
			//
			//mnuMunList
			//
			this.mnuMunList.Index = 1;
			this.mnuMunList.Text = "Municipality List";
			//
			//mnuRetestList
			//
			this.mnuRetestList.Index = 2;
			this.mnuRetestList.Text = "Retest List";
			//
			//mnuRetestLetters
			//
			this.mnuRetestLetters.Index = 3;
			this.mnuRetestLetters.Text = "Retest Letters";
			//
			//ToolBar1
			//
			this.ToolBar1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.ToolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] { this.tbbAddManager, this.tbbAddProp, this.tbbAddAssembly, this.ToolBarButton1, this.tbbEdit, this.ToolBarButton2, this.tbbDelete, this.ToolBarButton3, this.tbbFind, this.ToolBarButton4, this.tbbCut, this.tbbPaste, this.ToolBarButton5, this.tbbRefresh });
			this.ToolBar1.DropDownArrows = true;
			this.ToolBar1.ImageList = this.ImageList1;
			this.ToolBar1.Location = new System.Drawing.Point(0, 0);
			this.ToolBar1.Name = "ToolBar1";
			this.ToolBar1.ShowToolTips = true;
			this.ToolBar1.Size = new System.Drawing.Size(1012, 29);
			this.ToolBar1.TabIndex = 7;
			//
			//tbbAddManager
			//
			this.tbbAddManager.ImageIndex = 0;
			this.tbbAddManager.Tag = "tester";
			this.tbbAddManager.ToolTipText = "New Manager";
			//
			//tbbAddProp
			//
			this.tbbAddProp.ImageIndex = 1;
			this.tbbAddProp.ToolTipText = "New Property";
			//
			//tbbAddAssembly
			//
			this.tbbAddAssembly.ImageIndex = 2;
			this.tbbAddAssembly.ToolTipText = "New Assembly";
			//
			//ToolBarButton1
			//
			this.ToolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			//
			//tbbEdit
			//
			this.tbbEdit.ImageIndex = 5;
			this.tbbEdit.ToolTipText = "Edit";
			//
			//ToolBarButton2
			//
			this.ToolBarButton2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			//
			//tbbDelete
			//
			this.tbbDelete.ImageIndex = 4;
			this.tbbDelete.ToolTipText = "Delete";
			//
			//ToolBarButton3
			//
			this.ToolBarButton3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			//
			//tbbFind
			//
			this.tbbFind.ImageIndex = 3;
			this.tbbFind.ToolTipText = "Find";
			//
			//ToolBarButton4
			//
			this.ToolBarButton4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			//
			//tbbCut
			//
			this.tbbCut.ImageIndex = 11;
			this.tbbCut.ToolTipText = "Cut";
			//
			//tbbPaste
			//
			this.tbbPaste.ImageIndex = 12;
			this.tbbPaste.ToolTipText = "Paste";
			//
			//ToolBarButton5
			//
			this.ToolBarButton5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			//
			//tbbRefresh
			//
			this.tbbRefresh.ImageIndex = 13;
			this.tbbRefresh.ToolTipText = "Refresh TreeView";
			//
			//grpDetails
			//
			this.grpDetails.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right);
			this.grpDetails.Controls.Add(this.lblBottom);
			this.grpDetails.Controls.Add(this.llblEmail);
			this.grpDetails.Controls.Add(this.btnPhoneOrTest);
			this.grpDetails.Controls.Add(this.lblRight);
			this.grpDetails.Controls.Add(this.btnMap);
			this.grpDetails.Controls.Add(this.lblImage);
			this.grpDetails.Controls.Add(this.lblLeft);
			this.grpDetails.Location = new System.Drawing.Point(287, 32);
			this.grpDetails.Name = "grpDetails";
			this.grpDetails.Size = new System.Drawing.Size(724, 104);
			this.grpDetails.TabIndex = 11;
			this.grpDetails.TabStop = false;
			//
			//lblBottom
			//
			this.lblBottom.Location = new System.Drawing.Point(32, 80);
			this.lblBottom.Name = "lblBottom";
			this.lblBottom.Size = new System.Drawing.Size(608, 16);
			this.lblBottom.TabIndex = 22;
			//
			//llblEmail
			//
			this.llblEmail.Location = new System.Drawing.Point(376, 64);
			this.llblEmail.Name = "llblEmail";
			this.llblEmail.Size = new System.Drawing.Size(248, 16);
			this.llblEmail.TabIndex = 21;
			//
			//btnPhoneOrTest
			//
			this.btnPhoneOrTest.Location = new System.Drawing.Point(568, 29);
			this.btnPhoneOrTest.Name = "btnPhoneOrTest";
			this.btnPhoneOrTest.Size = new System.Drawing.Size(72, 32);
			this.btnPhoneOrTest.TabIndex = 17;
			this.btnPhoneOrTest.Text = "Phonebook";
			//
			//lblRight
			//
			this.lblRight.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
			this.lblRight.Location = new System.Drawing.Point(360, 16);
			this.lblRight.Name = "lblRight";
			this.lblRight.Size = new System.Drawing.Size(200, 48);
			this.lblRight.TabIndex = 16;
			//
			//btnMap
			//
			this.btnMap.Location = new System.Drawing.Point(240, 29);
			this.btnMap.Name = "btnMap";
			this.btnMap.Size = new System.Drawing.Size(64, 32);
			this.btnMap.TabIndex = 15;
			this.btnMap.Text = "Map";
			//
			//lblImage
			//
			this.lblImage.ImageAlign = System.Drawing.ContentAlignment.TopRight;
			this.lblImage.ImageList = this.ImageList1;
			this.lblImage.Location = new System.Drawing.Point(8, 16);
			this.lblImage.Name = "lblImage";
			this.lblImage.Size = new System.Drawing.Size(24, 56);
			this.lblImage.TabIndex = 14;
			//
			//lblLeft
			//
			this.lblLeft.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
			this.lblLeft.Location = new System.Drawing.Point(32, 16);
			this.lblLeft.Name = "lblLeft";
			this.lblLeft.Size = new System.Drawing.Size(200, 56);
			this.lblLeft.TabIndex = 13;
			//
			//grdTests
			//
			this.grdTests.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right);
			this.grdTests.CaptionBackColor = System.Drawing.Color.LightSlateGray;
			this.grdTests.CaptionForeColor = System.Drawing.Color.White;
			this.grdTests.ContextMenu = this.cmTests;
			this.grdTests.DataMember = "";
			this.grdTests.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.grdTests.Location = new System.Drawing.Point(288, 144);
			this.grdTests.Name = "grdTests";
			this.grdTests.Size = new System.Drawing.Size(724, 360);
			this.grdTests.TabIndex = 12;
			//
			//cmTests
			//
			//
			//frmMain
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1012, 669);
			this.ContextMenu = this.cmTests;
			this.Controls.Add(this.grdTests);
			this.Controls.Add(this.grpDetails);
			this.Controls.Add(this.ToolBar1);
			this.Controls.Add(this.MonthCalendar1);
			this.Controls.Add(this.TreeView1);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.Menu = this.MainMenu1;
			this.Name = "frmMain";
			this.Text = "Backflow Testing Manager";
			this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
			this.grpDetails.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize) this.grdTests).EndInit();
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		public static DataSet ds;
		private DataTable dtTests;
		private DataView dvTests;
		private string strLabelLeft;
		private string strLabelRight;
		private string strLabelBottom;
		private myTreeNode TheCutNode = null;
		private int intPrevMan;
		private int intPrevProp;
		
		private string ComName;
		private string Street;
		private string City;
		private string State;
		private string Zip;
		private string Contact;
		private string Phone;
		private string Fax;
		private string Email;
		
		
		// Handle the Form load event.
		private void frmMain_Load(object sender, System.EventArgs e)
		{
			// Fill the DataTables for ComboBoxes
			FillDataTables();
			// Fill the TreeView
			FillTreeView();
			RefreshTestList();
			this.WindowState = FormWindowState.Maximized;
			if (dvTests.Count > 0)
			{
				mnuDeleteTest.Enabled = true;
			}
			else
			{
				mnuDeleteTest.Enabled = false;
			}
			mnuPaste.Enabled = false;
			tbbPaste.Enabled = false;
			
		}
		
		private void FillDataTables()
		{
			// Instantiate the dataset
			ds = new DataSet(); // Create a new dataset
			// Establish connection object
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			SqlDataAdapter sda = new SqlDataAdapter(cmd);
			
			// Fill the tables for the comboboxes
			
			// Reset the CommandText to get the Assembly Sizes
			cmd.CommandText = "SELECT AssemSize, AssemSizeNo FROM AssemblySizes";
			sda.Fill(ds, "AssemSizes"); // Fill Assembly Sizes datatable
			// Reset the CommandText to get the Municipality
			cmd.CommandText = "SELECT munNo, munName FROM Municipalities ORDER BY munName";
			sda.Fill(ds, "Municip"); // Fill the Municipalities datatable
			cmd.CommandText = "SELECT tstrName, tstrNo FROM Testers";
			sda.Fill(ds, "Testers"); // Fill the Testers datatable.
			
		}
		
		private void FillTreeView()
		{
			// Create a New SQL Connection
			SqlConnection cn = new SqlConnection(connectionString);
			// Query which returns a resultset like the following
			// Managers No., Manager's Name, Property No., Property's Name, Assembly No., Assembly's Serial No.
			string strQuery = "SELECT Managers.manNo, Managers.manName, Managers.manDeleted, Properties.propNo, " + "Properties.propName, Properties.propDeleted, Properties.storeNo, Assemblies.assemNo, Assemblies.assemSerial, Assemblies.assemDeleted " + "FROM Managers LEFT JOIN Properties ON Managers.manNo = Properties.manNo " + "LEFT JOIN Assemblies ON Properties.propNo = Assemblies.propNo " + "WHERE manDeleted = 0 " + "ORDER BY Managers.manName, Properties.propName, Properties.storeNo";
			
			SqlCommand cmd = new SqlCommand(strQuery, cn);
			
			try
			{
				cn.Open(); // Open connection
				
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				// Display a wait cursor while the TreeNodes are being created.
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				
				// Suppress repainting the TreeView until all the objects have been created.
				TreeView1.BeginUpdate();
				
				// Clear the TreeView each time the method is called.
				TreeView1.Nodes.Clear();
				
				// Variables needed to iterate through the resultset
				int currentManager;
				int lastManager = 0;
				int currentProperty;
				int lastProperty = 0;
				int currentAssembly;
				int lastAssembly = 0;
				//				bool propDeleted;
				//				bool assemDeleted;
				//				string currentPropertyName;
				//				string PropertyName;
				string storeNo;
				
				// Ititial position values for populating the TreeView
				int j = - 1;
				int k = - 1;
				
				while (dr.Read())
				{
					currentManager = System.Convert.ToInt32(dr["manNo"]);
					if (currentManager != lastManager)
					{
						TreeView1.Nodes.Add(new myTreeNode(dr["manName"].ToString(), "Manager", currentManager, true, 0, 0));
						lastManager = currentManager;
						j++; // Increment the Manager Index
						k = - 1; // Reset the Property Index once a new manager is added
					}
					if (! Information.IsDBNull(dr["propNo"]))
					{
						currentProperty = System.Convert.ToInt32(dr["propNo"]);
						if (currentProperty != lastProperty && bool.Parse(dr["propDeleted"].ToString()) == false)
						{
							if (Information.IsDBNull(dr["storeNo"]))
							{
								storeNo = "";
							}
							else if (dr["storeNo"].ToString() == "")
							{
								storeNo = "";
							}
							else
							{
								storeNo = " - " + dr["storeNo"].ToString();
							}
							// This is around where I have to look into
							TreeView1.Nodes[j].Nodes.Add(new myTreeNode(dr["propName"].ToString() + storeNo, "Property", currentProperty, true, 1, 1));
							lastProperty = currentProperty;
							k++; // Increment the Property Index
						}
					}
					if (! Information.IsDBNull(dr["assemNo"]))
					{
						currentAssembly = System.Convert.ToInt32(dr["assemNo"]);
						if (currentAssembly != lastAssembly && bool.Parse(dr["assemDeleted"].ToString()) == false)
						{
							TreeView1.Nodes[j].Nodes[k].Nodes.Add(new myTreeNode(dr["assemSerial"].ToString() + "", "Assembly", currentAssembly, true, 2, 2));
						}
					}
				}
				dr.Close();
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			
			// Reset the cursor to the default for all controls.
			Cursor.Current = System.Windows.Forms.Cursors.Default;
			
			// Begin repainting the TreeView.
			TreeView1.EndUpdate();
			
		} //FillTreeView
		
		// Toolbar buttons are enabled or disabled based on the Node selected in the Treeview
		// in addition must update the labels after after a node is selected by calling FillLabels
		private void TreeView1_AfterSelect(System.Object sender, System.Windows.Forms.TreeViewEventArgs e)
		{
			
			object temp_object = (myTreeNode) TreeView1.SelectedNode;
			bftm.myTreeNode temp_bftmmyTreeNode =  (bftm.myTreeNode) temp_object;
			FillLabels(ref temp_bftmmyTreeNode);
			if (TreeView1.SelectedNode != null)
			{
				myTreeNode N = (myTreeNode) TreeView1.SelectedNode;
				if (N.Tag.ToString() == "Manager")
				{
					mnuAddProperty.Enabled = true;
					mnuAddAssembly.Enabled = false;
					mnuCut.Enabled = false;
					if (TheCutNode != null)
					{
						if (TheCutNode.Tag.ToString() == "Property")
						{
							mnuPaste.Enabled = true;
							tbbPaste.Enabled = true;
						}
						else
						{
							mnuPaste.Enabled = false;
							tbbPaste.Enabled = false;
						}
					}
					mnuPropToMan.Enabled = false;
					tbbAddProp.Enabled = true;
					tbbAddAssembly.Enabled = false;
					tbbCut.Enabled = false;
				}
				if (N.Tag.ToString() == "Property")
				{
					mnuAddProperty.Enabled = false;
					mnuAddAssembly.Enabled = true;
					mnuPropToMan.Enabled = true;
					
					if (TheCutNode != null)
					{
						if (TheCutNode.Tag.ToString() == "Assembly")
						{
							mnuPaste.Enabled = true;
							tbbPaste.Enabled = true;
						}
						else
						{
							mnuPaste.Enabled = false;
							tbbPaste.Enabled = false;
						}
						mnuCut.Enabled = false;
						tbbCut.Enabled = false;
					}
					else
					{
						mnuCut.Enabled = true;
						tbbCut.Enabled = true;
					}
					tbbAddProp.Enabled = false;
					tbbAddAssembly.Enabled = true;
				}
				if (N.Tag.ToString() == "Assembly")
				{
					mnuAddProperty.Enabled = false;
					mnuAddAssembly.Enabled = false;
					mnuPropToMan.Enabled = false;
					if (TheCutNode == null)
					{
						mnuCut.Enabled = true;
						tbbCut.Enabled = true;
					}
					else
					{
						mnuCut.Enabled = false;
						tbbCut.Enabled = false;
					}
					mnuPaste.Enabled = false;
					tbbAddProp.Enabled = false;
					tbbAddAssembly.Enabled = false;
					tbbPaste.Enabled = false;
				}
			}
		}
		
		
		// Fill lblAddress and lblContact based on the node selected in the TreeView
		private void FillLabels(ref myTreeNode N)
		{
			
			// If node is a Manager display the Manager's information
			if (N.Tag.ToString() == "Manager")
			{
				DisplayManagerInfo(N);
				// Else if Property display the Property's information
			}
			else if (N.Tag.ToString() == "Property")
			{
				DisplayPropertyInfo(N);
				// Else if Assembly display the Assembly's information
			}
			else if (N.Tag.ToString() == "Assembly")
			{
				DisplayAssemblyInfo(N);
			}
			
			lblLeft.Text = strLabelLeft;
			lblRight.Text = strLabelRight;
			lblBottom.Text = strLabelBottom;
			llblEmail.Text = Email;
			grpDetails.Text = N.Tag.ToString();
		}
		
		private void DisplayPropertyInfo(myTreeNode N)
		{
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			
			// Displays the image next to the label information
			lblImage.ImageIndex = 1;
			
			// Displays the buttons in the label area
			btnMap.Visible = true;
			btnPhoneOrTest.Text = "Phonebook";
			
			try
			{
				cn.Open(); // Open connection
				// Get the Property Information based on the Type and ID provided in the Selected Node (myTreeNode)
				cmd.CommandText = "SELECT propName, propStrt, propCity, propState, propZip, propCon, propPhone, " + "propFax, propEmail FROM Properties Where propNo = " + N.ID;
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				if (dr.Read())
				{
					// Null + Empty = Empty
					ComName = dr["propName"].ToString() + "";
					Street = dr["propStrt"].ToString() + "";
					City = dr["propCity"].ToString() + "";
					State = dr["propState"].ToString() + "";
					Zip = dr["propZip"].ToString() + "";
					Contact = dr["propCon"].ToString() + "";
					Phone = dr["propPhone"].ToString() + "";
					Fax = dr["propFax"].ToString() + "";
					Email = dr["propEmail"].ToString() + "";
				}
				
				dr.Close(); // Close DataReader
				
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				cn.Close(); // Close the connection
			}
			
			// Set the Property's information to the following strings
			strLabelLeft = ComName + ControlChars.CrLf + Street + ControlChars.CrLf + City + ", " + State + " " + Zip;
			strLabelRight = "Contact: " + Contact + ControlChars.CrLf + "Phone: " + Phone + ControlChars.CrLf + "Fax: " + Fax;
			strLabelBottom = "";
			
		}
		
		private void DisplayAssemblyInfo(myTreeNode N)
		{
			
			string Model = "";
			string Serial = "";
			string Usage = "";
			string InstallDate;
			string Location = "";
			string Type = "";
			string Manufacturer = "";
			string Size = "";
			string LastTest = "";
			
			// If the Selected Node is an Assembly then get the Assembly's info
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			
			lblImage.ImageIndex = 2;
			btnMap.Visible = false;
			btnPhoneOrTest.Text = "Add Test";
			
			try
			{
				cn.Open(); // Open connection
				cmd.CommandText = "SELECT Assemblies.assemMod, Assemblies.assemSerial, Assemblies.assemUsage, " + "Assemblies.assemInstDt, Assemblies.assemLoc, Assemblies.assemType, " + "Assemblies.AssemMan, AssemblySizes.AssemSize " + "FROM Assemblies " + "LEFT JOIN AssemblySizes ON AssemblySizes.AssemSizeNo = Assemblies.assemSizeNo " + "LEFT JOIN Tests ON Tests.assemNo = Assemblies.assemNo " + "WHERE Assemblies.assemNo = " + N.ID;
				
				SqlDataReader dr = cmd.ExecuteReader();
				
				if (dr.Read())
				{
					// Null + Empty = Empty
					Model = dr["assemMod"].ToString() + "";
					Serial = dr["assemSerial"].ToString() + "";
					Usage = dr["assemUsage"].ToString() + "";
					InstallDate = dr["assemInstDt"].ToString() + "";
					Location = dr["assemLoc"].ToString() + "";
					Type = dr["assemType"].ToString() + "";
					Manufacturer = dr["AssemMan"].ToString() + "";
					Size = dr["AssemSize"].ToString() + "";
				}
				Email = "";
				dr.Close(); // Close the DataReader
				cmd.CommandText = "SELECT MAX(Tests.testDate) FROM Tests " + "LEFT JOIN Assemblies ON Assemblies.assemNo = Tests.assemNo " + "WHERE Assemblies.assemNo = " + N.ID;
				
				if (Information.IsDBNull(cmd.ExecuteScalar()))
				{
					LastTest = null;
				}
				else
				{
					LastTest = cmd.ExecuteScalar().ToString();
				}
				
				
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			
			// Set the Assembly's information to the following strings
			strLabelLeft = "Serial: " + Serial + ControlChars.CrLf + "Type: " + Type + ControlChars.CrLf + "Manufacturer: " + Manufacturer + ControlChars.CrLf + "Model: " + Model;
			strLabelRight = "Size: " + Size + ControlChars.CrLf + "Usage: " + Usage + ControlChars.CrLf + "Last Test Date: " + LastTest + ControlChars.CrLf;
			strLabelBottom = "Location: " + Location;
			
		}
		private void DisplayManagerInfo(myTreeNode N)
		{
			
			// Create SQL Connection
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			
			// Displays the image next to the label information
			lblImage.ImageIndex = 0;
			// Displays the buttons in the label area
			btnMap.Visible = true;
			btnPhoneOrTest.Text = "Phonebook";
			
			try
			{
				// Open the DataConnection
				cn.Open();
				// Select the Manager information where the Manager No. is equal to the Manager No.
				// stored in the myTreeNode of the Selected Node.
				cmd.CommandText = "SELECT manName, manStrtAdd, manCity, manState, manZip, " + "manCntct, manPhone, manFax, manEmail FROM Managers Where manNo = " + ((myTreeNode) TreeView1.SelectedNode).ID;
				
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				//				string str;
				
				if (dr.Read())
				{
					// Null + Empty = Empty
					ComName = dr["manName"].ToString() + "";
					Street = dr["manStrtAdd"].ToString() + "";
					City = dr["manCity"].ToString() + "";
					State = dr["manState"].ToString() + "";
					Zip = dr["manZip"].ToString() + "";
					Contact = dr["manCntct"].ToString() + "";
					Phone = dr["manPhone"].ToString() + "";
					Fax = dr["manFax"].ToString() + "";
					Email = dr["manEmail"].ToString() + "";
				}
				dr.Close();
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			
			// Set the Managers info to the following strings
			strLabelLeft = ComName + ControlChars.CrLf + Street + ControlChars.CrLf + City + ", " + State + " " + Zip;
			strLabelRight = "Contact: " + Contact + ControlChars.CrLf + "Phone: " + Phone + ControlChars.CrLf + "Fax: " + Fax;
			strLabelBottom = "";
		}
		
		
		private void btnMap_Click(System.Object sender, System.EventArgs e)
		{
			// Get info to build mapquest string
			Street = Street.Trim();
			City = City.Trim();
			State = State.Trim();
			// Replace blank spaces in the strings with + for the Query String
			Street = Street.Replace(" ", "+");
			City = City.Replace(" ", "+");
			// Build mapquest string
			string MapQuest = "http://www.mapquest.com/maps/map.adp?country=US&countryid=250" + "&searchtype=address&address=" + Street + "&city=" + City + "&state=" + State.ToUpper() + "&zipcode=" + Zip + "\r\n";
			
			ProcessStartInfo psi = new ProcessStartInfo();
			psi.UseShellExecute = true;
			psi.FileName = MapQuest;
			try
			{
				// This throws an exception because the final webpage may differ from the one specified
				Process.Start(psi);
			}
			catch (Exception)
			{
				// This may suffice
			}
			
		}
		
		private void btnPhoneORTest_Click(System.Object sender, System.EventArgs e)
		{
			if (TreeView1.SelectedNode != null)
			{
				myTreeNode N = (myTreeNode) TreeView1.SelectedNode;
				if (N.Tag.ToString() == "Manager" || N.Tag.ToString() == "Property")
				{
					LookUpInPhonebook();
				}
				else if (N.Tag.ToString() == "Assembly")
				{
					AddTest(N.ID);
				}
			}
		}
		
		private void AddTest(int AssemblyID)
		{
			UpdateTests();
			DateTime TestDate = MonthCalendar1.SelectionStart;
			
			SqlConnection cn = new SqlConnection(connectionString);
			string sqlStatement = "INSERT INTO Tests (assemNo, testDate, Notes, PONo) VALUES " + "(\'" + AssemblyID + "\',\'" + TestDate + "\',\'\', \'\')";
			
			SqlCommand cmd = new SqlCommand(sqlStatement, cn);
			
			try
			{
				cn.Open();
				cmd.ExecuteNonQuery();
				cn.Close();
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			
			RefreshTestList();
			
		}
		
		private void RefreshTestList()
		{
			DateTime chosenDate = MonthCalendar1.SelectionStart;
			
			if (!System.Convert.ToBoolean(ds.Tables["Tests"] == null))
			{
				ds.Tables["Tests"].Clear();
			}
			
			SqlConnection cn = new SqlConnection(connectionString);
			string sqlStatement = "SELECT Managers.manName, Properties.propName, " + "Properties.propStrt, Properties.propCity, Assemblies.assemSerial, " + "Tests.testdate, Tests.testPass, Tests.tstrName, Tests.Notes, " + "Tests.testNo, Tests.testPerformed, Tests.testHours, Tests.PONo " + "FROM Managers " + "INNER JOIN Properties ON Properties.manNo = Managers.manNo " + "INNER JOIN Assemblies ON Assemblies.propNo = Properties.propNo " + "INNER JOIN Tests ON Tests.assemNo = Assemblies.assemNo " + "WHERE Tests.testdate = \'" + chosenDate + "\'";
			SqlCommand cmd = new SqlCommand(sqlStatement, cn);
			SqlDataAdapter sda = new SqlDataAdapter(cmd);
			sda.Fill(ds, "Tests"); // Fill the Tests datatable
			
			dtTests = ds.Tables["Tests"];
			dvTests = dtTests.DefaultView;
			dvTests.AllowNew = false;
			BindTests(chosenDate);
			
		}
		
		private void LookUpInPhonebook()
		{
			// Get info to build mapquest string
			ComName = ComName.Trim();
			City = City.Trim();
			State = State.Trim().ToUpper();
			ComName = ComName.Replace(" ", "+");
			City = City.Replace(" ", "+");
			// Build whitePages string
			
			string WhitePages = "http://www.smartpages.com/co/wp/index.jhtml?From=wp&QueryString=" + ComName + "&QueryType=2&CityZipAC=" + City + "&State=" + State + "&ClearLevel=Cloud9";
			
			ProcessStartInfo psi = new ProcessStartInfo();
			psi.UseShellExecute = true;
			psi.FileName = WhitePages;
			try
			{
				// This throws an exception because the final webpage may differ from the one specified
				Process.Start(psi);
			}
			catch (Exception)
			{
				// This may suffice
			}
			
		}
		
		private void ToolBar1_ButtonClick(System.Object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			
			// Evaluate the Button property to determine which button was clicked.
			switch (ToolBar1.Buttons.IndexOf(e.Button))
			{
				case 0:
					
					NewManager();
					break;
				case 1:
					
					NewProperty();
					break;
				case 2:
					
					NewAssembly();
					break;
				case 3:
					
					break;
					// Spacer
				case 4:
					
					Edit();
					break;
				case 5:
					
					break;
					// Spacer
				case 6:
					
					DeleteNode();
					break;
				case 7:
					
					break;
					// Spacer
				case 8:
					
					LookUp();
					break;
				case 9:
					
					break;
					// Spacer
				case 10:
					
					CutNode();
					break;
				case 11:
					
					PasteNode();
					break;
				case 12:
					
					break;
					// Spacer
				case 13:
					
					FillTreeView();
					break;
			}
			
		}
		
		private void Edit()
		{
			if (TreeView1.SelectedNode != null)
			{
				myTreeNode N = (myTreeNode) TreeView1.SelectedNode;
				if (N.Tag.ToString() == "Manager")
				{
					ManagerDetailForm frm = new ManagerDetailForm();
					frm.txtManNo.Text = N.ID.ToString(); // Use txtManNo to pass the manNo to the form
					if (frm.ShowDialog() == DialogResult.OK)
					{
						TreeView1.SelectedNode.Text = frm.txtManName.Text;
						FillLabels(ref N);
					}
				}
				if (N.Tag.ToString() == "Property")
				{
					PropertyDetailsForm frm = new PropertyDetailsForm();
					frm.txtPropNo.Text = N.ID.ToString(); // use txtPropNo to pass the propNo to the form
					if (frm.ShowDialog() == DialogResult.OK)
					{
						if (frm.txtPropStore.Text == null || frm.txtPropStore.Text.Trim() == "")
						{
							N.Text = frm.txtPropName.Text;
						}
						else
						{
							N.Text = frm.txtPropName.Text + " - " + frm.txtPropStore.Text;
						}
						
						FillLabels(ref N);
					}
				}
				if (N.Tag.ToString() == "Assembly")
				{
					AssemblyDetailForm frm = new AssemblyDetailForm();
					frm.txtAssemNo.Text = N.ID.ToString(); // use txtAssemNo to pass the assemNo to the form
					if (frm.ShowDialog() == DialogResult.OK)
					{
						TreeView1.SelectedNode.Text = frm.txtSerial.Text;
						FillLabels(ref N);
					}
				}
			}
			
		}
		
		private void NewManager()
		{
			ManagerAddForm frm = new ManagerAddForm();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				
				SqlConnection cn = new SqlConnection(connectionString);
				string sql = "SELECT MAX(manNo) FROM Managers";
				SqlCommand cmd = new SqlCommand(sql, cn);
				try
				{
					cn.Open();
					int manNo = System.Convert.ToInt32(cmd.ExecuteScalar());
					TreeView1.Nodes.Add(new myTreeNode(frm.txtManName.Text, "Manager", manNo, true, 0, 0));
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg);
				}
				catch (Exception ex)
				{
					// A generic exception has occured
					MessageBox.Show(ex.Message);
				}
				finally
				{
					// Close the connection
					cn.Close();
				}
			}
		}
		
		private void NewProperty()
		{
			
			PropertyAddForm frm = new PropertyAddForm();
			string strTitle = "New Property for " + ((myTreeNode) TreeView1.SelectedNode).Text;
			int manNo = System.Convert.ToInt32(((myTreeNode) TreeView1.SelectedNode).ID);
			frm.Text = strTitle;
			bftm.PropertyAddForm.manNumber = manNo;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				SqlConnection cn = new SqlConnection(connectionString);
				string sql = "SELECT MAX(propNo) FROM Properties";
				SqlCommand cmd = new SqlCommand(sql, cn);
				try
				{
					cn.Open();
					int propNo = System.Convert.ToInt32(cmd.ExecuteScalar());
					TreeView1.SelectedNode.Nodes.Add(new myTreeNode(frm.txtPropName.Text, "Property", propNo, true, 1, 1));
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg);
				}
				catch (Exception ex)
				{
					// A generic exception has occured
					MessageBox.Show(ex.Message);
				}
				finally
				{
					// Close the connection
					cn.Close();
				}
			}
			
		}
		
		private void NewAssembly()
		{
			string strTitle;
			strTitle = "New Assembly for " + TreeView1.SelectedNode.Text;
			int propNo = System.Convert.ToInt32(((myTreeNode) TreeView1.SelectedNode).ID);
			AssemblyAddForm frm = new AssemblyAddForm();
			frm.Text = strTitle;
			bftm.AssemblyAddForm.propNumber = propNo;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				SqlConnection cn = new SqlConnection(connectionString);
				string sql = "SELECT MAX(assemNo) FROM Assemblies";
				SqlCommand cmd = new SqlCommand(sql, cn);
				try
				{
					cn.Open();
					int assemNo = System.Convert.ToInt32(cmd.ExecuteScalar());
					TreeView1.SelectedNode.Nodes.Add(new myTreeNode(frm.txtSerial.Text, "Assembly", assemNo, true, 2, 2));
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg);
				}
				catch (Exception ex)
				{
					// A generic exception has occured
					MessageBox.Show(ex.Message);
				}
				finally
				{
					// Close the connection
					cn.Close();
				}
			}
		}
		
		// Deleted records remain in the database for reporting and history
		private void DeleteNode()
		{
			myTreeNode N = (myTreeNode) TreeView1.SelectedNode;
			if ((!(N == null)) && MessageBox.Show("Are you sure you want to delete " + ControlChars.CrLf + "the " + N.Tag.ToString() + " " + N.Text + "?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				
				// Establish the connection
				SqlConnection cn = new SqlConnection(connectionString);
				SqlCommand cmd = new SqlCommand();
				cmd.Connection = cn;
				try
				{
					cn.Open(); // Open the connection
					// If Manager is "deleted" then set the Properties and Assemblies to "deleted" as well
					if (N.Tag.ToString() == "Manager")
					{
						// Set the Manager's manDeleted to True
						cmd.CommandText = "UPDATE Managers SET manDeleted = 1 WHERE manNo = " + N.ID;
						cmd.ExecuteNonQuery();
						cmd.CommandText = "UPDATE Properties SET propDeleted = 1 WHERE manNo = " + N.ID;
						cmd.ExecuteNonQuery();
						cmd.CommandText = "UPDATE Assemblies SET assemDeleted = 1 " + "FROM Managers, Properties, Assemblies " + "WHERE Managers.manNo = Properties.manNo And " + "Properties.propNo = Assemblies.propNo And Managers.manNo = " + N.ID;
						cmd.ExecuteNonQuery();
					}
					// If the Property is "deleted" then set the Assemlies to "deleted" as well
					if (N.Tag.ToString() == "Property")
					{
						cmd.CommandText = "UPDATE Properties SET propDeleted = 1 WHERE propNo = " + ((myTreeNode) TreeView1.SelectedNode).ID;
						cmd.ExecuteNonQuery();
						cmd.CommandText = "UPDATE Assemblies SET assemDeleted = 1 WHERE propNo = " + ((myTreeNode) TreeView1.SelectedNode).ID;
						cmd.ExecuteNonQuery();
						if (TheCutNode != null)
						{
							if (N.ID == TheCutNode.ID)
							{
								TheCutNode = null;
							}
						}
					}
					// If the Assembly is "deleted" just mark it as "deleted"
					if (N.Tag.ToString() == "Assembly")
					{
						cmd.CommandText = "UPDATE Assemblies SET assemDeleted = 1 WHERE assemNo = " + ((myTreeNode) TreeView1.SelectedNode).ID;
						cmd.ExecuteNonQuery();
						if (TheCutNode != null)
						{
							if (N.Tag.ToString() == TheCutNode.Tag.ToString() && N.ID == TheCutNode.ID)
							{
								TheCutNode = null;
							}
						}
					}
					
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg);
					
					//     Catch ex As Exception
					// A generic exception has occured
					//        MessageBox.Show(ex.Message)
				}
				finally
				{
					// Close the connection
					cn.Close();
				}
				
				myTreeNode tn = (myTreeNode) TreeView1.SelectedNode; // reference the selected node
				TreeView1.Nodes.Remove(tn); // remove the selected node
				
			}
			
		}
		
		private void LookUp()
		{
			SearchForm frm = new SearchForm();
			this.AddOwnedForm(frm);
			frm.Show();
		}
		
		private void CutNode()
		{
			// Save the currently selected node to the holder, CutNode.
			TheCutNode = (myTreeNode) TreeView1.SelectedNode;
			mnuCut.Enabled = false;
			tbbCut.Enabled = false;
		}
		
		private void PasteNode()
		{
			myTreeNode N;
			N = (myTreeNode) TreeView1.SelectedNode;
			if (N.Tag.ToString() == "Manager" && TheCutNode.Tag.ToString() == "Property")
			{
				SqlConnection cn = new SqlConnection(connectionString);
				try
				{
					string sql = "SELECT manNo FROM Properties WHERE propNo = " + TheCutNode.ID;
					SqlCommand cmd = new SqlCommand(sql, cn);
					cn.Open();
					intPrevMan = System.Convert.ToInt32(cmd.ExecuteScalar());
					cmd.CommandText = "UPDATE Properties SET manNo = " + N.ID + " WHERE propNo = " + TheCutNode.ID;
					cmd.ExecuteNonQuery();
					cmd.CommandText = "UPDATE Properties SET propPrevManNo = " + intPrevMan + " WHERE propNo = " + TheCutNode.ID;
					cmd.ExecuteNonQuery();
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg);
				}
				catch (Exception ex)
				{
					// A generic exception has occured
					MessageBox.Show(ex.Message);
				}
				finally
				{
					// Close the connection
					cn.Close();
					// Remove Node from previous manager
					TheCutNode.Remove();
					// Add Node to the new manager
					TheCutNode.ForeColor = Color.Black;
					N.Nodes.Add(TheCutNode);
					TheCutNode = null;
					tbbPaste.Enabled = false;
					mnuPaste.Enabled = false;
				}
			}
			else if (N.Tag.ToString() == "Property" && TheCutNode.Tag.ToString() == "Assembly")
			{
				SqlConnection cn = new SqlConnection(connectionString);
				try
				{
					string sql = "SELECT propNo FROM Assemblies WHERE assemNo = " + TheCutNode.ID;
					SqlCommand cmd = new SqlCommand(sql, cn);
					cn.Open();
					intPrevProp = System.Convert.ToInt32(cmd.ExecuteScalar());
					cmd.CommandText = "UPDATE Assemblies SET propNo = " + N.ID + " WHERE assemNo = " + TheCutNode.ID;
					cmd.ExecuteNonQuery();
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg);
				}
				catch (Exception ex)
				{
					// A generic exception has occured
					MessageBox.Show(ex.Message);
				}
				finally
				{
					// Close the connection
					cn.Close();
					// Remove Node from previous manager
					TheCutNode.Remove();
					// Add Node to the new property
					TheCutNode.ForeColor = Color.Black;
					N.Nodes.Add(TheCutNode);
					TheCutNode = null;
					tbbPaste.Enabled = false;
					mnuPaste.Enabled = false;
				}
			}
		}
		
		private void cmTree_Popup(System.Object sender, System.EventArgs e)
		{
			myTreeNode N = (myTreeNode) TreeView1.SelectedNode;
			
			
			cmTree.MenuItems.Clear();
			if (N.Tag.ToString() == "Manager")
			{
				cmTree.MenuItems.Add("Edit", new System.EventHandler( mnuEdit_Click)).Enabled = mnuEdit.Enabled;
				cmTree.MenuItems.Add("Add Property", new System.EventHandler( mnuAddProperty_Click)).Enabled = mnuAddProperty.Enabled;
				cmTree.MenuItems.Add("Delete", new System.EventHandler( mnuDelete_Click)).Enabled = mnuDelete.Enabled;
				cmTree.MenuItems.Add("Paste", new System.EventHandler( mnuPaste_Click)).Enabled = mnuPaste.Enabled;
			}
			else if (N.Tag.ToString() == "Property")
			{
				cmTree.MenuItems.Add("Edit", new System.EventHandler( mnuEdit_Click)).Enabled = mnuEdit.Enabled;
				cmTree.MenuItems.Add("Add Assembly", new System.EventHandler( mnuAddAssembly_Click)).Enabled = mnuAddAssembly.Enabled;
				cmTree.MenuItems.Add("Delete", new System.EventHandler( mnuDelete_Click)).Enabled = mnuDelete.Enabled;
				cmTree.MenuItems.Add("Cut", new System.EventHandler( mnuCut_Click)).Enabled = mnuCut.Enabled;
				cmTree.MenuItems.Add("Paste", new System.EventHandler( mnuPaste_Click)).Enabled = mnuPaste.Enabled;
			}
			else if (N.Tag.ToString() == "Assembly")
			{
				cmTree.MenuItems.Add("Edit", new System.EventHandler( mnuEdit_Click)).Enabled = mnuEdit.Enabled;
				cmTree.MenuItems.Add("Delete", new System.EventHandler( mnuDelete_Click)).Enabled = mnuDelete.Enabled;
				cmTree.MenuItems.Add("Cut", new System.EventHandler( mnuCut_Click)).Enabled = mnuCut.Enabled;
			}
			
			
			
		}
		
		private void cmTests_Popup(object sender, System.EventArgs e)
		{
			cmTests.MenuItems.Clear();
			
			cmTests.MenuItems.Add("Delete Test", new System.EventHandler( mnuDeleteTest_Click)).Enabled = mnuDeleteTest.Enabled;
			
		}
		
		
		private void mnuAddManager_Click(System.Object sender, System.EventArgs e)
		{
			NewManager();
		}
		
		private void mnuAddProperty_Click(System.Object sender, System.EventArgs e)
		{
			NewProperty();
		}
		
		private void mnuAddAssembly_Click(System.Object sender, System.EventArgs e)
		{
			NewAssembly();
		}
		
		private void mnuEdit_Click(System.Object sender, System.EventArgs e)
		{
			Edit();
		}
		
		private void mnuDelete_Click(System.Object sender, System.EventArgs e)
		{
			DeleteNode();
		}
		
		private void mnuFind_Click(System.Object sender, System.EventArgs e)
		{
			LookUp();
		}
		
		private void mnuCut_Click(System.Object sender, System.EventArgs e)
		{
			CutNode();
		}
		
		private void mnuPaste_Click(System.Object sender, System.EventArgs e)
		{
			PasteNode();
		}
		
		private void MonthCalendar1_DateChanged(System.Object sender, System.Windows.Forms.DateRangeEventArgs e)
		{
			UpdateTests();
			RefreshTestList();
			if (dvTests.Count > 0)
			{
				mnuDeleteTest.Enabled = true;
			}
			else
			{
				mnuDeleteTest.Enabled = false;
			}
		}
		
		private void BindTests(DateTime testDate)
		{
			grdTests.CaptionText = "Total of " + dvTests.Count + " tests on " + testDate;
			grdTests.DataSource = dvTests;
			
			// You must clear out the TableStyles collection before
			grdTests.TableStyles.Clear();
			
			DataGridTableStyle grdTableStyle1 = new DataGridTableStyle();
			// You must always set the MappingName, even with a DataView that has
			// only one Table. If not, you will get no errors but the formatting
			// will not appear. To avoid mistakes, instead of typing the name of
			// the table used when creating the DataSet, you can access the
			// TableName property.
			grdTableStyle1.MappingName = dvTests.Table.TableName;
			
			// The use of column styles overrides the automatic generation of columns
			// for every column in the DataTable. When column style objects are used,
			// every column you want to display has to have an associate column style
			// object.
			
			DataGridTextBoxColumn grdColStyle1 = new DataGridTextBoxColumn();
			grdColStyle1.MappingName = "manName";
			grdColStyle1.HeaderText = "Manager";
			grdColStyle1.Width = 175;
			
			DataGridTextBoxColumn grdColStyle2 = new DataGridTextBoxColumn();
			grdColStyle2.MappingName = "propName";
			grdColStyle2.HeaderText = "Property";
			grdColStyle2.Width = 175;
			
			DataGridTextBoxColumn grdColStyle3 = new DataGridTextBoxColumn();
			grdColStyle3.MappingName = "propStrt";
			grdColStyle3.HeaderText = "Street Address";
			grdColStyle3.Width = 175;
			
			DataGridTextBoxColumn grdColStyle4 = new DataGridTextBoxColumn();
			grdColStyle4.MappingName = "propCity";
			grdColStyle4.HeaderText = "City";
			grdColStyle4.Width = 125;
			
			DataGridTextBoxColumn grdColStyle5 = new DataGridTextBoxColumn();
			grdColStyle5.MappingName = "assemSerial";
			grdColStyle5.HeaderText = "Serial No.";
			grdColStyle5.Width = 70;
			
			//   Dim grdColStyle6 As New DataGridTextBoxColumn
			//   With grdColStyle6
			//   .MappingName = "testDate"
			//   .HeaderText = "Test Date"
			//   .Width = 70
			//   End With
			
			DataGridBoolColumn grdColStyle7 = new DataGridBoolColumn();
			grdColStyle7.MappingName = "testPerformed";
			grdColStyle7.HeaderText = "Performed";
			grdColStyle7.AllowNull = false;
			grdColStyle7.Width = 60;
			
			DataGridBoolColumn grdColStyle8 = new DataGridBoolColumn();
			grdColStyle8.MappingName = "testPass";
			grdColStyle8.HeaderText = "Passed";
			grdColStyle8.AllowNull = false;
			grdColStyle8.Width = 60;
			
			//   Dim grdColStyle9 As New DataGridTextBoxColumn
			//   With grdColStyle9
			//   .MappingName = "testHours"
			//   .HeaderText = "Hours"
			//   .Width = 40
			//   End With
			
			DataGridComboBoxColumn grdColStyle10 = new DataGridComboBoxColumn();
			grdColStyle10.MappingName = "tstrName";
			grdColStyle10.HeaderText = "Tester";
			grdColStyle10.Width = 100;
			grdColStyle10.ColumnComboBox.DataSource = ds.Tables["Testers"].DefaultView;
			grdColStyle10.ColumnComboBox.DisplayMember = "tstrName";
			grdColStyle10.ColumnComboBox.ValueMember = "tstrName";
			
			DataGridTextBoxColumn grdColStyle11 = new DataGridTextBoxColumn();
			grdColStyle11.MappingName = "PONo";
			grdColStyle11.HeaderText = "P.O. No.";
			grdColStyle11.Width = 80;
			
			DataGridTextBoxColumn grdColStyle12 = new DataGridTextBoxColumn();
			grdColStyle12.MappingName = "Notes";
			grdColStyle12.HeaderText = "Notes";
			grdColStyle12.Width = 400;
			
			// Add the column style objects to the table style's collection of
			// column styles. Without this the styles do not take effect.
			grdTableStyle1.GridColumnStyles.AddRange(new DataGridColumnStyle[] { grdColStyle1, grdColStyle2, grdColStyle3, grdColStyle4, grdColStyle5, grdColStyle7, grdColStyle8, grdColStyle10, grdColStyle11, grdColStyle12 });
			
			grdTests.TableStyles.Add(grdTableStyle1);
			
			Binding bn;
			foreach (Binding tempLoopVar_bn in grdTests.DataBindings)
			{
				bn = tempLoopVar_bn;
				Trace.Write(bn.Control.ToString() + ", " + bn.PropertyName);
				Trace.Write("\r\n");
			}
		}
		
		private void RefillMunicipalityDT()
		{
			SqlConnection cn = new SqlConnection(connectionString);
			string sqlStatement = "SELECT munNo, munName FROM Municipalities ORDER BY munName";
			SqlCommand cmd = new SqlCommand(sqlStatement, cn);
			SqlDataAdapter sda = new SqlDataAdapter(cmd);
			MainForm.ds.Tables["Municip"].Clear();
			
			// Display a wait cursor while the TreeNodes are being created.
			Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
			
			sda.Fill(ds, "Municip"); // Fill the Municipalities datatable
			
			// Reset the cursor to the default for all controls.
			Cursor.Current = System.Windows.Forms.Cursors.Default;
		}
		
		private void mnuMunicipalities_Click(System.Object sender, System.EventArgs e)
		{
			MunicipalityForm frm = new MunicipalityForm();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				RefillMunicipalityDT();
			}
		}
		
		// This code makes certain that the proper context menu displays if the user selects the node via right-click.
		private void TreeView1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				myTreeNode N;
				N = (myTreeNode) TreeView1.GetNodeAt(e.X, e.Y);
				TreeView1.SelectedNode = N;
			}
		}
		
		private void PropertyToManager()
		{
			int ManNo = 0;
			int PropNo = 0;
			string Name = "";
			string Street = "";
			string City = "";
			string State = "";
			string Zip = "";
			string Contact = "";
			string Phone = "";
			string Fax = "";
			string Email = "";
			
			if (TreeView1.SelectedNode.Tag.ToString() == "Property")
			{
				
				SqlConnection cn = new SqlConnection(connectionString);
				try
				{
					string sqlStatement = "SELECT * FROM Properties WHERE propNo = " + ((myTreeNode) TreeView1.SelectedNode).ID;
					SqlCommand cmd = new SqlCommand(sqlStatement, cn);
					cn.Open();
					SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
					
					if (dr.Read())
					{
						// Null + Empty = Empty
						PropNo = System.Convert.ToInt32(dr["propNo"]);
						Name = dr["propName"].ToString() + "";
						Street = dr["propStrt"].ToString() + "";
						City = dr["propCity"].ToString() + "";
						State = dr["propState"].ToString() + "";
						Zip = dr["propZip"].ToString() + "";
						Contact = dr["propCon"].ToString() + "";
						Phone = dr["propPhone"].ToString() + "";
						Fax = dr["propFax"].ToString() + "";
						Email = dr["propEmail"].ToString() + "";
						ManNo = System.Convert.ToInt32(dr["manNo"]);
					}
					dr.Close();
					cn.Open();
					cmd.CommandText = "UPDATE Properties SET propPrevManNo = " + ManNo + " WHERE propNo = " + PropNo;
					cmd.ExecuteNonQuery();
					
					string strSQL;
					strSQL = "INSERT INTO Managers (manName, manStrtAdd, manCity, manState, " + "manZip, manCntct, manPhone, manFax, manEmail, manSchemeNo) VALUES ( ";
					strSQL += "\'" + Name.Replace("\'", "\'\'") + "\', ";
					strSQL += "\'" + Street + "\', ";
					strSQL += "\'" + City + "\', ";
					strSQL += "\'" + State + "\', ";
					strSQL += "\'" + Zip + "\', ";
					strSQL += "\'" + Contact + "\', ";
					strSQL += "\'" + Phone + "\', ";
					strSQL += "\'" + Fax + "\', ";
					strSQL += "\'" + Email + "\', ";
					strSQL += "1)";
					
					cmd.CommandText = strSQL;
					cmd.ExecuteNonQuery();
					
					cmd.CommandText = "SELECT MAX(manNo) FROM Managers";
					int manNumber = System.Convert.ToInt32(cmd.ExecuteScalar());
					
					cmd.CommandText = "UPDATE Properties SET manNo = " + manNumber + " WHERE propNo = " + PropNo;
					cmd.ExecuteNonQuery();
					
					myTreeNode M = new myTreeNode(Name, "Manager", manNumber, true, 0, 0);
					myTreeNode P = (myTreeNode) TreeView1.SelectedNode;
					P.Remove();
					M.Nodes.Add(P);
					TreeView1.Nodes.Add(M);
					
				}
				catch (SqlException ex)
				{
					//A SqlException has occured - display details
					int i;
					string msg = "";
					// 					string msg;
					for (i = 0; i <= ex.Errors.Count - 1; i++)
					{
						SqlError sqlErr = ex.Errors[i];
						msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
						msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
					}
					MessageBox.Show(msg + "this is it");
					
				}
				catch (Exception ex)
				{
					// A generic exception has occured
					MessageBox.Show(ex.Message);
				}
				finally
				{
					// Close the connection
					cn.Close();
				}
			}
		}
		
		
		private void mnuPropToMan_Click(System.Object sender, System.EventArgs e)
		{
			if (! ManagerExists())
			{
				PropertyToManager();
			}
			
		}
		
		private bool ManagerExists()
		{
			string ManagerName = "";
			string StreetAddress = "";
			int Matches;
			bool Booly = false;
			
			myTreeNode N = (myTreeNode) TreeView1.SelectedNode;
			
			SqlConnection cn = new SqlConnection(connectionString);
			string sqlStatement = "SELECT propName, propStrt FROM Properties WHERE propNo = " + N.ID;
			SqlCommand cmd = new SqlCommand(sqlStatement, cn);
			try
			{
				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();
				if (dr.Read())
				{
					ManagerName = Strings.Replace(dr["propName"].ToString(), "\'", "\'\'", 1, -1, 0);
					StreetAddress = dr["propStrt"].ToString();
				}
				dr.Close();
				cmd.CommandText = "SELECT COUNT(*) FROM Managers WHERE manName = \'" + ManagerName + "\' AND " + "manStrtAdd = \'" + StreetAddress + "\' AND manDeleted = 0";
				Matches = System.Convert.ToInt32(cmd.ExecuteScalar());
				if (Matches > 0)
				{
					Booly = true;
					MessageBox.Show("The manager " + ManagerName.Replace("\'\'", "\'") + " at " + StreetAddress + " already exists.", "Manager Already Exists", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
				}
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			return Booly;
		}
		
		private void llblEmail_LinkClicked(System.Object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (llblEmail.Text != null&& llblEmail.Text != "")
			{
				string strEmail;
				strEmail = "mailto:" + llblEmail.Text;
				strEmail = strEmail.Trim(null);
				
				ProcessStartInfo psi = new ProcessStartInfo();
				psi.UseShellExecute = true;
				psi.FileName = strEmail;
				try
				{
					// This throws an exception because the final webpage may differ from the one specified
					Process.Start(psi);
				}
				catch (Exception)
				{
					// This may suffice
				}
			}
		}
		
		private void mnuExit_Click(System.Object sender, System.EventArgs e)
		{
			this.Close();
		}
		
		private void mnuPricing_Click(System.Object sender, System.EventArgs e)
		{
			PricingForm frm = new PricingForm();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				
			}
		}
		
		private void UpdateTests()
		{
			
			int index;
			int lookUp;
			bool Performed;
			bool Passed;
			double Hours = 0;
			string Tester;
			string PO;
			string Notes;
			int Max = dtTests.Rows.Count - 1;
			//			string str;
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			
			try
			{
				cn.Open(); // Open Connection
				// this snippet sets the currently edited cells to end edit and therefore saving them
				grdTests.EndEdit(grdTests.TableStyles[0].GridColumnStyles[grdTests.CurrentCell.ColumnNumber], grdTests.CurrentCell.RowNumber, false);
				
				CurrencyManager cm = (CurrencyManager) grdTests.BindingContext[grdTests.DataSource];
				cm.EndCurrentEdit();
				cm.Refresh();
				
				for (index = 0; index <= Max; index++)
				{
					lookUp = System.Convert.ToInt32(dvTests[index]["testNo"]);
					Performed = bool.Parse(grdTests[index, 5].ToString());
					Passed = bool.Parse(grdTests[index, 6].ToString());
					// Hours = grdTests.Item(index, 8)
					Tester = grdTests[index, 7].ToString();
					PO = grdTests[index, 8].ToString();
					Notes = grdTests[index, 9].ToString();
					
					
					cmd.CommandText += "UPDATE Tests SET testPerformed = " + System.Convert.ToInt32(Performed) + ", " + "testPass = " + System.Convert.ToInt32(Passed) + ", " + "testHours = " + Hours + ", " + "tstrName = \'" + Tester + "\', " + "PONo = \'" + PO + "\', " + "Notes = \'" + Notes + "\'" + " WHERE testNo = " + lookUp;
					cmd.ExecuteNonQuery();
				}
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				// MessageBox.Show(msg)
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			
		}
		
		
		
		private void mnuDeleteTest_Click(System.Object sender, System.EventArgs e)
		{
			if (dvTests.Count > 0 && MessageBox.Show("Are you sure you want to delete " + ControlChars.CrLf + "the selected test?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				int TestNo = System.Convert.ToInt32(dvTests[grdTests.CurrentRowIndex]["testNo"]);
				DeleteTest(TestNo);
			}
			
		}
		
		private void DeleteTest(int TestNo)
		{
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			
			try
			{
				cn.Open(); // Open Connection
				cmd.CommandText = "DELETE FROM Tests WHERE testNo = " + TestNo;
				cmd.ExecuteNonQuery();
				
			}
			catch (SqlException ex)
			{
				//A SqlException has occured - display details
				int i;
				string msg = "";
				// 				string msg;
				for (i = 0; i <= ex.Errors.Count - 1; i++)
				{
					SqlError sqlErr = ex.Errors[i];
					msg = "Message = " + sqlErr.Message + ControlChars.CrLf;
					msg += "Source = " + sqlErr.Source + ControlChars.CrLf;
				}
				MessageBox.Show(msg);
				
			}
			catch (Exception ex)
			{
				// A generic exception has occured
				MessageBox.Show(ex.Message);
			}
			finally
			{
				// Close the connection
				cn.Close();
			}
			
			RefreshTestList();
			
		}
		
		
		
		private void mnuRetestLetters_Click(System.Object sender, System.EventArgs e)
		{
			DateChooserForm frm = new DateChooserForm();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				ReportsForm frmRprt = new ReportsForm();
				frmRprt.ReportType = "RetestLetters";
				frmRprt.ShowDialog();
			}
		}
		
		private void mnuRetestList_Click(System.Object sender, System.EventArgs e)
		{
			DateChooserForm frm = new DateChooserForm();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				ReportsForm frmRprt = new ReportsForm();
				frmRprt.ReportType = "RetestList";
				frmRprt.ShowDialog();
			}
		}
		
		private void mnuManList_Click(System.Object sender, System.EventArgs e)
		{
			ReportsForm frm = new ReportsForm();
			frm.ReportType = "ManagerList";
			frm.ShowDialog();
		}
		
		private void mnuMunList_Click(System.Object sender, System.EventArgs e)
		{
			ReportsForm frm = new ReportsForm();
			frm.ReportType = "MunicipalityList";
			frm.ShowDialog();
		}
		
		
		protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
		{
			UpdateTests();
			RefreshTestList();
		}
	}
	
}
