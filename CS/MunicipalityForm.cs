using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;
using System.Data.SqlClient;
using System.Text.RegularExpressions;


namespace bftm
{
	public class MunicipalityForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public MunicipalityForm()
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
		internal System.Windows.Forms.TextBox txtName;
		internal System.Windows.Forms.TextBox txtDept;
		internal System.Windows.Forms.TextBox txtStreet;
		internal System.Windows.Forms.TextBox txtCity;
		internal System.Windows.Forms.TextBox txtZipCode;
		internal System.Windows.Forms.ComboBox cboPreface;
		internal System.Windows.Forms.ComboBox cboState;
		internal System.Windows.Forms.TextBox txtContact;
		internal System.Windows.Forms.TextBox txtPhone;
		internal System.Windows.Forms.TextBox txtFax;
		internal System.Windows.Forms.TextBox txtEmail;
		internal System.Windows.Forms.Label Label6;
		internal System.Windows.Forms.Button btnAdd;
		internal System.Windows.Forms.Button btnSave;
		internal System.Windows.Forms.Button btnDelete;
		internal System.Windows.Forms.ListView lvMunicipalities;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Label Label7;
		internal System.Windows.Forms.Label Label13;
		internal System.Windows.Forms.Label Label14;
		internal System.Windows.Forms.Label Label15;
		internal System.Windows.Forms.Label Label16;
		internal System.Windows.Forms.Label Label17;
		internal System.Windows.Forms.ErrorProvider ErrorProvider1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MunicipalityForm));
			this.cboPreface = new System.Windows.Forms.ComboBox();
			base.Load += new System.EventHandler(frmMun_Load);
			base.Closing += new System.ComponentModel.CancelEventHandler(frmMun_Closing);
			this.txtName = new System.Windows.Forms.TextBox();
			this.txtDept = new System.Windows.Forms.TextBox();
			this.txtStreet = new System.Windows.Forms.TextBox();
			this.txtCity = new System.Windows.Forms.TextBox();
			this.cboState = new System.Windows.Forms.ComboBox();
			this.txtZipCode = new System.Windows.Forms.TextBox();
			this.txtContact = new System.Windows.Forms.TextBox();
			this.txtPhone = new System.Windows.Forms.TextBox();
			this.txtFax = new System.Windows.Forms.TextBox();
			this.txtEmail = new System.Windows.Forms.TextBox();
			this.Label6 = new System.Windows.Forms.Label();
			this.btnAdd = new System.Windows.Forms.Button();
			this.btnAdd.Click += new System.EventHandler(btnAdd_Click);
			this.btnSave = new System.Windows.Forms.Button();
			this.btnSave.Click += new System.EventHandler(btnSave_Click);
			this.btnDelete = new System.Windows.Forms.Button();
			this.btnDelete.Click += new System.EventHandler(btnDelete_Click);
			this.lvMunicipalities = new System.Windows.Forms.ListView();
			this.lvMunicipalities.SelectedIndexChanged += new System.EventHandler(lvMunicipalities_SelectedIndexChanged);
			this.Label4 = new System.Windows.Forms.Label();
			this.Label3 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.Label1 = new System.Windows.Forms.Label();
			this.Label7 = new System.Windows.Forms.Label();
			this.Label13 = new System.Windows.Forms.Label();
			this.Label14 = new System.Windows.Forms.Label();
			this.Label15 = new System.Windows.Forms.Label();
			this.Label16 = new System.Windows.Forms.Label();
			this.Label17 = new System.Windows.Forms.Label();
			this.ErrorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.SuspendLayout();
			//
			//cboPreface
			//
			this.cboPreface.Items.AddRange(new object[] { "Village Of", "City Of" });
			this.cboPreface.Location = new System.Drawing.Point(96, 150);
			this.cboPreface.Name = "cboPreface";
			this.cboPreface.Size = new System.Drawing.Size(152, 21);
			this.cboPreface.TabIndex = 1;
			//
			//txtName
			//
			this.txtName.Location = new System.Drawing.Point(96, 174);
			this.txtName.Name = "txtName";
			this.txtName.Size = new System.Drawing.Size(280, 20);
			this.txtName.TabIndex = 2;
			this.txtName.Text = "";
			//
			//txtDept
			//
			this.txtDept.Location = new System.Drawing.Point(96, 198);
			this.txtDept.Name = "txtDept";
			this.txtDept.Size = new System.Drawing.Size(280, 20);
			this.txtDept.TabIndex = 3;
			this.txtDept.Text = "";
			//
			//txtStreet
			//
			this.txtStreet.Location = new System.Drawing.Point(96, 222);
			this.txtStreet.Name = "txtStreet";
			this.txtStreet.Size = new System.Drawing.Size(280, 20);
			this.txtStreet.TabIndex = 4;
			this.txtStreet.Text = "";
			//
			//txtCity
			//
			this.txtCity.Location = new System.Drawing.Point(96, 246);
			this.txtCity.Name = "txtCity";
			this.txtCity.Size = new System.Drawing.Size(152, 20);
			this.txtCity.TabIndex = 5;
			this.txtCity.Text = "";
			//
			//cboState
			//
			this.cboState.Items.AddRange(new object[] { "IL", "IN", "IA", "MI" });
			this.cboState.Location = new System.Drawing.Point(320, 246);
			this.cboState.Name = "cboState";
			this.cboState.Size = new System.Drawing.Size(56, 21);
			this.cboState.TabIndex = 6;
			//
			//txtZipCode
			//
			this.txtZipCode.Location = new System.Drawing.Point(96, 270);
			this.txtZipCode.Name = "txtZipCode";
			this.txtZipCode.Size = new System.Drawing.Size(152, 20);
			this.txtZipCode.TabIndex = 7;
			this.txtZipCode.Text = "";
			//
			//txtContact
			//
			this.txtContact.Location = new System.Drawing.Point(96, 294);
			this.txtContact.Name = "txtContact";
			this.txtContact.Size = new System.Drawing.Size(280, 20);
			this.txtContact.TabIndex = 8;
			this.txtContact.Text = "";
			//
			//txtPhone
			//
			this.txtPhone.Location = new System.Drawing.Point(96, 318);
			this.txtPhone.Name = "txtPhone";
			this.txtPhone.Size = new System.Drawing.Size(152, 20);
			this.txtPhone.TabIndex = 9;
			this.txtPhone.Text = "";
			//
			//txtFax
			//
			this.txtFax.Location = new System.Drawing.Point(96, 342);
			this.txtFax.Name = "txtFax";
			this.txtFax.Size = new System.Drawing.Size(152, 20);
			this.txtFax.TabIndex = 10;
			this.txtFax.Text = "";
			//
			//txtEmail
			//
			this.txtEmail.Location = new System.Drawing.Point(96, 366);
			this.txtEmail.Name = "txtEmail";
			this.txtEmail.Size = new System.Drawing.Size(152, 20);
			this.txtEmail.TabIndex = 11;
			this.txtEmail.Text = "";
			//
			//Label6
			//
			this.Label6.Location = new System.Drawing.Point(280, 254);
			this.Label6.Name = "Label6";
			this.Label6.Size = new System.Drawing.Size(32, 16);
			this.Label6.TabIndex = 17;
			this.Label6.Text = "State";
			this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//btnAdd
			//
			this.btnAdd.Location = new System.Drawing.Point(400, 190);
			this.btnAdd.Name = "btnAdd";
			this.btnAdd.Size = new System.Drawing.Size(72, 32);
			this.btnAdd.TabIndex = 23;
			this.btnAdd.Text = "New";
			//
			//btnSave
			//
			this.btnSave.Location = new System.Drawing.Point(400, 238);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(72, 32);
			this.btnSave.TabIndex = 24;
			this.btnSave.Text = "Save";
			//
			//btnDelete
			//
			this.btnDelete.Location = new System.Drawing.Point(400, 286);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(72, 32);
			this.btnDelete.TabIndex = 25;
			this.btnDelete.Text = "Delete";
			//
			//lvMunicipalities
			//
			this.lvMunicipalities.FullRowSelect = true;
			this.lvMunicipalities.Location = new System.Drawing.Point(8, 8);
			this.lvMunicipalities.MultiSelect = false;
			this.lvMunicipalities.Name = "lvMunicipalities";
			this.lvMunicipalities.Size = new System.Drawing.Size(368, 128);
			this.lvMunicipalities.TabIndex = 26;
			this.lvMunicipalities.View = System.Windows.Forms.View.Details;
			//
			//Label4
			//
			this.Label4.Location = new System.Drawing.Point(8, 222);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(80, 16);
			this.Label4.TabIndex = 30;
			this.Label4.Text = "Street Address";
			this.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label3
			//
			this.Label3.Location = new System.Drawing.Point(8, 198);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(80, 16);
			this.Label3.TabIndex = 29;
			this.Label3.Text = "Department";
			this.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label2
			//
			this.Label2.Location = new System.Drawing.Point(8, 174);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(80, 16);
			this.Label2.TabIndex = 28;
			this.Label2.Text = "Name";
			this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(8, 150);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(80, 16);
			this.Label1.TabIndex = 27;
			this.Label1.Text = "Preface";
			this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label7
			//
			this.Label7.Location = new System.Drawing.Point(8, 246);
			this.Label7.Name = "Label7";
			this.Label7.Size = new System.Drawing.Size(80, 16);
			this.Label7.TabIndex = 31;
			this.Label7.Text = "City";
			this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label13
			//
			this.Label13.Location = new System.Drawing.Point(8, 294);
			this.Label13.Name = "Label13";
			this.Label13.Size = new System.Drawing.Size(80, 16);
			this.Label13.TabIndex = 33;
			this.Label13.Text = "Contact";
			this.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label14
			//
			this.Label14.Location = new System.Drawing.Point(8, 318);
			this.Label14.Name = "Label14";
			this.Label14.Size = new System.Drawing.Size(80, 16);
			this.Label14.TabIndex = 34;
			this.Label14.Text = "Phone";
			this.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label15
			//
			this.Label15.Location = new System.Drawing.Point(8, 342);
			this.Label15.Name = "Label15";
			this.Label15.Size = new System.Drawing.Size(80, 16);
			this.Label15.TabIndex = 35;
			this.Label15.Text = "Fax";
			this.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label16
			//
			this.Label16.Location = new System.Drawing.Point(8, 366);
			this.Label16.Name = "Label16";
			this.Label16.Size = new System.Drawing.Size(80, 16);
			this.Label16.TabIndex = 36;
			this.Label16.Text = "Email";
			this.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label17
			//
			this.Label17.Location = new System.Drawing.Point(8, 270);
			this.Label17.Name = "Label17";
			this.Label17.Size = new System.Drawing.Size(80, 16);
			this.Label17.TabIndex = 32;
			this.Label17.Text = "Zip Code";
			this.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//ErrorProvider1
			//
			this.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
			this.ErrorProvider1.ContainerControl = this;
			//
			//frmMun
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 397);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.Label7);
			this.Controls.Add(this.Label13);
			this.Controls.Add(this.Label14);
			this.Controls.Add(this.Label15);
			this.Controls.Add(this.Label16);
			this.Controls.Add(this.Label17);
			this.Controls.Add(this.lvMunicipalities);
			this.Controls.Add(this.btnDelete);
			this.Controls.Add(this.btnSave);
			this.Controls.Add(this.btnAdd);
			this.Controls.Add(this.Label6);
			this.Controls.Add(this.txtEmail);
			this.Controls.Add(this.txtFax);
			this.Controls.Add(this.txtPhone);
			this.Controls.Add(this.txtContact);
			this.Controls.Add(this.txtZipCode);
			this.Controls.Add(this.cboState);
			this.Controls.Add(this.txtCity);
			this.Controls.Add(this.txtStreet);
			this.Controls.Add(this.txtDept);
			this.Controls.Add(this.txtName);
			this.Controls.Add(this.cboPreface);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MaximumSize = new System.Drawing.Size(512, 424);
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(440, 424);
			this.Name = "frmMun";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Municipalities";
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		private int municipalityNo;
		string mode = "update";
		
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		// Handle the Form load event.
		private void frmMun_Load(object sender, System.EventArgs e)
		{
			PopulateMunicipalityList();
			lvMunicipalities.Items[0].Selected = true;
		}
		
		private void PopulateMunicipalityList()
		{
			SqlConnection cn = new SqlConnection(connectionString);
			string sql = "SELECT munName, munNo FROM Municipalities WHERE munDeleted = 0 ORDER BY munName";
			SqlCommand cmd = new SqlCommand(sql, cn);
			// lvMunicipalities.Items.Clear()
			lvMunicipalities.Clear();
			
			// Set to details view
			ColumnHeader columnheader; // Used for creating column headers.
			ListViewItem listviewitem; // Used for creating ListView items.
			
			// Make sure that the view is set to show details.
			lvMunicipalities.View = View.Details;
			
			try
			{
				cn.Open();
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					
					// Create some ListView items consisting of first and last names.
					listviewitem = new myListViewItem((dr.item["munName"].ToString() + ""), "Municipality", System.Convert.ToInt32(dr.item["munNo"]), 1);
					this.lvMunicipalities.Items.Add(listviewitem);
					
				}
				// lstMunicipalities.SetSelected(0, True)
				dr.Close();
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
			// Create some column headers for the data.
			columnheader = new ColumnHeader();
			columnheader.Text = "Municipality";
			this.lvMunicipalities.Columns.Add(columnheader);
			
			// Loop through and size each column header to fit the column header text.
			foreach (ColumnHeader tempLoopVar_columnheader in this.lvMunicipalities.Columns)
			{
				columnheader = tempLoopVar_columnheader;
				columnheader.Width = - 2;
			}
			
		}
		
		private void lvMunicipalities_SelectedIndexChanged(System.Object sender, System.EventArgs e)
		{
			ClearErrors();
			if (this.lvMunicipalities.SelectedIndices.Count <= 0)
			{
				return;
			}
			int selNdx = this.lvMunicipalities.SelectedIndices[0];
			if (System.Convert.ToInt32(selNdx) >= 0)
			{
				municipalityNo = ((myListViewItem) lvMunicipalities.Items[selNdx]).intID;
				PopulateFields(municipalityNo);
				mode = "update";
			}
			btnDelete.Enabled = true;
		}
		
		private void ClearErrors()
		{
			ErrorProvider1.SetError(txtName, "");
			ErrorProvider1.SetError(txtStreet, "");
			ErrorProvider1.SetError(txtCity, "");
			ErrorProvider1.SetError(txtZipCode, "");
			ErrorProvider1.SetError(cboState, "");
			ErrorProvider1.SetError(txtPhone, "");
			ErrorProvider1.SetError(txtFax, "");
			ErrorProvider1.SetError(txtEmail, "");
		}
		
		private void PopulateFields(int munNumber)
		{
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(("SELECT * FROM Municipalities WHERE munNo = " + munNumber), cn);
			try
			{
				cn.Open(); // Open Connection
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				if (dr.Read())
				{
					txtName.Text = dr["munName"].ToString() + "";
					txtDept.Text = dr["munDept"].ToString() + "";
					txtStreet.Text = dr["munStrtAddr"].ToString() + "";
					txtCity.Text = dr["munCity"].ToString() + "";
					txtZipCode.Text = dr["munZip"].ToString() + "";
					txtContact.Text = dr["munCntct"].ToString() + "";
					txtPhone.Text = dr["munPhone"].ToString() + "";
					txtFax.Text = dr["munFax"].ToString() + "";
					txtEmail.Text = dr["munEmail"].ToString() + "";
					cboPreface.Text = dr["munPref"].ToString() + "";
					cboState.Text = dr["munState"].ToString() + "";
					
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
			
		}
		
		private void btnAdd_Click(System.Object sender, System.EventArgs e)
		{
			ClearForm();
			mode = "add";
			btnDelete.Enabled = false;
		}
		
		private void ClearForm()
		{
			txtName.Text = "";
			txtDept.Text = "";
			txtStreet.Text = "";
			txtCity.Text = "";
			txtZipCode.Text = "";
			txtContact.Text = "";
			txtPhone.Text = "";
			txtFax.Text = "";
			txtEmail.Text = "";
			cboPreface.Text = "";
			cboState.Text = "";
		}
		
		private void AddOrUpdate()
		{
			if (mode == "add")
			{
				AddMunicipality();
			}
			else
			{
				UpdateMunicipality(municipalityNo);
			}
		}
		
		private void AddMunicipality()
		{
			// Build the INSERT INTO string
			string strInsert = "INSERT INTO Municipalities " + "(munName, munCity, munFax, munPhone, munState, munStrtAddr, munCntct, " + "munPref, munZip, munEmail, munDept) " + "VALUES (";
			
			strInsert += "\'" + txtName.Text + "\', ";
			strInsert += "\'" + txtCity.Text + "\', ";
			strInsert += "\'" + txtFax.Text + "\', ";
			strInsert += "\'" + txtPhone.Text + "\', ";
			strInsert += "\'" + cboState.Text + "\', ";
			strInsert += "\'" + txtStreet.Text + "\', ";
			strInsert += "\'" + txtContact.Text + "\', ";
			strInsert += "\'" + cboPreface.Text + "\', ";
			strInsert += "\'" + txtZipCode.Text + "\', ";
			strInsert += "\'" + txtEmail.Text + "\', ";
			strInsert += "\'" + txtDept.Text + "\')";
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(strInsert, cn);
			cn.Open(); // Open the Connection
			try
			{
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
			ClearForm();
			RefillMunicipalityDT();
		}
		
		private void UpdateMunicipality(int MunicipalityNumber)
		{
			// Build the Update String
			string strUpdate = "UPDATE Municipalities SET ";
			strUpdate += "munName = \'" + txtName.Text + "\', ";
			strUpdate += "munCity = \'" + txtCity.Text + "\', ";
			strUpdate += "munFax = \'" + txtFax.Text + "\', ";
			strUpdate += "munPhone = \'" + txtPhone.Text + "\', ";
			strUpdate += "munState = \'" + cboState.Text + "\', ";
			strUpdate += "munStrtAddr = \'" + txtStreet.Text + "\', ";
			strUpdate += "munCntct = \'" + txtContact.Text + "\', ";
			strUpdate += "munPref = \'" + cboPreface.Text + "\', ";
			strUpdate += "munZip = \'" + txtZipCode.Text + "\', ";
			strUpdate += "munEmail = \'" + txtEmail.Text + "\', ";
			strUpdate += "munDept = \'" + txtDept.Text + "\' ";
			strUpdate += "WHERE munNo = " + MunicipalityNumber;
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(strUpdate, cn);
			cn.Open(); // Open the Connection
			try
			{
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
			ClearForm();
			RefillMunicipalityDT();
			
		}
		
		private void RefillMunicipalityDT()
		{
			SqlConnection cn = new SqlConnection(connectionString);
			string sqlStatement = "SELECT munNo, munName FROM Municipalities ORDER BY munName";
			SqlCommand cmd = new SqlCommand(sqlStatement, cn);
			SqlDataAdapter sda = new SqlDataAdapter(cmd);
			MainForm.ds.Tables["Municip"].Clear();
			
			sda.Fill(MainForm.ds, "Municip"); // Fill the Municipalities datatable
		}
		
		
		private void btnSave_Click(System.Object sender, System.EventArgs e)
		{
			if (FormIsValid())
			{
				AddOrUpdate();
				PopulateMunicipalityList();
				lvMunicipalities.Items[0].Selected = true;
			}
			else
			{
				DialogResult = DialogResult.None;
			}
		}
		
		private void frmMun_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			DialogResult = DialogResult.OK;
		}
		
		private bool FormIsValid()
		{
			if (NameIsValid() && StreetIsValid() && CityIsValid() && ZipIsValid() && StateIsValid() && PhoneIsValid() && FaxIsValid() && EmailIsValid())
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		private bool NameIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtName, "");
			if (txtName.Text.Trim() == "" || txtName.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtName, "Name is a required field");
			}
			return Booly;
		}
		
		private bool StreetIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtStreet, "");
			if (txtStreet.Text.Trim() == "" || txtStreet.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtStreet, "Required Field: e.g. \'1234 North Damen Avenue\'");
			}
			
			char individualChr;
			
			foreach (char tempLoopVar_individualChr in txtStreet.Text)
			{
				individualChr = tempLoopVar_individualChr;
				if (individualChr.CompareTo('\u002E') == 0)
				{
					Booly = false;
					ErrorProvider1.SetError(txtStreet, "Street addresses may not contain periods." + "No abbreviations!");
				}
			}
			return Booly;
		}
		
		private bool CityIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtCity, "");
			Regex CityRegex = new Regex("([A-Z][a-z]+[ ]?)+");
			if (! CityRegex.IsMatch(txtCity.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(txtCity, "Please correct City name (Capitalization, No Numbers, etc.)");
			}
			if (txtCity.Text.Trim() == "" || txtCity.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtCity, "Required Field: e.g. \'Orland Park\'");
			}
			return Booly;
		}
		
		private bool ZipIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtZipCode, "");
			Regex ZipRegex = new Regex("(^\\d{5}$)|(^\\d{5}-\\d{4}$)");
			if (! ZipRegex.IsMatch(txtZipCode.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(txtZipCode, "Zip code must be in the following forms: " + ControlChars.NewLine + "of \'60462\' or \'60462-1499\'");
			}
			if (txtZipCode.Text.Trim() == "" || txtZipCode.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtZipCode, "Required Field: e.g. \'60462\' or \'60462-1499\'");
			}
			return Booly;
		}
		
		private bool StateIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(cboState, "");
			Regex StateRegex = new Regex("^(?-i:A[LKSZRAP]|C[AOT]|D[EC]|F[LM]|G[AU]|HI|I[ADLN]|K[SY]|LA" + "|M[ADEHINOPST]|N[CDEHJMVY]|O[HKR]|P[ARW]|RI|S[CD]|T[NX]|UT|" + "V[AIT]|W[AIVY])$");
			if (! StateRegex.IsMatch(cboState.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(cboState, "State abbreviation must be in the following form: \'IL\'");
			}
			if (cboState.Text.Trim() == "" || cboState.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(cboState, "Required Field: e.g. \'IL\'");
			}
			return Booly;
		}
		
		private bool PhoneIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtPhone, "");
			if (txtPhone.Text.Trim() != ""&& txtPhone.Text != null)
			{
				Regex PhoneRegex = new Regex("^\\([1-9]\\d{2}\\)\\s?\\d{3}\\-\\d{4}( Ext.[0-9]+)?$");
				if (! PhoneRegex.IsMatch(txtPhone.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtPhone, "Phone numbers must be in the following form: " + ControlChars.NewLine + "\'(708)123-4567\' or \'(708)123-4567 Ext.456\'");
				}
			}
			return Booly;
		}
		
		private bool FaxIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtFax, "");
			if (txtFax.Text.Trim() != ""&& txtFax.Text != null)
			{
				Regex FaxRegex = new Regex("^\\([1-9]\\d{2}\\)\\s?\\d{3}\\-\\d{4}$");
				if (! FaxRegex.IsMatch(txtFax.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtFax, "Fax numbers must be in the following form: " + ControlChars.NewLine + "\'(708)123-4567\'");
				}
			}
			return Booly;
		}
		
		private bool EmailIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtEmail, "");
			if (txtEmail.Text.Trim() != ""&& txtEmail.Text != null)
			{
				Regex EmailRegex = new Regex("^([a-zA-Z0-9_\\-])([a-zA-Z0-9_\\-\\.]*)@(\\[((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\\.){3}|((([a-zA-Z0-9\\-]+)\\.)+))([a-zA-Z]{2,}|(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\\])$");
				if (! EmailRegex.IsMatch(txtEmail.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtEmail, "Email addresses must be in the following form: " + ControlChars.NewLine + "\'johnq@hotmail.com\'");
				}
			}
			return Booly;
		}
		
		private void btnDelete_Click(System.Object sender, System.EventArgs e)
		{
			if (MessageBox.Show("Are you sure you want to delete the Municipality?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				
				int selNdx = this.lvMunicipalities.SelectedIndices[0];
				if (selNdx >= 0)
				{
					myListViewItem municipality = (myListViewItem) lvMunicipalities.Items[selNdx];
					string strUpdate = "UPDATE Properties SET munNo = 294 WHERE " + "munNo = " + municipality.intID;
					SqlConnection cn = new SqlConnection(connectionString);
					SqlCommand cmd = new SqlCommand(strUpdate, cn);
					try
					{
						cn.Open(); // Open the Connection
						if (municipality.intID != 294)
						{
							cmd.ExecuteNonQuery();
							cmd.CommandText = "DELETE FROM Municipalities WHERE munNo = " + municipality.intID;
							cmd.ExecuteNonQuery();
						}
					}
					catch (SqlException ex)
					{
						//A SqlException has occured - display details
						int i;
						string msg = "";
						// 						string msg;
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
					
					municipality.Remove();
					ClearErrors();
					ClearForm();
					mode = "add";
					btnDelete.Enabled = false;
				}
			}
			
		}
		
	}
	
}
