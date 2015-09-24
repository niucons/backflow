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
	public class PropertyAddForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public PropertyAddForm()
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
		internal System.Windows.Forms.Button btnOK;
		internal System.Windows.Forms.Button btnCancel;
		internal System.Windows.Forms.Label lblManNotes;
		internal System.Windows.Forms.RichTextBox rtbPropNotes;
		internal System.Windows.Forms.Label Label10;
		internal System.Windows.Forms.TextBox txtPropStore;
		internal System.Windows.Forms.Label Label8;
		internal System.Windows.Forms.Label Label7;
		internal System.Windows.Forms.Label Label6;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.TextBox txtPropFaxNo;
		internal System.Windows.Forms.TextBox txtPropPhone;
		internal System.Windows.Forms.TextBox txtPropContact;
		internal System.Windows.Forms.TextBox txtPropZipCode;
		internal System.Windows.Forms.TextBox txtPropCity;
		internal System.Windows.Forms.TextBox txtPropStrtAdd;
		internal System.Windows.Forms.TextBox txtPropName;
		internal System.Windows.Forms.Label Label11;
		internal System.Windows.Forms.ComboBox cboMunicipality;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.ComboBox cboState;
		internal System.Windows.Forms.Button btnManToProp;
		internal System.Windows.Forms.ErrorProvider ErrorProvider1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PropertyAddForm));
			this.btnOK = new System.Windows.Forms.Button();
			base.Load += new System.EventHandler(frmPropAdd_Load);
			this.btnOK.Click += new System.EventHandler(btnOK_Click);
			this.btnCancel = new System.Windows.Forms.Button();
			this.lblManNotes = new System.Windows.Forms.Label();
			this.rtbPropNotes = new System.Windows.Forms.RichTextBox();
			this.Label10 = new System.Windows.Forms.Label();
			this.txtPropStore = new System.Windows.Forms.TextBox();
			this.Label8 = new System.Windows.Forms.Label();
			this.Label7 = new System.Windows.Forms.Label();
			this.Label6 = new System.Windows.Forms.Label();
			this.Label5 = new System.Windows.Forms.Label();
			this.Label4 = new System.Windows.Forms.Label();
			this.Label3 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.txtPropFaxNo = new System.Windows.Forms.TextBox();
			this.txtPropPhone = new System.Windows.Forms.TextBox();
			this.txtPropContact = new System.Windows.Forms.TextBox();
			this.txtPropZipCode = new System.Windows.Forms.TextBox();
			this.txtPropCity = new System.Windows.Forms.TextBox();
			this.txtPropStrtAdd = new System.Windows.Forms.TextBox();
			this.txtPropName = new System.Windows.Forms.TextBox();
			this.Label11 = new System.Windows.Forms.Label();
			this.cboMunicipality = new System.Windows.Forms.ComboBox();
			this.Label1 = new System.Windows.Forms.Label();
			this.cboState = new System.Windows.Forms.ComboBox();
			this.btnManToProp = new System.Windows.Forms.Button();
			this.btnManToProp.Click += new System.EventHandler(btnManToProp_Click);
			this.ErrorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.SuspendLayout();
			//
			//btnOK
			//
			this.btnOK.Location = new System.Drawing.Point(308, 426);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(56, 24);
			this.btnOK.TabIndex = 11;
			this.btnOK.Text = "OK";
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(244, 426);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(56, 24);
			this.btnCancel.TabIndex = 78;
			this.btnCancel.Text = "Cancel";
			//
			//lblManNotes
			//
			this.lblManNotes.Location = new System.Drawing.Point(12, 258);
			this.lblManNotes.Name = "lblManNotes";
			this.lblManNotes.Size = new System.Drawing.Size(144, 16);
			this.lblManNotes.TabIndex = 89;
			this.lblManNotes.Text = "Property Specific Notes";
			//
			//rtbPropNotes
			//
			this.rtbPropNotes.Location = new System.Drawing.Point(12, 282);
			this.rtbPropNotes.Name = "rtbPropNotes";
			this.rtbPropNotes.Size = new System.Drawing.Size(352, 128);
			this.rtbPropNotes.TabIndex = 10;
			this.rtbPropNotes.Text = "";
			//
			//Label10
			//
			this.Label10.Location = new System.Drawing.Point(32, 120);
			this.Label10.Name = "Label10";
			this.Label10.Size = new System.Drawing.Size(56, 16);
			this.Label10.TabIndex = 88;
			this.Label10.Text = "Store No.";
			this.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtPropStore
			//
			this.txtPropStore.HideSelection = false;
			this.txtPropStore.Location = new System.Drawing.Point(96, 120);
			this.txtPropStore.Name = "txtPropStore";
			this.txtPropStore.Size = new System.Drawing.Size(152, 20);
			this.txtPropStore.TabIndex = 5;
			this.txtPropStore.Text = "";
			//
			//Label8
			//
			this.Label8.Location = new System.Drawing.Point(264, 72);
			this.Label8.Name = "Label8";
			this.Label8.Size = new System.Drawing.Size(32, 24);
			this.Label8.TabIndex = 87;
			this.Label8.Text = "State";
			this.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label7
			//
			this.Label7.Location = new System.Drawing.Point(0, 192);
			this.Label7.Name = "Label7";
			this.Label7.Size = new System.Drawing.Size(88, 16);
			this.Label7.TabIndex = 86;
			this.Label7.Text = "Fax No.";
			this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label6
			//
			this.Label6.Location = new System.Drawing.Point(0, 168);
			this.Label6.Name = "Label6";
			this.Label6.Size = new System.Drawing.Size(88, 16);
			this.Label6.TabIndex = 85;
			this.Label6.Text = "Phone No.";
			this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label5
			//
			this.Label5.Location = new System.Drawing.Point(0, 144);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(88, 16);
			this.Label5.TabIndex = 84;
			this.Label5.Text = "Contact Person";
			this.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label4
			//
			this.Label4.Location = new System.Drawing.Point(0, 96);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(88, 16);
			this.Label4.TabIndex = 83;
			this.Label4.Text = "Postal Code";
			this.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label3
			//
			this.Label3.Location = new System.Drawing.Point(0, 72);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(88, 16);
			this.Label3.TabIndex = 82;
			this.Label3.Text = "City";
			this.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label2
			//
			this.Label2.Location = new System.Drawing.Point(8, 48);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(80, 16);
			this.Label2.TabIndex = 81;
			this.Label2.Text = "Street Address";
			this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtPropFaxNo
			//
			this.txtPropFaxNo.HideSelection = false;
			this.txtPropFaxNo.Location = new System.Drawing.Point(96, 192);
			this.txtPropFaxNo.Name = "txtPropFaxNo";
			this.txtPropFaxNo.Size = new System.Drawing.Size(152, 20);
			this.txtPropFaxNo.TabIndex = 8;
			this.txtPropFaxNo.Text = "";
			//
			//txtPropPhone
			//
			this.txtPropPhone.HideSelection = false;
			this.txtPropPhone.Location = new System.Drawing.Point(96, 168);
			this.txtPropPhone.Name = "txtPropPhone";
			this.txtPropPhone.Size = new System.Drawing.Size(152, 20);
			this.txtPropPhone.TabIndex = 7;
			this.txtPropPhone.Text = "";
			//
			//txtPropContact
			//
			this.txtPropContact.HideSelection = false;
			this.txtPropContact.Location = new System.Drawing.Point(96, 144);
			this.txtPropContact.Name = "txtPropContact";
			this.txtPropContact.Size = new System.Drawing.Size(152, 20);
			this.txtPropContact.TabIndex = 6;
			this.txtPropContact.Text = "";
			//
			//txtPropZipCode
			//
			this.txtPropZipCode.Location = new System.Drawing.Point(96, 96);
			this.txtPropZipCode.Name = "txtPropZipCode";
			this.txtPropZipCode.Size = new System.Drawing.Size(152, 20);
			this.txtPropZipCode.TabIndex = 4;
			this.txtPropZipCode.Text = "";
			//
			//txtPropCity
			//
			this.txtPropCity.Location = new System.Drawing.Point(96, 72);
			this.txtPropCity.Name = "txtPropCity";
			this.txtPropCity.Size = new System.Drawing.Size(152, 20);
			this.txtPropCity.TabIndex = 2;
			this.txtPropCity.Text = "";
			//
			//txtPropStrtAdd
			//
			this.txtPropStrtAdd.Location = new System.Drawing.Point(96, 48);
			this.txtPropStrtAdd.Name = "txtPropStrtAdd";
			this.txtPropStrtAdd.Size = new System.Drawing.Size(256, 20);
			this.txtPropStrtAdd.TabIndex = 1;
			this.txtPropStrtAdd.Text = "";
			//
			//txtPropName
			//
			this.txtPropName.ForeColor = System.Drawing.SystemColors.WindowText;
			this.txtPropName.Location = new System.Drawing.Point(96, 24);
			this.txtPropName.Name = "txtPropName";
			this.txtPropName.Size = new System.Drawing.Size(256, 20);
			this.txtPropName.TabIndex = 0;
			this.txtPropName.Text = "";
			//
			//Label11
			//
			this.Label11.Location = new System.Drawing.Point(0, 24);
			this.Label11.Name = "Label11";
			this.Label11.Size = new System.Drawing.Size(88, 16);
			this.Label11.TabIndex = 80;
			this.Label11.Text = "Property Name";
			this.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//cboMunicipality
			//
			this.cboMunicipality.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboMunicipality.Location = new System.Drawing.Point(96, 216);
			this.cboMunicipality.Name = "cboMunicipality";
			this.cboMunicipality.Size = new System.Drawing.Size(152, 21);
			this.cboMunicipality.TabIndex = 9;
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(0, 216);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(88, 16);
			this.Label1.TabIndex = 90;
			this.Label1.Text = "Municipality";
			this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//cboState
			//
			this.cboState.Items.AddRange(new object[] { "IL", "IN", "IA", "WI", "MI" });
			this.cboState.Location = new System.Drawing.Point(304, 72);
			this.cboState.Name = "cboState";
			this.cboState.Size = new System.Drawing.Size(48, 21);
			this.cboState.TabIndex = 3;
			//
			//btnManToProp
			//
			this.btnManToProp.Location = new System.Drawing.Point(12, 426);
			this.btnManToProp.Name = "btnManToProp";
			this.btnManToProp.Size = new System.Drawing.Size(154, 24);
			this.btnManToProp.TabIndex = 116;
			this.btnManToProp.Text = "Manager Info to Property";
			//
			//ErrorProvider1
			//
			this.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
			this.ErrorProvider1.ContainerControl = this;
			//
			//frmPropAdd
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 461);
			this.Controls.Add(this.btnManToProp);
			this.Controls.Add(this.cboState);
			this.Controls.Add(this.cboMunicipality);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.lblManNotes);
			this.Controls.Add(this.rtbPropNotes);
			this.Controls.Add(this.Label10);
			this.Controls.Add(this.txtPropStore);
			this.Controls.Add(this.txtPropFaxNo);
			this.Controls.Add(this.txtPropPhone);
			this.Controls.Add(this.txtPropContact);
			this.Controls.Add(this.txtPropZipCode);
			this.Controls.Add(this.txtPropCity);
			this.Controls.Add(this.txtPropStrtAdd);
			this.Controls.Add(this.txtPropName);
			this.Controls.Add(this.Label8);
			this.Controls.Add(this.Label7);
			this.Controls.Add(this.Label6);
			this.Controls.Add(this.Label5);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.Label11);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmPropAdd";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "New Property";
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		public static int manNumber; // Set through frmMain
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		
		// Handle the Form load event.
		private void frmPropAdd_Load(object sender, System.EventArgs e)
		{
			FillComboBoxes();
		}
		
		private void FillComboBoxes()
		{
			cboMunicipality.ValueMember = "munNo";
			cboMunicipality.DisplayMember = "munName";
			cboMunicipality.DataSource = MainForm.ds.Tables["Municip"];
			cboMunicipality.SelectedValue = 294;
		}
		
		public void AddProperty()
		{
			// Build the INSERT INTO string
			string strInsert = "INSERT INTO Properties " + "(munNo, manNo, propName, propStrt, propCity, propState, propZip, storeNo, " + "propCon, propPhone, propFax, propAdded, propNotes) " + "VALUES (";
			
			strInsert += cboMunicipality.SelectedValue.ToString() + ", ";
			strInsert += manNumber + ", ";
			strInsert += "\'" + txtPropName.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + txtPropStrtAdd.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + txtPropCity.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.cboState.Text + "\', ";
			strInsert += "\'" + this.txtPropZipCode.Text + "\', ";
			strInsert += "\'" + txtPropStore.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + txtPropContact.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.txtPropPhone.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.txtPropFaxNo.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + DateTime.Today.ToShortDateString() + "\', ";
			strInsert += "\'" + rtbPropNotes.Text.Replace("\'", "\'\'") + "\')";
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(strInsert, cn);
			try
			{
				cn.Open(); // Open the Connection
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
				DialogResult = DialogResult.OK;
			}
			
		}
		
		private void btnManToProp_Click(System.Object sender, System.EventArgs e)
		{
			GetManData();
		}
		
		private void GetManData()
		{
			
			string PropName = "";
			string Street = "";
			string City = "";
			string State = "";
			string Zip = "";
			string Store = "";
			string Contact = "";
			string Phone = "";
			string Fax = "";
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(("SELECT * FROM Managers WHERE manNo = " + manNumber), cn);
			// Run the query; get the DataReader object.
			
			try
			{
				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					PropName = dr["manName"].ToString() + "";
					Street = dr["manStrtAdd"].ToString() + "";
					City = dr["manCity"].ToString() + "";
					State = dr["manState"].ToString() + "";
					Zip = dr["manZip"].ToString() + "";
					Store = dr["manSuite"].ToString() + "";
					Contact = dr["manCntct"].ToString() + "";
					Phone = dr["manPhone"].ToString() + "";
					Fax = dr["manFax"].ToString() + "";
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
			
			this.txtPropName.Text = PropName;
			this.txtPropStrtAdd.Text = Street;
			this.txtPropCity.Text = City;
			this.cboState.Text = State;
			this.txtPropZipCode.Text = Zip;
			this.txtPropStore.Text = Store;
			this.txtPropContact.Text = Contact;
			this.txtPropPhone.Text = Phone;
			this.txtPropFaxNo.Text = Fax;
			
			this.cboMunicipality.SelectedValue = 294;
			
		} //GetManData
		
		private bool FormIsValid()
		{
			if (NameIsValid() && StreetIsValid() && CityIsValid() && ZipIsValid() && StateIsValid())
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
			if (1 == 2)
			{
				ArrayList SimilarProperties = SimilarPropertyList[];
				ErrorProvider1.SetError(txtPropName, "");
				if (SimilarProperties.Count > 0)
				{
					Booly = false;
					string LineItem;
					string FullText;
					FullText = "Property Already Exists: " + ControlChars.NewLine;
					foreach (string tempLoopVar_LineItem in SimilarProperties)
					{
						LineItem = tempLoopVar_LineItem;
						FullText += LineItem;
					}
					ErrorProvider1.SetError(txtPropName, FullText);
				}
			}
			if (txtPropName.Text.Trim() == "" || txtPropName.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtPropName, "Required Field: e.g. \'Best Buy\'");
			}
			return Booly;
		}
		
		private bool StreetIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtPropStrtAdd, "");
			if (txtPropStrtAdd.Text.Trim() == "" || txtPropStrtAdd.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtPropStrtAdd, "Required Field: e.g. \'1234 North Damen Avenue\'");
			}
			
			char individualChr;
			
			foreach (char tempLoopVar_individualChr in txtPropStrtAdd.Text)
			{
				individualChr = tempLoopVar_individualChr;
				if (individualChr.CompareTo('\u002E') == 0)
				{
					Booly = false;
					ErrorProvider1.SetError(txtPropStrtAdd, "Street addresses may not contain periods." + "No abbreviations!");
				}
			}
			return Booly;
		}
		
		private bool CityIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtPropCity, "");
			Regex CityRegex = new Regex("([A-Z][a-z]+[ ]?)+");
			if (! CityRegex.IsMatch(txtPropCity.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(txtPropCity, "Please correct City name (Capitalization, No Numbers, etc.)");
			}
			if (txtPropCity.Text.Trim() == "" || txtPropCity.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtPropCity, "Required Field: e.g. \'Orland Park\'");
			}
			return Booly;
		}
		
		private bool ZipIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtPropZipCode, "");
			Regex ZipRegex = new Regex("(^\\d{5}$)|(^\\d{5}-\\d{4}$)");
			if (! ZipRegex.IsMatch(txtPropZipCode.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(txtPropZipCode, "Zip code must be in the following forms: " + ControlChars.NewLine + "of \'60462\' or \'60462-1499\'");
			}
			if (txtPropZipCode.Text.Trim() == "" || txtPropCity.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtPropZipCode, "Required Field: e.g. \'60462\' or \'60462-1499\'");
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
			ErrorProvider1.SetError(txtPropPhone, "");
			if (txtPropPhone.Text.Trim() != ""&& txtPropPhone.Text != null)
			{
				Regex PhoneRegex = new Regex("^\\([1-9]\\d{2}\\)\\s?\\d{3}\\-\\d{4}( Ext.[0-9]+)?$");
				if (! PhoneRegex.IsMatch(txtPropPhone.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtPropPhone, "Phone numbers must be in the following form: " + ControlChars.NewLine + "\'(708)123-4567\' or \'(708)123-4567 Ext.456\'");
				}
			}
			return Booly;
		}
		
		private bool FaxIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtPropFaxNo, "");
			if (txtPropFaxNo.Text.Trim() != ""&& txtPropFaxNo.Text != null)
			{
				Regex FaxRegex = new Regex("^\\([1-9]\\d{2}\\)\\s?\\d{3}\\-\\d{4}$");
				if (! FaxRegex.IsMatch(txtPropFaxNo.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtPropFaxNo, "Fax numbers must be in the following form: " + ControlChars.NewLine + "\'(708)123-4567\'");
				}
			}
			return Booly;
		}
		
		private void btnOK_Click(System.Object sender, System.EventArgs e)
		{
			if (FormIsValid())
			{
				AddProperty();
			}
			else
			{
				DialogResult = DialogResult.None;
			}
		}
		
		private ArrayList SimilarPropertyList()
		{
			
			ArrayList resultList = new ArrayList();
			SqlConnection cn = new SqlConnection(connectionString);
			string strQuery = "SELECT propName, propStrt FROM Properties WHERE (propName = \'" + txtPropName.Text.Replace("\'", "\'\'") + "\' AND propStrt = \'" + txtPropStrtAdd.Text.Replace("\'", "\'\'") + "\') AND propDeleted = 0";
			
			SqlCommand cmd = new SqlCommand(strQuery, cn);
			try
			{
				cn.Open(); // Open Connection
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					resultList.Add(dr["propName"].ToString() + " at " + dr["propStrt"].ToString() + ControlChars.NewLine);
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
			
			return resultList;
			
		}
		
	}
	
}
