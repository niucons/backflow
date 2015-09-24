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
	public class ManagerDetailForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public ManagerDetailForm()
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
		internal System.Windows.Forms.TextBox txtManNo;
		internal System.Windows.Forms.Button btnOK;
		internal System.Windows.Forms.Button btnCancel;
		internal System.Windows.Forms.Label lblManNotes;
		internal System.Windows.Forms.RichTextBox rtbManNotes;
		internal System.Windows.Forms.Label Label10;
		internal System.Windows.Forms.TextBox txtManOffice;
		internal System.Windows.Forms.Label Label8;
		internal System.Windows.Forms.Label Label7;
		internal System.Windows.Forms.Label Label6;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.CheckBox chkManActive;
		internal System.Windows.Forms.TextBox txtManFaxNo;
		internal System.Windows.Forms.TextBox txtManPhone;
		internal System.Windows.Forms.TextBox txtManContact;
		internal System.Windows.Forms.TextBox txtManZipCode;
		internal System.Windows.Forms.TextBox txtManCity;
		internal System.Windows.Forms.TextBox txtManStrtAdd;
		internal System.Windows.Forms.TextBox txtManName;
		internal System.Windows.Forms.Label Label11;
		internal System.Windows.Forms.ComboBox cboState;
		internal System.Windows.Forms.GroupBox grpPricing;
		internal System.Windows.Forms.TextBox txtPriceList;
		internal System.Windows.Forms.Label lblPricingScheme;
		internal System.Windows.Forms.Button btnRight;
		internal System.Windows.Forms.Button btnLeft;
		internal System.Windows.Forms.TextBox txtManEmail;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.ErrorProvider ErrorProvider1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ManagerDetailForm));
			this.txtManNo = new System.Windows.Forms.TextBox();
			base.Load += new System.EventHandler(frmManDetails_Load);
			this.btnOK = new System.Windows.Forms.Button();
			this.btnOK.Click += new System.EventHandler(btnOK_Click);
			this.btnCancel = new System.Windows.Forms.Button();
			this.lblManNotes = new System.Windows.Forms.Label();
			this.rtbManNotes = new System.Windows.Forms.RichTextBox();
			this.Label10 = new System.Windows.Forms.Label();
			this.txtManOffice = new System.Windows.Forms.TextBox();
			this.Label8 = new System.Windows.Forms.Label();
			this.Label7 = new System.Windows.Forms.Label();
			this.Label6 = new System.Windows.Forms.Label();
			this.Label5 = new System.Windows.Forms.Label();
			this.Label4 = new System.Windows.Forms.Label();
			this.Label3 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.chkManActive = new System.Windows.Forms.CheckBox();
			this.txtManFaxNo = new System.Windows.Forms.TextBox();
			this.txtManPhone = new System.Windows.Forms.TextBox();
			this.txtManContact = new System.Windows.Forms.TextBox();
			this.txtManZipCode = new System.Windows.Forms.TextBox();
			this.txtManCity = new System.Windows.Forms.TextBox();
			this.txtManStrtAdd = new System.Windows.Forms.TextBox();
			this.txtManName = new System.Windows.Forms.TextBox();
			this.Label11 = new System.Windows.Forms.Label();
			this.cboState = new System.Windows.Forms.ComboBox();
			this.grpPricing = new System.Windows.Forms.GroupBox();
			this.txtPriceList = new System.Windows.Forms.TextBox();
			this.lblPricingScheme = new System.Windows.Forms.Label();
			this.btnRight = new System.Windows.Forms.Button();
			this.btnRight.Click += new System.EventHandler(btnRight_Click);
			this.btnLeft = new System.Windows.Forms.Button();
			this.btnLeft.Click += new System.EventHandler(btnLeft_Click);
			this.txtManEmail = new System.Windows.Forms.TextBox();
			this.Label1 = new System.Windows.Forms.Label();
			this.ErrorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.grpPricing.SuspendLayout();
			this.SuspendLayout();
			//
			//txtManNo
			//
			this.txtManNo.Enabled = false;
			this.txtManNo.Location = new System.Drawing.Point(11, 408);
			this.txtManNo.Name = "txtManNo";
			this.txtManNo.Size = new System.Drawing.Size(72, 20);
			this.txtManNo.TabIndex = 65;
			this.txtManNo.Text = "";
			this.txtManNo.Visible = false;
			//
			//btnOK
			//
			this.btnOK.Location = new System.Drawing.Point(251, 408);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(56, 24);
			this.btnOK.TabIndex = 13;
			this.btnOK.Text = "OK";
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(187, 408);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(56, 24);
			this.btnCancel.TabIndex = 14;
			this.btnCancel.Text = "Cancel";
			//
			//lblManNotes
			//
			this.lblManNotes.Location = new System.Drawing.Point(16, 248);
			this.lblManNotes.Name = "lblManNotes";
			this.lblManNotes.Size = new System.Drawing.Size(144, 16);
			this.lblManNotes.TabIndex = 64;
			this.lblManNotes.Text = "Manager Specific Notes";
			//
			//rtbManNotes
			//
			this.rtbManNotes.Location = new System.Drawing.Point(7, 264);
			this.rtbManNotes.Name = "rtbManNotes";
			this.rtbManNotes.Size = new System.Drawing.Size(352, 128);
			this.rtbManNotes.TabIndex = 11;
			this.rtbManNotes.Text = "";
			//
			//Label10
			//
			this.Label10.Location = new System.Drawing.Point(39, 112);
			this.Label10.Name = "Label10";
			this.Label10.Size = new System.Drawing.Size(56, 16);
			this.Label10.TabIndex = 63;
			this.Label10.Text = "Office No.";
			this.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtManOffice
			//
			this.txtManOffice.HideSelection = false;
			this.txtManOffice.Location = new System.Drawing.Point(103, 112);
			this.txtManOffice.Name = "txtManOffice";
			this.txtManOffice.Size = new System.Drawing.Size(152, 20);
			this.txtManOffice.TabIndex = 6;
			this.txtManOffice.Text = "";
			//
			//Label8
			//
			this.Label8.Location = new System.Drawing.Point(267, 64);
			this.Label8.Name = "Label8";
			this.Label8.Size = new System.Drawing.Size(32, 24);
			this.Label8.TabIndex = 62;
			this.Label8.Text = "State";
			this.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label7
			//
			this.Label7.Location = new System.Drawing.Point(7, 184);
			this.Label7.Name = "Label7";
			this.Label7.Size = new System.Drawing.Size(88, 16);
			this.Label7.TabIndex = 61;
			this.Label7.Text = "Fax No.";
			this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label6
			//
			this.Label6.Location = new System.Drawing.Point(7, 160);
			this.Label6.Name = "Label6";
			this.Label6.Size = new System.Drawing.Size(88, 16);
			this.Label6.TabIndex = 60;
			this.Label6.Text = "Phone No.";
			this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label5
			//
			this.Label5.Location = new System.Drawing.Point(7, 136);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(88, 16);
			this.Label5.TabIndex = 59;
			this.Label5.Text = "Contact Person";
			this.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label4
			//
			this.Label4.Location = new System.Drawing.Point(7, 88);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(88, 16);
			this.Label4.TabIndex = 58;
			this.Label4.Text = "Postal Code";
			this.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label3
			//
			this.Label3.Location = new System.Drawing.Point(7, 64);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(88, 16);
			this.Label3.TabIndex = 57;
			this.Label3.Text = "City";
			this.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label2
			//
			this.Label2.Location = new System.Drawing.Point(15, 40);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(80, 16);
			this.Label2.TabIndex = 56;
			this.Label2.Text = "Street Address";
			this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//chkManActive
			//
			this.chkManActive.AccessibleName = "asdfasdf";
			this.chkManActive.Location = new System.Drawing.Point(247, 240);
			this.chkManActive.Name = "chkManActive";
			this.chkManActive.Size = new System.Drawing.Size(112, 24);
			this.chkManActive.TabIndex = 12;
			this.chkManActive.Text = "Active Account";
			//
			//txtManFaxNo
			//
			this.txtManFaxNo.HideSelection = false;
			this.txtManFaxNo.Location = new System.Drawing.Point(103, 184);
			this.txtManFaxNo.Name = "txtManFaxNo";
			this.txtManFaxNo.Size = new System.Drawing.Size(152, 20);
			this.txtManFaxNo.TabIndex = 9;
			this.txtManFaxNo.Text = "";
			//
			//txtManPhone
			//
			this.txtManPhone.HideSelection = false;
			this.txtManPhone.Location = new System.Drawing.Point(103, 160);
			this.txtManPhone.Name = "txtManPhone";
			this.txtManPhone.Size = new System.Drawing.Size(152, 20);
			this.txtManPhone.TabIndex = 8;
			this.txtManPhone.Text = "";
			//
			//txtManContact
			//
			this.txtManContact.HideSelection = false;
			this.txtManContact.Location = new System.Drawing.Point(103, 136);
			this.txtManContact.Name = "txtManContact";
			this.txtManContact.Size = new System.Drawing.Size(152, 20);
			this.txtManContact.TabIndex = 7;
			this.txtManContact.Text = "";
			//
			//txtManZipCode
			//
			this.txtManZipCode.Location = new System.Drawing.Point(103, 88);
			this.txtManZipCode.Name = "txtManZipCode";
			this.txtManZipCode.Size = new System.Drawing.Size(152, 20);
			this.txtManZipCode.TabIndex = 5;
			this.txtManZipCode.Text = "";
			//
			//txtManCity
			//
			this.txtManCity.Location = new System.Drawing.Point(103, 64);
			this.txtManCity.Name = "txtManCity";
			this.txtManCity.Size = new System.Drawing.Size(152, 20);
			this.txtManCity.TabIndex = 3;
			this.txtManCity.Text = "";
			//
			//txtManStrtAdd
			//
			this.txtManStrtAdd.Location = new System.Drawing.Point(103, 40);
			this.txtManStrtAdd.Name = "txtManStrtAdd";
			this.txtManStrtAdd.Size = new System.Drawing.Size(256, 20);
			this.txtManStrtAdd.TabIndex = 2;
			this.txtManStrtAdd.Text = "";
			//
			//txtManName
			//
			this.txtManName.ForeColor = System.Drawing.SystemColors.WindowText;
			this.txtManName.Location = new System.Drawing.Point(103, 16);
			this.txtManName.Name = "txtManName";
			this.txtManName.Size = new System.Drawing.Size(256, 20);
			this.txtManName.TabIndex = 1;
			this.txtManName.Text = "";
			//
			//Label11
			//
			this.Label11.Location = new System.Drawing.Point(7, 16);
			this.Label11.Name = "Label11";
			this.Label11.Size = new System.Drawing.Size(88, 16);
			this.Label11.TabIndex = 55;
			this.Label11.Text = "Company Name";
			this.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//cboState
			//
			this.cboState.Items.AddRange(new object[] { "IL", "IN", "IA", "WI", "MI" });
			this.cboState.Location = new System.Drawing.Point(311, 64);
			this.cboState.Name = "cboState";
			this.cboState.Size = new System.Drawing.Size(48, 21);
			this.cboState.TabIndex = 4;
			//
			//grpPricing
			//
			this.grpPricing.Controls.Add(this.txtPriceList);
			this.grpPricing.Controls.Add(this.lblPricingScheme);
			this.grpPricing.Controls.Add(this.btnRight);
			this.grpPricing.Controls.Add(this.btnLeft);
			this.grpPricing.Location = new System.Drawing.Point(374, 8);
			this.grpPricing.Name = "grpPricing";
			this.grpPricing.Size = new System.Drawing.Size(128, 384);
			this.grpPricing.TabIndex = 68;
			this.grpPricing.TabStop = false;
			this.grpPricing.Text = "Price List";
			//
			//txtPriceList
			//
			this.txtPriceList.BackColor = System.Drawing.SystemColors.Control;
			this.txtPriceList.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtPriceList.Location = new System.Drawing.Point(21, 24);
			this.txtPriceList.Multiline = true;
			this.txtPriceList.Name = "txtPriceList";
			this.txtPriceList.Size = new System.Drawing.Size(96, 288);
			this.txtPriceList.TabIndex = 81;
			this.txtPriceList.TabStop = false;
			this.txtPriceList.Text = "";
			//
			//lblPricingScheme
			//
			this.lblPricingScheme.Location = new System.Drawing.Point(56, 336);
			this.lblPricingScheme.Name = "lblPricingScheme";
			this.lblPricingScheme.Size = new System.Drawing.Size(24, 24);
			this.lblPricingScheme.TabIndex = 82;
			this.lblPricingScheme.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			//
			//btnRight
			//
			this.btnRight.Location = new System.Drawing.Point(80, 344);
			this.btnRight.Name = "btnRight";
			this.btnRight.Size = new System.Drawing.Size(30, 24);
			this.btnRight.TabIndex = 16;
			this.btnRight.Text = ">";
			//
			//btnLeft
			//
			this.btnLeft.Location = new System.Drawing.Point(24, 344);
			this.btnLeft.Name = "btnLeft";
			this.btnLeft.Size = new System.Drawing.Size(30, 24);
			this.btnLeft.TabIndex = 15;
			this.btnLeft.Text = "<";
			//
			//txtManEmail
			//
			this.txtManEmail.HideSelection = false;
			this.txtManEmail.Location = new System.Drawing.Point(104, 208);
			this.txtManEmail.Name = "txtManEmail";
			this.txtManEmail.Size = new System.Drawing.Size(152, 20);
			this.txtManEmail.TabIndex = 10;
			this.txtManEmail.Text = "";
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(8, 208);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(88, 16);
			this.Label1.TabIndex = 74;
			this.Label1.Text = "Email Address";
			this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//ErrorProvider1
			//
			this.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
			this.ErrorProvider1.ContainerControl = this;
			//
			//frmManDetails
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 445);
			this.Controls.Add(this.txtManEmail);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.grpPricing);
			this.Controls.Add(this.cboState);
			this.Controls.Add(this.txtManNo);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.lblManNotes);
			this.Controls.Add(this.rtbManNotes);
			this.Controls.Add(this.Label10);
			this.Controls.Add(this.txtManOffice);
			this.Controls.Add(this.Label8);
			this.Controls.Add(this.Label7);
			this.Controls.Add(this.Label6);
			this.Controls.Add(this.Label5);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.chkManActive);
			this.Controls.Add(this.txtManFaxNo);
			this.Controls.Add(this.txtManPhone);
			this.Controls.Add(this.txtManContact);
			this.Controls.Add(this.txtManZipCode);
			this.Controls.Add(this.txtManCity);
			this.Controls.Add(this.txtManStrtAdd);
			this.Controls.Add(this.txtManName);
			this.Controls.Add(this.Label11);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmManDetails";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Manager Details";
			this.grpPricing.ResumeLayout(false);
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		private int PricingScheme;
		private int MaxSchemeNo;
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		// Handle the Form load event.
		private void frmManDetails_Load(object sender, System.EventArgs e)
		{
			FillManagersData();
			SetMaxSchemeNo();
			UpdatePriceList();
			btnLeft.Enabled = false;
		}
		
		private void SetMaxSchemeNo()
		{
			SqlConnection cn = new SqlConnection(connectionString);
			string strQuery = "SELECT MAX(manSchemeNo) FROM PricingSchemes";
			SqlCommand cmd = new SqlCommand(strQuery, cn);
			try
			{
				cn.Open(); // Open Connection
				MaxSchemeNo = System.Convert.ToInt32(cmd.ExecuteScalar());
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
		
		public void UpdatePriceList()
		{
			if (PricingScheme != 1)
			{
				bool BackflowByPass = false;
				string str = "Backflow Preventer" + ControlChars.CrLf + ControlChars.CrLf;
				SqlConnection cn = new SqlConnection(connectionString);
				string strSQL = "SELECT DISTINCT PricingSchemes.manSchemeNo, AssemblySizes.AssemSizeNo, " + "AssemblySizes.AssemSize, Price.price, AssemblyUsage.assemUsagePriceNo FROM Price " + "LEFT JOIN AssemblySizes ON AssemblySizes.AssemSizeNo = Price.assemSizeNo " + "LEFT JOIN PricingSchemes ON PricingSchemes.manSchemeNo = Price.manSchemeNo " + "LEFT JOIN AssemblyUsagePrice ON AssemblyUsagePrice.assemUsagePriceNo = Price.assemUsagePriceNo " + "LEFT JOIN AssemblyUsage ON AssemblyUsage.assemUsagePriceNo = AssemblyUsagePrice.assemUsagePriceNo " + "WHERE PricingSchemes.manSchemeNo = " + PricingScheme + "ORDER BY  AssemblyUsage.assemUsagePriceNo, AssemblySizes.AssemSizeNo";
				
				SqlCommand cmd = new SqlCommand(strSQL, cn);
				try
				{
					cn.Open(); // Open Connection
					// Run the query; get the DataReader object.
					SqlDataReader dr = cmd.ExecuteReader();
					
					while (dr.Read())
					{
						if (System.Convert.ToInt32(dr["assemUsagePriceNo"]) == 2 && BackflowByPass == false)
						{
							str += ControlChars.CrLf + "Fire Line By-Pass" + ControlChars.CrLf + ControlChars.CrLf;
							BackflowByPass = true;
						}
						str += dr["AssemSize"].ToString() + ControlChars.Tab;
						str += System.Convert.ToDouble(dr["price"]).ToString("c") + ControlChars.CrLf;
					}
					dr.Close();
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
				
				lblPricingScheme.Text = (PricingScheme - 1).ToString();
				txtPriceList.Text = str;
			}
			else
			{
				txtPriceList.Text = "There are currently no prices associated with this Manager." + "It would be prudent to set a scheme now by clicking the arrows below";
				btnLeft.Enabled = false;
			}
			
			
		}
		
		private void btnRight_Click(System.Object sender, System.EventArgs e)
		{
			if (PricingScheme > 1)
			{
				btnLeft.Enabled = true;
			}
			
			if (PricingScheme < MaxSchemeNo)
			{
				PricingScheme++;
				UpdatePriceList();
				if (PricingScheme == MaxSchemeNo)
				{
					btnRight.Enabled = false;
				}
			}
		}
		
		private void btnLeft_Click(System.Object sender, System.EventArgs e)
		{
			btnRight.Enabled = true;
			if (PricingScheme > 2)
			{
				PricingScheme--;
				UpdatePriceList();
				if (PricingScheme == 2)
				{
					btnLeft.Enabled = false;
				}
			}
		}
		
		private void FillManagersData()
		{
			string ManName = "";
			string Street = "";
			string City = "";
			string State = "";
			string Zip = "";
			string Office = "";
			string Contact = "";
			string Phone = "";
			string Fax = "";
			string Email = "";
			string Notes = "";
			bool CurrentAccount = false;
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(("SELECT * FROM Managers WHERE manNo = " + txtManNo.Text), cn);
			try
			{
				cn.Open(); // Open Connection
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				if (dr.Read())
				{
					ManName = dr["manName"].ToString() + "";
					Street = dr["manStrtAdd"].ToString() + "";
					City = dr["manCity"].ToString() + "";
					State = dr["manState"].ToString() + "";
					Zip = dr["manZip"].ToString() + "";
					Office = dr["manSuite"].ToString() + "";
					Contact = dr["manCntct"].ToString() + "";
					Phone = dr["manPhone"].ToString() + "";
					Fax = dr["manFax"].ToString() + "";
					Email = dr["manEmail"].ToString() + "";
					Notes = dr["manNotes"].ToString() + "";
					CurrentAccount = bool.Parse(dr["manCurAcct"].ToString());
					if (Information.IsDBNull(dr["manSchemeNo"]))
					{
						PricingScheme = 0;
					}
					else
					{
						PricingScheme = System.Convert.ToInt32(dr["manSchemeNo"]);
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
			
			this.txtManName.Text = ManName;
			this.txtManStrtAdd.Text = Street;
			this.txtManCity.Text = City;
			this.cboState.Text = State;
			this.txtManZipCode.Text = Zip;
			this.txtManOffice.Text = Office;
			this.txtManContact.Text = Contact;
			this.txtManPhone.Text = Phone;
			this.txtManFaxNo.Text = Fax;
			this.txtManEmail.Text = Email;
			this.rtbManNotes.Text = Notes;
			this.chkManActive.Checked = CurrentAccount;
			
		} //FillManagersData
		
		private void btnOK_Click(System.Object sender, System.EventArgs e)
		{
			if (FormIsValid())
			{
				UpdateManager();
			}
			else
			{
				DialogResult = DialogResult.None;
			}
		}
		
		private ArrayList SimilarManagerList()
		{
			
			ArrayList resultList = new ArrayList();
			SqlConnection cn = new SqlConnection(connectionString);
			string strQuery = "SELECT manName, manStrtAdd, manNo FROM Managers WHERE (manName = \'" + txtManName.Text.Replace("\'", "\'\'") + "\' AND manStrtAdd = \'" + txtManStrtAdd.Text.Replace("\'", "\'\'") + "\') AND manDeleted = 0";
			SqlCommand cmd = new SqlCommand(strQuery, cn);
			try
			{
				cn.Open(); // Open Connection
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					if (dr["manNo"].ToString() != txtManNo.Text)
					{
						resultList.Add(dr["manName"].ToString() + " at " + dr["manStrtAdd"].ToString() + ControlChars.NewLine);
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
			return resultList;
			
		}
		
		private void UpdateManager()
		{
			// Code to build UPDATE string
			string strUpdate = "UPDATE Managers SET ";
			
			strUpdate += "manName = \'" + this.txtManName.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manStrtAdd = \'" + this.txtManStrtAdd.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manCity = \'" + this.txtManCity.Text + "\', ";
			strUpdate += "manState = \'" + this.cboState.Text + "\', ";
			strUpdate += "manZip = \'" + this.txtManZipCode.Text + "\', ";
			strUpdate += "manSuite = \'" + this.txtManOffice.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manCntct = \'" + this.txtManContact.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manPhone = \'" + this.txtManPhone.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manFax = \'" + this.txtManFaxNo.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manEmail = \'" + this.txtManEmail.Text + "\', ";
			strUpdate += "manNotes = \'" + this.rtbManNotes.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "manCurAcct = " + System.Convert.ToInt32(this.chkManActive.Checked).ToString() + ", ";
			strUpdate += "manSchemeNo = " + PricingScheme + " ";
			strUpdate += "WHERE manNo = " + this.txtManNo.Text;
			
			// Me.rtbManNotes.Text = strUpdate
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(strUpdate, cn);
			
			try
			{
				cn.Open();
				cmd.ExecuteNonQuery();
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
			
			DialogResult = DialogResult.OK;
			
		}
		
		private bool FormIsValid()
		{
			if (NameIsValid() && StreetIsValid() && CityIsValid() && ZipIsValid() && StateIsValid() && EmailIsValid())
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
			
			
			ArrayList SimilarManagers = SimilarManagerList[];
			ErrorProvider1.SetError(txtManName, "");
			if (SimilarManagers.Count > 0)
			{
				Booly = false;
				string LineItem;
				string FullText;
				FullText = "Manager Already Exists: " + ControlChars.NewLine;
				foreach (string tempLoopVar_LineItem in SimilarManagers)
				{
					LineItem = tempLoopVar_LineItem;
					FullText += LineItem;
				}
				ErrorProvider1.SetError(txtManName, FullText);
			}
			
			if (txtManName.Text.Trim() == "" || txtManName.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtManName, "Required Field: e.g. \'Best Buy\'");
			}
			return Booly;
		}
		
		private bool StreetIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtManStrtAdd, "");
			if (txtManStrtAdd.Text.Trim() == "" || txtManStrtAdd.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtManStrtAdd, "Required Field: e.g. \'1234 North Damen Avenue\'");
			}
			
			char individualChr;
			
			foreach (char tempLoopVar_individualChr in txtManStrtAdd.Text)
			{
				individualChr = tempLoopVar_individualChr;
				if (individualChr.CompareTo('\u002E') == 0)
				{
					Booly = false;
					ErrorProvider1.SetError(txtManStrtAdd, "Street addresses may not contain periods." + "No abbreviations!");
				}
			}
			return Booly;
		}
		
		private bool CityIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtManCity, "");
			Regex CityRegex = new Regex("([A-Z][a-z]+[ ]?)+");
			if (! CityRegex.IsMatch(txtManCity.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(txtManCity, "Please correct City name (Capitalization, No Numbers, etc.)");
			}
			if (txtManCity.Text.Trim() == "" || txtManCity.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtManCity, "Required Field: e.g. \'Orland Park\'");
			}
			return Booly;
		}
		
		private bool ZipIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtManZipCode, "");
			Regex ZipRegex = new Regex("(^\\d{5}$)|(^\\d{5}-\\d{4}$)");
			if (! ZipRegex.IsMatch(txtManZipCode.Text))
			{
				Booly = false;
				ErrorProvider1.SetError(txtManZipCode, "Zip code must be in the following forms: " + ControlChars.NewLine + "of \'60462\' or \'60462-1499\'");
			}
			if (txtManZipCode.Text.Trim() == "" || txtManCity.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtManZipCode, "Required Field: e.g. \'60462\' or \'60462-1499\'");
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
			ErrorProvider1.SetError(txtManPhone, "");
			if (txtManPhone.Text.Trim() != ""&& txtManPhone.Text != null)
			{
				Regex PhoneRegex = new Regex("^\\([1-9]\\d{2}\\)\\s?\\d{3}\\-\\d{4}( Ext.[0-9]+)?$");
				if (! PhoneRegex.IsMatch(txtManPhone.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtManPhone, "Phone numbers must be in the following form: " + ControlChars.NewLine + "\'(708)123-4567\' or \'(708)123-4567 Ext.456\'");
				}
			}
			return Booly;
		}
		
		private bool FaxIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtManFaxNo, "");
			if (txtManFaxNo.Text.Trim() != ""&& txtManFaxNo.Text != null)
			{
				Regex FaxRegex = new Regex("^\\([1-9]\\d{2}\\)\\s?\\d{3}\\-\\d{4}$");
				if (! FaxRegex.IsMatch(txtManFaxNo.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtManFaxNo, "Fax numbers must be in the following form: " + ControlChars.NewLine + "\'(708)123-4567\'");
				}
			}
			return Booly;
		}
		
		private bool EmailIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtManEmail, "");
			if (txtManEmail.Text.Trim() != ""&& txtManEmail.Text != null)
			{
				Regex EmailRegex = new Regex("^([a-zA-Z0-9_\\-])([a-zA-Z0-9_\\-\\.]*)@(\\[((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\\.){3}|((([a-zA-Z0-9\\-]+)\\.)+))([a-zA-Z]{2,}|(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\\])$");
				if (! EmailRegex.IsMatch(txtManEmail.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtManEmail, "Email addresses must be in the following form: " + ControlChars.NewLine + "\'johnq@hotmail.com\'");
				}
			}
			return Booly;
		}
	}
	
}
