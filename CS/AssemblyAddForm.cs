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
	public class AssemblyAddForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public AssemblyAddForm()
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
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Label Label10;
		internal System.Windows.Forms.Label Label7;
		internal System.Windows.Forms.Label Label6;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.ComboBox cboSize;
		internal System.Windows.Forms.ComboBox cboManufacturer;
		internal System.Windows.Forms.ComboBox cboType;
		internal System.Windows.Forms.Label lblNotes;
		internal System.Windows.Forms.RichTextBox rtbAssemNotes;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.TextBox txtInstallDate;
		internal System.Windows.Forms.TextBox txtModel;
		internal System.Windows.Forms.TextBox txtSerial;
		internal System.Windows.Forms.TextBox txtLocation;
		internal System.Windows.Forms.ComboBox cboUsage;
		internal System.Windows.Forms.ErrorProvider ErrorProvider1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(AssemblyAddForm));
			this.cboSize = new System.Windows.Forms.ComboBox();
			base.Load += new System.EventHandler(AssemAdd_Load);
			this.cboManufacturer = new System.Windows.Forms.ComboBox();
			this.cboType = new System.Windows.Forms.ComboBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnOK.Click += new System.EventHandler(btnOK_Click);
			this.btnCancel = new System.Windows.Forms.Button();
			this.Label1 = new System.Windows.Forms.Label();
			this.txtInstallDate = new System.Windows.Forms.TextBox();
			this.Label10 = new System.Windows.Forms.Label();
			this.Label7 = new System.Windows.Forms.Label();
			this.Label6 = new System.Windows.Forms.Label();
			this.Label5 = new System.Windows.Forms.Label();
			this.Label4 = new System.Windows.Forms.Label();
			this.Label3 = new System.Windows.Forms.Label();
			this.txtModel = new System.Windows.Forms.TextBox();
			this.txtSerial = new System.Windows.Forms.TextBox();
			this.lblNotes = new System.Windows.Forms.Label();
			this.rtbAssemNotes = new System.Windows.Forms.RichTextBox();
			this.Label2 = new System.Windows.Forms.Label();
			this.txtLocation = new System.Windows.Forms.TextBox();
			this.cboUsage = new System.Windows.Forms.ComboBox();
			this.ErrorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.SuspendLayout();
			//
			//cboSize
			//
			this.cboSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboSize.Location = new System.Drawing.Point(120, 112);
			this.cboSize.Name = "cboSize";
			this.cboSize.Size = new System.Drawing.Size(152, 21);
			this.cboSize.TabIndex = 4;
			//
			//cboManufacturer
			//
			this.cboManufacturer.Items.AddRange(new object[] { "Ames", "Conbraco", "Febco", "Hersey", "Hunter Ames", "Watts", "Wilkens" });
			this.cboManufacturer.Location = new System.Drawing.Point(120, 65);
			this.cboManufacturer.Name = "cboManufacturer";
			this.cboManufacturer.Size = new System.Drawing.Size(152, 21);
			this.cboManufacturer.TabIndex = 2;
			//
			//cboType
			//
			this.cboType.Items.AddRange(new object[] { "RP", "RPDA", "DC", "DCDA" });
			this.cboType.Location = new System.Drawing.Point(120, 41);
			this.cboType.Name = "cboType";
			this.cboType.Size = new System.Drawing.Size(152, 21);
			this.cboType.TabIndex = 1;
			//
			//btnOK
			//
			this.btnOK.Location = new System.Drawing.Point(308, 392);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(56, 24);
			this.btnOK.TabIndex = 9;
			this.btnOK.Text = "OK";
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(244, 392);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(56, 24);
			this.btnCancel.TabIndex = 119;
			this.btnCancel.Text = "Cancel";
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(24, 163);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(88, 16);
			this.Label1.TabIndex = 118;
			this.Label1.Text = "Install Date";
			this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtInstallDate
			//
			this.txtInstallDate.HideSelection = false;
			this.txtInstallDate.Location = new System.Drawing.Point(120, 159);
			this.txtInstallDate.Name = "txtInstallDate";
			this.txtInstallDate.Size = new System.Drawing.Size(152, 20);
			this.txtInstallDate.TabIndex = 6;
			this.txtInstallDate.Text = "";
			//
			//Label10
			//
			this.Label10.Location = new System.Drawing.Point(16, 70);
			this.Label10.Name = "Label10";
			this.Label10.Size = new System.Drawing.Size(96, 16);
			this.Label10.TabIndex = 116;
			this.Label10.Text = "Manufacturer";
			this.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label7
			//
			this.Label7.Location = new System.Drawing.Point(24, 140);
			this.Label7.Name = "Label7";
			this.Label7.Size = new System.Drawing.Size(88, 16);
			this.Label7.TabIndex = 115;
			this.Label7.Text = "Usage";
			this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label6
			//
			this.Label6.Location = new System.Drawing.Point(24, 117);
			this.Label6.Name = "Label6";
			this.Label6.Size = new System.Drawing.Size(88, 16);
			this.Label6.TabIndex = 114;
			this.Label6.Text = "Size";
			this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label5
			//
			this.Label5.Location = new System.Drawing.Point(24, 93);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(88, 16);
			this.Label5.TabIndex = 113;
			this.Label5.Text = "Model";
			this.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label4
			//
			this.Label4.Location = new System.Drawing.Point(24, 46);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(88, 16);
			this.Label4.TabIndex = 112;
			this.Label4.Text = "Type";
			this.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label3
			//
			this.Label3.Location = new System.Drawing.Point(24, 22);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(88, 16);
			this.Label3.TabIndex = 111;
			this.Label3.Text = "Serial";
			this.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtModel
			//
			this.txtModel.HideSelection = false;
			this.txtModel.Location = new System.Drawing.Point(120, 89);
			this.txtModel.Name = "txtModel";
			this.txtModel.Size = new System.Drawing.Size(152, 20);
			this.txtModel.TabIndex = 3;
			this.txtModel.Text = "";
			//
			//txtSerial
			//
			this.txtSerial.Location = new System.Drawing.Point(120, 18);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(152, 20);
			this.txtSerial.TabIndex = 0;
			this.txtSerial.Text = "";
			//
			//lblNotes
			//
			this.lblNotes.Location = new System.Drawing.Point(12, 224);
			this.lblNotes.Name = "lblNotes";
			this.lblNotes.Size = new System.Drawing.Size(144, 16);
			this.lblNotes.TabIndex = 125;
			this.lblNotes.Text = "Assembly Specific Notes";
			//
			//rtbAssemNotes
			//
			this.rtbAssemNotes.Location = new System.Drawing.Point(12, 248);
			this.rtbAssemNotes.Name = "rtbAssemNotes";
			this.rtbAssemNotes.Size = new System.Drawing.Size(352, 128);
			this.rtbAssemNotes.TabIndex = 8;
			this.rtbAssemNotes.Text = "";
			//
			//Label2
			//
			this.Label2.Location = new System.Drawing.Point(24, 186);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(88, 16);
			this.Label2.TabIndex = 127;
			this.Label2.Text = "Location";
			this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtLocation
			//
			this.txtLocation.HideSelection = false;
			this.txtLocation.Location = new System.Drawing.Point(120, 182);
			this.txtLocation.Name = "txtLocation";
			this.txtLocation.Size = new System.Drawing.Size(248, 20);
			this.txtLocation.TabIndex = 7;
			this.txtLocation.Text = "";
			//
			//cboUsage
			//
			this.cboUsage.Items.AddRange(new object[] { "Boiler Feed", "Car Wash", "Filter System", "Fire Line", "Fire Line By-Pass", "Hose Bib", "Lawn Sprinkler", "Potable" });
			this.cboUsage.Location = new System.Drawing.Point(120, 136);
			this.cboUsage.Name = "cboUsage";
			this.cboUsage.Size = new System.Drawing.Size(152, 21);
			this.cboUsage.TabIndex = 5;
			//
			//ErrorProvider1
			//
			this.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
			this.ErrorProvider1.ContainerControl = this;
			//
			//frmAssemAdd
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 429);
			this.Controls.Add(this.cboUsage);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.txtLocation);
			this.Controls.Add(this.lblNotes);
			this.Controls.Add(this.rtbAssemNotes);
			this.Controls.Add(this.cboSize);
			this.Controls.Add(this.cboManufacturer);
			this.Controls.Add(this.cboType);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.txtInstallDate);
			this.Controls.Add(this.Label10);
			this.Controls.Add(this.Label7);
			this.Controls.Add(this.Label6);
			this.Controls.Add(this.Label5);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.txtModel);
			this.Controls.Add(this.txtSerial);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAssemAdd";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "New Assembly";
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		public static int propNumber;
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		// Handle the Form load event.
		private void AssemAdd_Load(object sender, System.EventArgs e)
		{
			FillComboBoxes();
		}
		
		private void FillComboBoxes()
		{
			
			cboSize.ValueMember = "AssemSizeNo";
			cboSize.DisplayMember = "AssemSize";
			cboSize.DataSource = MainForm.ds.Tables["AssemSizes"];
			
		}
		
		private void btnOK_Click(System.Object sender, System.EventArgs e)
		{
			if (FormIsValid())
			{
				AddAssembly();
			}
			else
			{
				DialogResult = DialogResult.None;
			}
		}
		
		public void AddAssembly()
		{
			
			int UsagePriceNo = 1;
			if (cboUsage.Text.Trim() == "Fire Line By-Pass")
			{
				UsagePriceNo = 2;
			}
			
			// Build the INSERT INTO string
			string strInsert = "INSERT INTO Assemblies " + "(propNo, assemSerial, assemType, assemMan, assemMod, assemSizeNo, assemUsage, " + "assemUsagePriceNo, assemInstDt, assemLoc, assemNotes) " + "VALUES (";
			
			strInsert += propNumber + ", ";
			strInsert += "\'" + this.txtSerial.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.cboType.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.cboManufacturer.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.txtModel.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += this.cboSize.SelectedValue.ToString() + ", ";
			strInsert += "\'" + this.cboUsage.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += UsagePriceNo.ToString() + ", ";
			strInsert += "\'" + this.txtInstallDate.Text + "\', ";
			strInsert += "\'" + this.txtLocation.Text.Replace("\'", "\'\'") + "\', ";
			strInsert += "\'" + this.rtbAssemNotes.Text.Replace("\'", "\'\'") + "\')";
			
			
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
				DialogResult = DialogResult.OK;
				
			}
		}
		
		private bool FormIsValid()
		{
			if (SerialIsValid() && ModelIsValid() && UsageIsValid() && InstallDateIsValid() && TypeIsValid() && ManufacturerIsValid())
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		private bool TypeIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(cboType, "");
			if (cboType.Text.Trim() == "" || cboType.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(cboType, "Type is a required field");
			}
			return Booly;
		}
		
		private bool ManufacturerIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(cboManufacturer, "");
			if (cboManufacturer.Text.Trim() == "" || cboManufacturer.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(cboManufacturer, "Manufacturer is a required field");
			}
			return Booly;
		}
		
		private bool SerialIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtSerial, "");
			if (txtSerial.Text.Trim() == "" || txtSerial.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtSerial, "Serial No. is a required field");
			}
			return Booly;
		}
		
		private bool ModelIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtModel, "");
			if (txtModel.Text.Trim() == "" || txtModel.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(txtModel, "Model No. is a required field");
			}
			return Booly;
		}
		
		
		private bool UsageIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(cboUsage, "");
			if (cboUsage.Text.Trim() == "" || cboUsage.Text == null)
			{
				Booly = false;
				ErrorProvider1.SetError(cboUsage, "Usage is a required field.");
			}
			return Booly;
		}
		
		private bool InstallDateIsValid()
		{
			bool Booly = true;
			ErrorProvider1.SetError(txtInstallDate, "");
			
			if (txtInstallDate.Text.Trim() != ""&& txtInstallDate.Text != null)
			{
				Regex InstallDateRegex = new Regex("(?n:^(?=\\d)((?<month>(0?[13578])|1[02]|(0?[469]|11)(?!.31)|0" + "?2(?(.29)(?=.29.((1[6-9]|[2-9]\\d)(0[48]|[2468][048]|[13579][" + "26])|(16|[2468][048]|[3579][26])00))|(?!.3[01])))(?<sep>[-./" + "])(?<day>0?[1-9]|[12]\\d|3[01])\\k<sep>(?<year>(1[6-9]|[2-9]\\d" + ")\\d{2})(?(?=\\x20\\d)\\x20|$))?(?<time>((0?[1-9]|1[012])(:[0-5]" + "\\d){0,2}(?i:\\x20[AP]M))|([01]\\d|2[0-3])(:[0-5]\\d){1,2})?$)");
				
				if (! InstallDateRegex.IsMatch(txtInstallDate.Text))
				{
					Booly = false;
					ErrorProvider1.SetError(txtInstallDate, "Dates must be in the form of \'2/12/2004\'");
				}
				
			}
			
			return Booly;
		}
		
	}
	
}
