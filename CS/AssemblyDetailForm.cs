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
	public class AssemblyDetailForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public AssemblyDetailForm()
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
		internal System.Windows.Forms.TextBox txtAssemNo;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.TextBox txtLocation;
		internal System.Windows.Forms.Label lblNotes;
		internal System.Windows.Forms.RichTextBox rtbAssemNotes;
		internal System.Windows.Forms.ComboBox cboSize;
		internal System.Windows.Forms.ComboBox cboManufacturer;
		internal System.Windows.Forms.ComboBox cboType;
		internal System.Windows.Forms.Button btnOK;
		internal System.Windows.Forms.Button btnCancel;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.TextBox txtInstallDate;
		internal System.Windows.Forms.Label Label10;
		internal System.Windows.Forms.Label Label7;
		internal System.Windows.Forms.Label Label6;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.TextBox txtModel;
		internal System.Windows.Forms.TextBox txtSerial;
		internal System.Windows.Forms.ComboBox cboUsage;
		internal System.Windows.Forms.ErrorProvider ErrorProvider1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(AssemblyDetailForm));
			this.txtAssemNo = new System.Windows.Forms.TextBox();
			base.Load += new System.EventHandler(AssemAdd_Load);
			this.Label2 = new System.Windows.Forms.Label();
			this.txtLocation = new System.Windows.Forms.TextBox();
			this.lblNotes = new System.Windows.Forms.Label();
			this.rtbAssemNotes = new System.Windows.Forms.RichTextBox();
			this.cboSize = new System.Windows.Forms.ComboBox();
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
			this.cboUsage = new System.Windows.Forms.ComboBox();
			this.ErrorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.SuspendLayout();
			//
			//txtAssemNo
			//
			this.txtAssemNo.Location = new System.Drawing.Point(16, 392);
			this.txtAssemNo.Name = "txtAssemNo";
			this.txtAssemNo.Size = new System.Drawing.Size(88, 20);
			this.txtAssemNo.TabIndex = 108;
			this.txtAssemNo.Text = "";
			this.txtAssemNo.Visible = false;
			//
			//Label2
			//
			this.Label2.Location = new System.Drawing.Point(24, 183);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(88, 16);
			this.Label2.TabIndex = 147;
			this.Label2.Text = "Location";
			this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtLocation
			//
			this.txtLocation.HideSelection = false;
			this.txtLocation.Location = new System.Drawing.Point(120, 179);
			this.txtLocation.Name = "txtLocation";
			this.txtLocation.Size = new System.Drawing.Size(248, 20);
			this.txtLocation.TabIndex = 8;
			this.txtLocation.Text = "";
			//
			//lblNotes
			//
			this.lblNotes.Location = new System.Drawing.Point(12, 221);
			this.lblNotes.Name = "lblNotes";
			this.lblNotes.Size = new System.Drawing.Size(144, 16);
			this.lblNotes.TabIndex = 145;
			this.lblNotes.Text = "Assembly Specific Notes";
			//
			//rtbAssemNotes
			//
			this.rtbAssemNotes.Location = new System.Drawing.Point(12, 245);
			this.rtbAssemNotes.Name = "rtbAssemNotes";
			this.rtbAssemNotes.Size = new System.Drawing.Size(352, 128);
			this.rtbAssemNotes.TabIndex = 9;
			this.rtbAssemNotes.Text = "";
			//
			//cboSize
			//
			this.cboSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboSize.Location = new System.Drawing.Point(120, 109);
			this.cboSize.Name = "cboSize";
			this.cboSize.Size = new System.Drawing.Size(152, 21);
			this.cboSize.TabIndex = 5;
			//
			//cboManufacturer
			//
			this.cboManufacturer.Items.AddRange(new object[] { "Ames", "Conbraco", "Febco", "Hersey", "Hunter Ames", "Watts", "Wilkens" });
			this.cboManufacturer.Location = new System.Drawing.Point(120, 62);
			this.cboManufacturer.Name = "cboManufacturer";
			this.cboManufacturer.Size = new System.Drawing.Size(152, 21);
			this.cboManufacturer.TabIndex = 3;
			//
			//cboType
			//
			this.cboType.Items.AddRange(new object[] { "RP", "RPDA", "DC", "DCDA" });
			this.cboType.Location = new System.Drawing.Point(120, 38);
			this.cboType.Name = "cboType";
			this.cboType.Size = new System.Drawing.Size(152, 21);
			this.cboType.TabIndex = 2;
			//
			//btnOK
			//
			this.btnOK.Location = new System.Drawing.Point(308, 389);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(56, 24);
			this.btnOK.TabIndex = 10;
			this.btnOK.Text = "OK";
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(244, 389);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(56, 24);
			this.btnCancel.TabIndex = 139;
			this.btnCancel.Text = "Cancel";
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(24, 160);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(88, 16);
			this.Label1.TabIndex = 138;
			this.Label1.Text = "Install Date";
			this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtInstallDate
			//
			this.txtInstallDate.HideSelection = false;
			this.txtInstallDate.Location = new System.Drawing.Point(120, 156);
			this.txtInstallDate.Name = "txtInstallDate";
			this.txtInstallDate.Size = new System.Drawing.Size(152, 20);
			this.txtInstallDate.TabIndex = 7;
			this.txtInstallDate.Text = "";
			//
			//Label10
			//
			this.Label10.Location = new System.Drawing.Point(16, 67);
			this.Label10.Name = "Label10";
			this.Label10.Size = new System.Drawing.Size(96, 16);
			this.Label10.TabIndex = 136;
			this.Label10.Text = "Manufacturer";
			this.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label7
			//
			this.Label7.Location = new System.Drawing.Point(24, 137);
			this.Label7.Name = "Label7";
			this.Label7.Size = new System.Drawing.Size(88, 16);
			this.Label7.TabIndex = 135;
			this.Label7.Text = "Usage";
			this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label6
			//
			this.Label6.Location = new System.Drawing.Point(24, 114);
			this.Label6.Name = "Label6";
			this.Label6.Size = new System.Drawing.Size(88, 16);
			this.Label6.TabIndex = 134;
			this.Label6.Text = "Size";
			this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label5
			//
			this.Label5.Location = new System.Drawing.Point(24, 90);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(88, 16);
			this.Label5.TabIndex = 133;
			this.Label5.Text = "Model";
			this.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label4
			//
			this.Label4.Location = new System.Drawing.Point(24, 43);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(88, 16);
			this.Label4.TabIndex = 132;
			this.Label4.Text = "Type";
			this.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label3
			//
			this.Label3.Location = new System.Drawing.Point(24, 19);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(88, 16);
			this.Label3.TabIndex = 131;
			this.Label3.Text = "Serial";
			this.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//txtModel
			//
			this.txtModel.HideSelection = false;
			this.txtModel.Location = new System.Drawing.Point(120, 86);
			this.txtModel.Name = "txtModel";
			this.txtModel.Size = new System.Drawing.Size(152, 20);
			this.txtModel.TabIndex = 4;
			this.txtModel.Text = "";
			//
			//txtSerial
			//
			this.txtSerial.Location = new System.Drawing.Point(120, 15);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(152, 20);
			this.txtSerial.TabIndex = 1;
			this.txtSerial.Text = "";
			//
			//cboUsage
			//
			this.cboUsage.Items.AddRange(new object[] { "Boiler Feed", "Car Wash", "Filter System", "Fire Line", "Fire Line By-Pass", "Hose Bib", "Lawn Sprinkler", "Potable" });
			this.cboUsage.Location = new System.Drawing.Point(120, 133);
			this.cboUsage.Name = "cboUsage";
			this.cboUsage.Size = new System.Drawing.Size(152, 21);
			this.cboUsage.TabIndex = 6;
			//
			//ErrorProvider1
			//
			this.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
			this.ErrorProvider1.ContainerControl = this;
			//
			//frmAssemDetail
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 429);
			this.Controls.Add(this.cboUsage);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.txtLocation);
			this.Controls.Add(this.txtInstallDate);
			this.Controls.Add(this.txtModel);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.txtAssemNo);
			this.Controls.Add(this.lblNotes);
			this.Controls.Add(this.rtbAssemNotes);
			this.Controls.Add(this.cboSize);
			this.Controls.Add(this.cboManufacturer);
			this.Controls.Add(this.cboType);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.Label10);
			this.Controls.Add(this.Label7);
			this.Controls.Add(this.Label6);
			this.Controls.Add(this.Label5);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.Label3);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAssemDetail";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Assembly Details";
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		public static int assemblyNumber;
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		private string strSerial;
		private string Type;
		private string Manufacturer;
		private string strModel;
		private int intSize;
		private string strUsage;
		private string strInstallDate;
		private string strLocation;
		private string strNotes;
		
		
		// Handle the Form load event.
		private void AssemAdd_Load(object sender, System.EventArgs e)
		{
			FillComboBoxes();
			FillAssemData();
		}
		
		private void FillComboBoxes()
		{
			
			cboSize.ValueMember = "AssemSizeNo";
			cboSize.DisplayMember = "AssemSize";
			cboSize.DataSource = MainForm.ds.Tables["AssemSizes"];
			
		}
		
		private void FillAssemData()
		{
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(("SELECT * FROM Assemblies WHERE assemNo = " + this.txtAssemNo.Text), cn);
			// Run the query; get the DataReader object.
			cn.Open();
			SqlDataReader dr = cmd.ExecuteReader();
			
			while (dr.Read())
			{
				strSerial = dr["assemSerial"].ToString() + "";
				Type = dr["assemType"].ToString() + "";
				Manufacturer = dr["assemMan"].ToString() + "";
				strModel = dr["assemMod"].ToString() + "";
				if (! Information.IsDBNull(dr["assemSizeNo"]))
				{
					intSize = System.Convert.ToInt32(dr["assemSizeNo"]);
				}
				else
				{
					intSize = 1;
				}
				
				strUsage = dr["assemUsage"].ToString() + "";
				if (Information.IsDBNull(dr["assemInstDt"]))
				{
					strInstallDate = "";
				}
				else if (dr["assemInstDt"].ToString() == "1/1/1900")
				{
					strInstallDate = "";
				}
				else
				{
					strInstallDate = dr["assemInstDt"].ToString() + "";
				}
				
				strLocation = dr["assemLoc"].ToString() + "";
				strNotes = dr["assemNotes"].ToString() + "";
			}
			dr.Close();
			cn.Close();
			
			this.txtSerial.Text = strSerial;
			this.cboType.Text = Type;
			this.cboManufacturer.Text = Manufacturer;
			this.txtModel.Text = strModel;
			this.cboSize.SelectedValue = intSize;
			this.cboUsage.Text = strUsage;
			this.txtInstallDate.Text = strInstallDate;
			this.txtLocation.Text = strLocation;
			this.rtbAssemNotes.Text = strNotes;
			
		} //FillData
		
		public void UpdateAssembly()
		{
			
			int UsagePriceNo = 1;
			//  If cboUsage.SelectedValue = 8 Then
			// UsagePriceNo = 2
			//End If
			
			// Build the INSERT INTO string
			string strUpdate = "UPDATE Assemblies SET ";
			
			strUpdate += "assemSerial = \'" + txtSerial.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "assemType = \'" + cboType.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "assemMan = \'" + cboManufacturer.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "assemMod = \'" + txtModel.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "assemSizeNo = " + cboSize.SelectedValue.ToString() + ", ";
			strUpdate += "assemUsage = \'" + cboUsage.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "assemUsagePriceNo = " + UsagePriceNo + ", ";
			strUpdate += "assemInstDt = \'" + txtInstallDate.Text + "\', ";
			strUpdate += "assemLoc = \'" + txtLocation.Text.Replace("\'", "\'\'") + "\', ";
			strUpdate += "assemNotes = \'" + rtbAssemNotes.Text.Replace("\'", "\'\'") + "\' ";
			strUpdate += "WHERE assemNo = " + txtAssemNo.Text;
			
			
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
				DialogResult = DialogResult.OK;
				
			}
			
		}
		
		private void btnOK_Click(System.Object sender, System.EventArgs e)
		{
			if (FormIsValid())
			{
				UpdateAssembly();
			}
			else
			{
				DialogResult = DialogResult.None;
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
