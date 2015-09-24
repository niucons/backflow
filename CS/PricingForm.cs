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
	public class PricingForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public PricingForm()
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
		internal System.Windows.Forms.Button btnLeft;
		internal System.Windows.Forms.Button btnRight;
		internal System.Windows.Forms.DataGrid grdScheme;
		internal System.Windows.Forms.Button btnSave;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PricingForm));
			this.grdScheme = new System.Windows.Forms.DataGrid();
			base.Load += new System.EventHandler(frmPricing_Load);
			this.btnLeft = new System.Windows.Forms.Button();
			this.btnLeft.Click += new System.EventHandler(btnLeft_Click);
			this.btnRight = new System.Windows.Forms.Button();
			this.btnRight.Click += new System.EventHandler(btnRight_Click);
			this.btnSave = new System.Windows.Forms.Button();
			this.btnSave.Click += new System.EventHandler(btnSave_Click);
			((System.ComponentModel.ISupportInitialize) this.grdScheme).BeginInit();
			this.SuspendLayout();
			//
			//grdScheme
			//
			this.grdScheme.BackgroundColor = System.Drawing.SystemColors.AppWorkspace;
			this.grdScheme.DataMember = "";
			this.grdScheme.GridLineColor = System.Drawing.SystemColors.Control;
			this.grdScheme.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.grdScheme.Location = new System.Drawing.Point(8, 8);
			this.grdScheme.Name = "grdScheme";
			this.grdScheme.Size = new System.Drawing.Size(176, 288);
			this.grdScheme.TabIndex = 0;
			//
			//btnLeft
			//
			this.btnLeft.Location = new System.Drawing.Point(192, 256);
			this.btnLeft.Name = "btnLeft";
			this.btnLeft.Size = new System.Drawing.Size(30, 24);
			this.btnLeft.TabIndex = 2;
			this.btnLeft.Text = "<";
			//
			//btnRight
			//
			this.btnRight.Location = new System.Drawing.Point(232, 256);
			this.btnRight.Name = "btnRight";
			this.btnRight.Size = new System.Drawing.Size(30, 24);
			this.btnRight.TabIndex = 3;
			this.btnRight.Text = ">";
			//
			//btnSave
			//
			this.btnSave.Location = new System.Drawing.Point(192, 208);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 4;
			this.btnSave.Text = "Save";
			//
			//frmPricing
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(264, 301);
			this.Controls.Add(this.btnSave);
			this.Controls.Add(this.btnRight);
			this.Controls.Add(this.btnLeft);
			this.Controls.Add(this.grdScheme);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmPricing";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Pricing Schemes";
			((System.ComponentModel.ISupportInitialize) this.grdScheme).EndInit();
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		private int PricingScheme = 2;
		private int MaxSchemeNo;
		private DataTable dtSchemes;
		private DataView dvSchemes;
		private SqlDataAdapter sda;
		
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		// Handle the Form load event.
		private void frmPricing_Load(object sender, System.EventArgs e)
		{
			FillPricingDataTable();
			SetMaxSchemeNo();
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
		
		private void FillPricingDataTable()
		{
			SqlConnection cn = new SqlConnection(connectionString);
			string sqlStatement = "SELECT DISTINCT Price.priceNo, PricingSchemes.manSchemeNo, AssemblySizes.AssemSizeNo, " + "AssemblySizes.AssemSize, Price.price, AssemblyUsage.assemUsagePriceNo FROM Price " + "LEFT JOIN AssemblySizes ON AssemblySizes.AssemSizeNo = Price.assemSizeNo " + "LEFT JOIN PricingSchemes ON PricingSchemes.manSchemeNo = Price.manSchemeNo " + "LEFT JOIN AssemblyUsagePrice ON AssemblyUsagePrice.assemUsagePriceNo = Price.assemUsagePriceNo " + "LEFT JOIN AssemblyUsage ON AssemblyUsage.assemUsagePriceNo = AssemblyUsagePrice.assemUsagePriceNo " + "ORDER BY  AssemblyUsage.assemUsagePriceNo, AssemblySizes.AssemSizeNo";
			SqlCommand cmd = new SqlCommand(sqlStatement, cn);
			sda = new SqlDataAdapter(cmd);
			
			
			if (!System.Convert.ToBoolean(MainForm.ds.Tables["PricingSchemes"] == null))
			{
				MainForm.ds.Tables["PricingSchemes"].Clear();
			}
			
			sda.Fill(MainForm.ds, "PricingSchemes");
			
			dtSchemes = MainForm.ds.Tables["PricingSchemes"];
			
			dvSchemes = dtSchemes.DefaultView;
			dvSchemes.AllowDelete = false;
			dvSchemes.AllowEdit = true;
			dvSchemes.AllowNew = false;
			
			BindSchemesGrid();
		}
		
		public void BindSchemesGrid()
		{
			
			// Filter the OrderDetails data based on the currently selected OrderID.
			dvSchemes.RowFilter = "manSchemeNo = " + PricingScheme;
			
			grdScheme.CaptionText = System.Convert.ToInt32("Scheme No. " + PricingScheme) - System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(System.Convert.ToInt32(1.ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString()).ToString();
			grdScheme.DataSource = dvSchemes;
			
			// You must clear out the TableStyles collection before
			grdScheme.TableStyles.Clear();
			
			DataGridTableStyle grdTableStyle1 = new DataGridTableStyle();
			// You must always set the MappingName, even with a DataView that has
			// only one Table. If not, you will get no errors but the formatting
			// will not appear. To avoid mistakes, instead of typing the name of
			// the table used when creating the DataSet, you can access the
			// TableName property.
			grdTableStyle1.MappingName = dvSchemes.Table.TableName;
			
			// The use of column styles overrides the automatic generation of columns
			// for every column in the DataTable. When column style objects are used,
			// every column you want to display has to have an associate column style
			// object.
			DataGridTextBoxColumn grdColStyle1 = new DataGridTextBoxColumn();
			grdColStyle1.MappingName = "AssemSize";
			grdColStyle1.HeaderText = "Size";
			grdColStyle1.Width = 65;
			
			DataGridTextBoxColumn grdColStyle2 = new DataGridTextBoxColumn();
			grdColStyle2.MappingName = "price";
			grdColStyle2.HeaderText = "Price";
			grdColStyle2.Width = 70;
			grdColStyle2.Format = "c";
			
			// Add the column style objects to the table style's collection of
			// column styles. Without this the styles do not take effect.
			grdTableStyle1.GridColumnStyles.AddRange(new DataGridColumnStyle[] { grdColStyle1, grdColStyle2 });
			
			grdScheme.TableStyles.Add(grdTableStyle1);
			
			Binding bn;
			foreach (Binding tempLoopVar_bn in grdScheme.DataBindings)
			{
				bn = tempLoopVar_bn;
				Trace.Write(bn.Control.ToString() + ", " + bn.PropertyName);
				Trace.Write("\r\n");
			}
		}
		
		private void NextScheme()
		{
			PricingScheme++;
			BindSchemesGrid();
		}
		
		private void PreviousScheme()
		{
			PricingScheme--;
			BindSchemesGrid();
		}
		
		private void btnRight_Click(System.Object sender, System.EventArgs e)
		{
			btnLeft.Enabled = true;
			if (PricingScheme < MaxSchemeNo)
			{
				NextScheme();
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
				PreviousScheme();
				if (PricingScheme == 2)
				{
					btnLeft.Enabled = false;
				}
			}
		}
		
		
		
		private void btnSave_Click(System.Object sender, System.EventArgs e)
		{
			int index;
			int lookUp;
			double Price;
			
			SqlConnection cn = new SqlConnection(connectionString);
			//			string strQuery = "SELECT MAX(manSchemeNo) FROM PricingSchemes";
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			
			try
			{
				cn.Open(); // Open Connection
				
				for (index = 0; index <= grdScheme.VisibleRowCount - 1; index++)
				{
					lookUp = System.Convert.ToInt32(dvSchemes[index]["priceNo"]);
					Price = System.Convert.ToDouble(grdScheme[index, 1]);
					cmd.CommandText = "UPDATE Price SET price = " + Price + " WHERE priceNo = " + lookUp;
					cmd.ExecuteNonQuery();
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
			
		}
		
	}
	
}
