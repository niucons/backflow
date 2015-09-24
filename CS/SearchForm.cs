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
	public class SearchForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public SearchForm()
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
		internal System.Windows.Forms.Button btnSearch;
		internal System.Windows.Forms.ListView ListView1;
		internal System.Windows.Forms.ProgressBar ProgressBar1;
		internal System.Windows.Forms.TextBox txtSearch;
		internal System.Windows.Forms.ImageList ImageList1;
		internal System.Windows.Forms.Label Label1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(SearchForm));
			this.txtSearch = new System.Windows.Forms.TextBox();
			this.btnSearch = new System.Windows.Forms.Button();
			this.btnSearch.Click += new System.EventHandler(btnSearch_Click);
			this.ListView1 = new System.Windows.Forms.ListView();
			this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
			this.ProgressBar1 = new System.Windows.Forms.ProgressBar();
			this.Label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			//
			//txtSearch
			//
			this.txtSearch.Location = new System.Drawing.Point(8, 30);
			this.txtSearch.Name = "txtSearch";
			this.txtSearch.Size = new System.Drawing.Size(328, 20);
			this.txtSearch.TabIndex = 0;
			this.txtSearch.Text = "";
			//
			//btnSearch
			//
			this.btnSearch.Image = (System.Drawing.Image) resources.GetObject("btnSearch.Image");
			this.btnSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSearch.Location = new System.Drawing.Point(352, 24);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(64, 32);
			this.btnSearch.TabIndex = 1;
			this.btnSearch.Text = "Search ";
			this.btnSearch.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//ListView1
			//
			this.ListView1.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right);
			this.ListView1.Location = new System.Drawing.Point(8, 64);
			this.ListView1.MultiSelect = false;
			this.ListView1.Name = "ListView1";
			this.ListView1.Size = new System.Drawing.Size(592, 176);
			this.ListView1.SmallImageList = this.ImageList1;
			this.ListView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
			this.ListView1.TabIndex = 2;
			this.ListView1.View = System.Windows.Forms.View.Details;
			//
			//ImageList1
			//
			this.ImageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.ImageList1.ImageStream = (System.Windows.Forms.ImageListStreamer) resources.GetObject("ImageList1.ImageStream");
			this.ImageList1.TransparentColor = System.Drawing.Color.Transparent;
			//
			//ProgressBar1
			//
			this.ProgressBar1.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.ProgressBar1.Location = new System.Drawing.Point(456, 253);
			this.ProgressBar1.Name = "ProgressBar1";
			this.ProgressBar1.Size = new System.Drawing.Size(144, 16);
			this.ProgressBar1.TabIndex = 3;
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(8, 8);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(328, 16);
			this.Label1.TabIndex = 4;
			this.Label1.Text = "Enter the piece of information that you would lilke to search by.";
			//
			//frmSearch
			//
			this.AcceptButton = this.btnSearch;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(608, 277);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.ProgressBar1);
			this.Controls.Add(this.ListView1);
			this.Controls.Add(this.btnSearch);
			this.Controls.Add(this.txtSearch);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.Name = "frmSearch";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Search";
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		
		private void Search()
		{
			ListView1.Clear();
			ProgressBar1.Value = 15;
			
			// Set to details view
			ColumnHeader columnheader; // Used for creating column headers.
			ListViewItem listviewitem; // Used for creating ListView items.
			
			// Make sure that the view is set to show details.
			ListView1.View = View.Details;
			
			int intManNo;
			string strManager;
			string strProperty;
			int intPropNo;
			int intAssemNo;
			string strAssembly;
			
			SqlConnection cn = new SqlConnection(connectionString);
			string sql = "SELECT manNo, manName FROM Managers WHERE ( manName LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR manStrtAdd LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR manSuite LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR manCity LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR manZip LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR manCntct LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + ") AND manDeleted = 0";
			SqlCommand cmd = new SqlCommand(sql, cn);
			ProgressBar1.Value = 30;
			try
			{
				cn.Open();
				// Run the query; get the DataReader object.
				SqlDataReader dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					
					strManager = dr["manName"].ToString() + "";
					intManNo = System.Convert.ToInt32(dr["manNo"]);
					
					// Create some ListView items consisting of first and last names.
					listviewitem = new myListViewItem(strManager, "Manager", intManNo, 0);
					listviewitem.SubItems.Add(strManager);
					this.ListView1.Items.Add(listviewitem);
					
				}
				dr.Close();
				ProgressBar1.Value = 45;
				
				
				cmd.CommandText = "SELECT propNo, propName, storeNo, manName FROM Properties " + "LEFT JOIN Managers ON Managers.manNo = Properties.manNo WHERE ( " + "propName LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR storeNo LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR propCon LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR propStrt LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR propCity LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR propZip LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + ") AND propDeleted = 0";
				
				dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					
					intPropNo = System.Convert.ToInt32(dr["propNo"]);
					// strProperty = dr.Item("propName") & ""
					strManager = dr["manName"].ToString() + "";
					
					if (Information.IsDBNull(dr["storeNo"]))
					{
						strProperty = dr["propName"].ToString() + "";
					}
					else if (dr["storeNo"].ToString() == "")
					{
						strProperty = dr["propName"].ToString() + "";
					}
					else
					{
						strProperty = dr["propName"].ToString() + " - " + dr["storeNo"].ToString();
					}
					
					// Create some ListView items consisting of first and last names.
					listviewitem = new myListViewItem(strProperty, "Property", intPropNo, 1);
					listviewitem.SubItems.Add(strManager + " \\ " + strProperty);
					this.ListView1.Items.Add(listviewitem);
					
				}
				dr.Close();
				ProgressBar1.Value = 60;
				
				cmd.CommandText = "SELECT assemNo, assemSerial, manName, propName, storeNo FROM Assemblies " + "LEFT JOIN Properties ON Properties.propNo = Assemblies.propNo " + "LEFT JOIN Managers ON Managers.manNo = Properties.manNo WHERE ( " + "assemSerial LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + "OR assemMod LIKE \'%" + txtSearch.Text.Replace("\'", "\'\'") + "%\' " + ") AND assemDeleted = 0";
				
				dr = cmd.ExecuteReader();
				
				while (dr.Read())
				{
					
					intAssemNo = System.Convert.ToInt32(dr["assemNo"]);
					strAssembly = dr["assemSerial"].ToString() + "";
					// strProperty = dr.Item("propName") & ""
					strManager = dr["manName"].ToString() + "";
					
					if (Information.IsDBNull(dr["storeNo"]))
					{
						strProperty = dr["propName"].ToString() + "";
					}
					else if (dr["storeNo"].ToString() == "")
					{
						strProperty = dr["propName"].ToString() + "";
					}
					else
					{
						strProperty = dr["propName"].ToString() + " - " + dr["storeNo"].ToString();
					}
					
					// Create some ListView items consisting of first and last names.
					listviewitem = new myListViewItem(strAssembly, "Assembly", intAssemNo, 2);
					listviewitem.SubItems.Add(strManager + " \\ " + strProperty + " \\ " + strAssembly);
					this.ListView1.Items.Add(listviewitem);
					
				}
				dr.Close();
				ProgressBar1.Value = 75;
				
				// Create some column headers for the data.
				columnheader = new ColumnHeader();
				columnheader.Text = "Item";
				this.ListView1.Columns.Add(columnheader);
				
				columnheader = new ColumnHeader();
				columnheader.Text = "Location";
				this.ListView1.Columns.Add(columnheader);
				
				// Loop through and size each column header to fit the column header text.
				foreach (ColumnHeader tempLoopVar_columnheader in this.ListView1.Columns)
				{
					columnheader = tempLoopVar_columnheader;
					columnheader.Width = - 2;
				}
				
				ProgressBar1.Value = 100;
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
				ProgressBar1.Value = 0;
			}
			
		}
		
		private void btnSearch_Click(System.Object sender, System.EventArgs e)
		{
			string str;
			str = txtSearch.Text;
			if (str != null)
			{
				Search();
			}
			
		}
		
		private void Button1_Click(System.Object sender, System.EventArgs e)
		{
			ProgressBar1.Value = 50;
		}
		
	}
	
}
