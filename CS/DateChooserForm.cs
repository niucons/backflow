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
	public class DateChooserForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public DateChooserForm()
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
		internal System.Windows.Forms.DateTimePicker dtpStart;
		internal System.Windows.Forms.DateTimePicker dtpEnd;
		internal System.Windows.Forms.Button btnOK;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Button btnCancel;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(DateChooserForm));
			this.dtpStart = new System.Windows.Forms.DateTimePicker();
			this.dtpEnd = new System.Windows.Forms.DateTimePicker();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnOK.Click += new System.EventHandler(btnOK_Click);
			this.Label1 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.btnCancel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			//
			//dtpStart
			//
			this.dtpStart.CalendarForeColor = System.Drawing.SystemColors.GrayText;
			this.dtpStart.CalendarTitleBackColor = System.Drawing.Color.LightSlateGray;
			this.dtpStart.Location = new System.Drawing.Point(96, 24);
			this.dtpStart.Name = "dtpStart";
			this.dtpStart.Size = new System.Drawing.Size(232, 20);
			this.dtpStart.TabIndex = 0;
			//
			//dtpEnd
			//
			this.dtpEnd.CalendarForeColor = System.Drawing.SystemColors.GrayText;
			this.dtpEnd.CalendarTitleBackColor = System.Drawing.Color.LightSlateGray;
			this.dtpEnd.Location = new System.Drawing.Point(96, 72);
			this.dtpEnd.Name = "dtpEnd";
			this.dtpEnd.Size = new System.Drawing.Size(232, 20);
			this.dtpEnd.TabIndex = 1;
			//
			//btnOK
			//
			this.btnOK.Location = new System.Drawing.Point(128, 120);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(56, 24);
			this.btnOK.TabIndex = 3;
			this.btnOK.Text = "OK";
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(24, 24);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(64, 24);
			this.Label1.TabIndex = 4;
			this.Label1.Text = "Start Date:";
			this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//Label2
			//
			this.Label2.Location = new System.Drawing.Point(16, 72);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(72, 24);
			this.Label2.TabIndex = 5;
			this.Label2.Text = "Finish Date:";
			this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(200, 120);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(56, 24);
			this.btnCancel.TabIndex = 6;
			this.btnCancel.Text = "Cancel";
			//
			//frmDateChsr
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 157);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.dtpEnd);
			this.Controls.Add(this.dtpStart);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDateChsr";
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Date Chooser";
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		private void btnOK_Click(System.Object sender, System.EventArgs e)
		{
			DateTime StartDate = dtpStart.Value.Date;
			DateTime EndDate = dtpEnd.Value.Date;
			
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand();
			cmd.Connection = cn;
			cmd.CommandText = "DELETE FROM Dates WHERE Record = 1";
			
			try
			{
				cn.Open();
				cmd.ExecuteNonQuery();
				cmd.CommandText = "INSERT INTO Dates (Record, StartDate, FinishDate, SingleDate) VALUES (1,\'" + StartDate + "\',\'" + EndDate + "\',\'" + EndDate + "\')";
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
			
			this.DialogResult = DialogResult.OK;
		}
	}
	
}
