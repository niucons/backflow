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
	public class ReportsForm : System.Windows.Forms.Form
	{
		
		#region " Windows Form Designer generated code "
		
		public ReportsForm()
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
		internal CrystalDecisions.Windows.Forms.CrystalReportViewer CrystalReportViewer1;
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ReportsForm));
			this.CrystalReportViewer1 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
			this.SuspendLayout();
			//
			//CrystalReportViewer1
			//
			this.CrystalReportViewer1.ActiveViewIndex = - 1;
			this.CrystalReportViewer1.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right);
			this.CrystalReportViewer1.Location = new System.Drawing.Point(8, 8);
			this.CrystalReportViewer1.Name = "CrystalReportViewer1";
			this.CrystalReportViewer1.ReportSource = null;
			this.CrystalReportViewer1.Size = new System.Drawing.Size(1000, 520);
			this.CrystalReportViewer1.TabIndex = 0;
			//
			//frmReports
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1012, 533);
			this.Controls.Add(this.CrystalReportViewer1);
			this.Icon = (System.Drawing.Icon) resources.GetObject("$this.Icon");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReports";
			this.ShowInTaskbar = false;
			this.Text = "Backflow Testing Reporting";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.ResumeLayout(false);
			
		}
		
		#endregion
		
		public string ReportType;
		
		protected string connectionString = "User Id=sa;Data Source=(local)\\VSDotNet;" + "Password=noName#2; Initial Catalog=qb; Integrated Security=SSPI";
		
		protected override void OnLoad(System.EventArgs e)
		{
			if (ReportType == "ManagerList")
			{
				ManagerCallSheet report = new ManagerCallSheet();
				CrystalReportViewer1.ReportSource = report;
			}
			if (ReportType == "MunicipalityList")
			{
				Municipalities report = new Municipalities();
				CrystalReportViewer1.ReportSource = report;
			}
			if (ReportType == "RetestList")
			{
				RetestList report = new RetestList();
				CrystalReportViewer1.ReportSource = report;
			}
			if (ReportType == "RetestLetters")
			{
				RetestLetters report = new RetestLetters();
				CrystalReportViewer1.ReportSource = report;
			}
			
			
		}
		
	}
	
}
