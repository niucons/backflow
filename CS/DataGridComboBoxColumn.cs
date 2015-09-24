using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;
using System.ComponentModel;
using System.Data.Common;
using System.Data.OleDb;




// Step 1. Derive a custom column style from DataGridTextBoxColumn
//	a) add a ComboBox member
//  b) track when the combobox has focus in Enter and Leave events
//  c) override Edit to allow the ComboBox to replace the TextBox
//  d) override Commit to save the changed data
namespace bftm
{
	public class DataGridComboBoxColumn : DataGridTextBoxColumn
	{
		
		public NoKeyUpCombo ColumnComboBox;
		private System.Windows.Forms.CurrencyManager _source;
		private int _rowNum;
		private bool _isEditing;
		public static int RowCnt;
		
		
		public DataGridComboBoxColumn()
		{
			_source = null;
			_isEditing = false;
			RowCnt = - 1;
			
			ColumnComboBox = new NoKeyUpCombo();
			ColumnComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
			
			ColumnComboBox.Leave += new System.EventHandler(LeaveComboBox);
			ColumnComboBox.SelectionChangeCommitted += new System.EventHandler(ComboStartEditing);
		} //New
		
		private void HandleScroll(object sender, EventArgs e)
		{
			if (ColumnComboBox.Visible)
			{
				ColumnComboBox.Hide();
			}
		} //HandleScroll
		
		private void ComboStartEditing(object sender, EventArgs e)
		{
			_isEditing = true;
			base.ColumnStartedEditing((System.Windows.Forms.Control) sender);
		} //ComboMadeCurrent
		
		
		private void LeaveComboBox(object sender, EventArgs e)
		{
			if (_isEditing)
			{
				SetColumnValueAtRow(_source, _rowNum, ColumnComboBox.Text);
				_isEditing = false;
				Invalidate();
				
			}
			ColumnComboBox.Hide();
			this.DataGridTableStyle.DataGrid.Scroll += new System.EventHandler(new EventHandler( HandleScroll));
		} //LeaveComboBox
		
		
		protected override void Edit(System.Windows.Forms.CurrencyManager @source, int rowNum, System.Drawing.Rectangle bounds, bool @readOnly, string instantText, bool cellIsVisible)
		{
			
			base.Edit(@source, rowNum, bounds, @readOnly, instantText, cellIsVisible);
			
			_rowNum = rowNum;
			_source = @source;
			
			ColumnComboBox.Parent = this.TextBox.Parent;
			ColumnComboBox.Location = this.TextBox.Location;
			ColumnComboBox.Size = new Size(this.TextBox.Size.Width, ColumnComboBox.Size.Height);
			ColumnComboBox.SelectedIndex = ColumnComboBox.FindStringExact(this.TextBox.Text);
			ColumnComboBox.Text = this.TextBox.Text;
			this.TextBox.Visible = false;
			ColumnComboBox.Visible = true;
			this.DataGridTableStyle.DataGrid.Scroll += new System.EventHandler(HandleScroll);
			
			ColumnComboBox.BringToFront();
			ColumnComboBox.Focus();
		} //Edit
		
		
		protected override bool Commit(System.Windows.Forms.CurrencyManager dataSource, int rowNum)
		{
			
			if (_isEditing)
			{
				_isEditing = false;
				SetColumnValueAtRow(dataSource, rowNum, ColumnComboBox.Text);
			}
			return true;
		} //Commit
		
		
		protected override void ConcedeFocus()
		{
			Console.WriteLine("ConcedeFocus");
			base.ConcedeFocus();
		} //ConcedeFocus
		
		protected override object GetColumnValueAtRow(System.Windows.Forms.CurrencyManager @source, int rowNum)
		{
			
			object s = base.GetColumnValueAtRow(@source, rowNum);
			DataView dv = (DataView) this.ColumnComboBox.DataSource;
			int rowCount = dv.Count;
			int i = 0;
			object s1;
			
			//if things are slow, you could order your dataview
			//& use binary search instead of this linear one
			while (i < rowCount)
			{
				s1 = dv[i][this.ColumnComboBox.ValueMember];
				if ((!System.Convert.ToBoolean(s1 == DBNull.Value)) &&(!System.Convert.ToBoolean(s == DBNull.Value)) && s == s1)
				{
					break;
				}
				i++;
			}
			
			if (i < rowCount)
			{
				return dv[i][this.ColumnComboBox.DisplayMember];
			}
			return DBNull.Value;
		} //GetColumnValueAtRow
		
		
		protected override void SetColumnValueAtRow(System.Windows.Forms.CurrencyManager @source, int rowNum, object value)
		{
			object s = value;
			
			DataView dv = (DataView) this.ColumnComboBox.DataSource;
			int rowCount = dv.Count;
			int i = 0;
			object s1;
			
			//if things are slow, you could order your dataview
			//& use binary search instead of this linear one
			while (i < rowCount)
			{
				s1 = dv[i][this.ColumnComboBox.DisplayMember];
				if ((!System.Convert.ToBoolean(s1 == DBNull.Value)) && s == s1)
				{
					break;
				}
				i++;
			}
			if (i < rowCount)
			{
				s = dv[i][this.ColumnComboBox.ValueMember];
			}
			else
			{
				s = DBNull.Value;
			}
			base.SetColumnValueAtRow(@source, rowNum, s);
		} //SetColumnValueAtRow
		
		
	} //DataGridComboBoxColumn
	
	
	
}
