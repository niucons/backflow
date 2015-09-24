using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;

namespace bftm
{
	public class myListViewItem : ListViewItem
	{
		
		
		// Public FilePath As String
		public string strName;
		public string strType;
		public int intID;
		public bool boolActive;
		
		public myListViewItem(string Name) : base(Name)
		{
		}
		
		public myListViewItem(string Name, string Type, int ID, int imageIndex) : base(Name, imageIndex)
		{
			
			base.Tag = Type;
			intID = ID;
			
		}
		
		public myListViewItem(string Name, string Type, int ID)
		{
			strName = Name;
			base.Tag = Type;
			intID = ID;
			this.Text = Name;
			
		}
		
	}
	
}
