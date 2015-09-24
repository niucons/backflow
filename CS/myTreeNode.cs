using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;

namespace bftm
{
	public class myTreeNode : TreeNode
	{
		
		
		int m_ID;
		bool m_Active;
		
		public myTreeNode(string Name) : base(Name)
		{
		}
		
		public myTreeNode(string Name, string Type, int ID, bool Active) : base(Name)
		{
			base.Tag = Type;
			m_ID = ID;
			m_Active = Active;
		}
		
		public myTreeNode(string Name, string Type, int ID, bool Active, int ImageIndex, int SelectedImageIndex) : base(Name, ImageIndex, SelectedImageIndex)
		{
			base.Tag = Type;
			m_ID = ID;
			m_Active = Active;
		}
		
		public int ID
		{
			get
			{
				return m_ID;
			}
			set
			{
				m_ID = value;
			}
		}
		
		public bool isActive
		{
			get
			{
				return m_Active;
			}
			set
			{
				m_Active = value;
			}
		}
		
	}
	
	
}
