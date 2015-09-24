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



namespace bftm
{
	public class NoKeyUpCombo : ComboBox
	{
		
		private int WM_KEYUP = 0x101;
		
		
		protected override void WndProc(ref System.Windows.Forms.Message m)
		{
			if (m.Msg == WM_KEYUP)
			{
				//ignore keyup to avoid problem with tabbing & dropdownlist;
				return;
			}
			object temp_object = m;
			System.Windows.Forms.Message temp_SystemWindowsFormsMessage =  (System.Windows.Forms.Message) temp_object;
			base.WndProc(ref temp_SystemWindowsFormsMessage);
		} //WndProc
	} //NoKeyUpCombo
	
	
}
