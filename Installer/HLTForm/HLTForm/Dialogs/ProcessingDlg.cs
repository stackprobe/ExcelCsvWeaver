using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HLTStudio.Dialogs
{
	public partial class ProcessingDlg : Form
	{
		#region WndProc

		[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		protected override void WndProc(ref Message m)
		{
			const int WM_SYSCOMMAND = 0x112;
			const long SC_CLOSE = 0xF060L;
			const long WP_MASK = 0xFFF0L;

			if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & WP_MASK) == SC_CLOSE)
				return;

			base.WndProc(ref m);
		}

		#endregion

		public static void Run(string nameOfProcess, Action routine)
		{
			using (ProcessingDlg f = new ProcessingDlg())
			{
				f.P_Routine = routine;

				f.MainMessage.Text = $"「{Consts.APPLICATION_LONG_NAME}」を{nameOfProcess}しています...";
				f.ShowDialog();

				if (f.P_Ex != null)
				{
					throw new Exception($"{nameOfProcess}中にエラーが発生しました。", f.P_Ex);
				}
			}
		}

		private Action P_Routine;
		private Exception P_Ex = null;

		private ProcessingDlg()
		{
			InitializeComponent();
		}

		private void ProcessingDlg_Load(object sender, EventArgs e)
		{
			this.MinimumSize = this.Size;
		}

		private void ProcessingDlg_Shown(object sender, EventArgs e)
		{
			new Thread(() =>
			{
				Thread delayTh = new Thread(() =>
				{
					Thread.Sleep(1000);
				});

				delayTh.Start();

				try
				{
					P_Routine();
				}
				catch (Exception ex)
				{
					P_Ex = ex;
				}

				delayTh.Join();

				this.BeginInvoke((MethodInvoker)delegate
				{
					this.Close();
				});
			})
			.Start();
		}
	}
}
