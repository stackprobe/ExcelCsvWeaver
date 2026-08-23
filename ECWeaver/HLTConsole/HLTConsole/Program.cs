using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using HLTStudio.ECWOperations;
using HLTStudio.ECWArguments;
using HLTStudio.Commons;
using HLTStudio.Tests;

namespace HLTStudio
{
	class Program
	{
		static void Main(string[] args)
		{
			ProcMain.CUIMain(new Program().Main2);
		}

		private void Main2(ArgsReader ar)
		{
			if (ProcMain.DEBUG && !ar.HasArgs())
			{
				Main3();
			}
			else
			{
				Main4(ar);
			}
			SCommon.OpenOutputDirIfCreated();
		}

		private void Main3()
		{
#if DEBUG
			// -- choose one --

			//Main4(new ArgsReader(new string[] { }));
			//Main4(new ArgsReader(new string[] { }));
			//Main4(new ArgsReader(new string[] { }));
			new Test0001().Test01();

			// --
#endif
			SCommon.Pause();
		}

		private void Main4(ArgsReader ar)
		{
			try
			{
				Main5(ar);
			}
			catch (Exception ex)
			{
				ProcMain.WriteLog(ex);

				MessageBox.Show(ex.ToString(), $"{Path.GetFileNameWithoutExtension(ProcMain.SelfFile)} / エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Main5(ArgsReader ar)
		{
			ECWeaverArgs args = ECWeaverArgs.Read(ar);
			new ECWeaverProcessor().Run(args);
		}
	}
}
