using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using HLTStudio.Commons;
using IWshRuntimeLibrary;

namespace HLTStudio.Modules
{
	// memo: 参照の追加 -> COM -> Windows Script Host Object Model

	public static class ShortcutCreator
	{
		public static string GetShortcutPath()
		{
			// デスクトップパスを取得
			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

			// ショートカットファイルのパス
			string shortcutPath = Path.Combine(desktopPath, Consts.APPLICATION_NAME + ".lnk");
			//shortcutPath = SCommon.ToCreatablePath(shortcutPath);

			return shortcutPath;
		}

		public static void Run(string mainProgram)
		{
			// ショートカットファイルのパス
			string shortcutPath = GetShortcutPath();

			// 既存削除
			SCommon.DeletePath(shortcutPath);

			// WScript.Shell を生成
			var shell = new WshShell();

			// ショートカットを作成
			IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

			// 実行するターゲット（例：アプリ本体の実行ファイル）
			shortcut.TargetPath = mainProgram;

			// 作業フォルダ（省略可）
			shortcut.WorkingDirectory = SCommon.ToParentPath(mainProgram);

			// アイコン（省略可）
			shortcut.IconLocation = mainProgram + ",0";

			// コメント（省略可）
			shortcut.Description = Consts.APPLICATION_NAME;

			// 保存
			shortcut.Save();

			// 後始末
			Marshal.FinalReleaseComObject(shortcut);
			Marshal.FinalReleaseComObject(shell);
		}
	}
}
