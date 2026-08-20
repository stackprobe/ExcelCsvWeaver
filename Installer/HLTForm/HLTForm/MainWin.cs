using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HLTStudio.Commons;
using HLTStudio.Dialogs;
using HLTStudio.Modules;
using HLTStudio.Tools;

namespace HLTStudio
{
	public partial class MainWin : Form
	{
		public MainWin()
		{
			InitializeComponent();
		}

		private void MainWin_Load(object sender, EventArgs e)
		{
			this.MinimumSize = this.Size;

			this.Text = $"{Consts.APPLICATION_NAME} インストーラ";
			this.LMainMessage.Text = $"「{Consts.APPLICATION_LONG_NAME}」をインストールします。";
			this.TxtInstallDir.Text = Consts.DEFAULT_INSTALL_DIR;
		}

		private void MainWin_Shown(object sender, EventArgs e)
		{
			this.RefreshView();
			this.SetInitialFocus();
		}

		private void RefreshView()
		{
			bool alreadyInstalled = File.Exists(Path.Combine(this.TxtInstallDir.Text, Consts.INSTALLED_SIGNATURE));

			if (alreadyInstalled)
			{
				this.BtnInstall.Text = "再インストール";
				this.BtnUninstall.Visible = true;
			}
			else
			{
				this.BtnInstall.Text = "インストール";
				this.BtnUninstall.Visible = false;
			}
		}

		private void SetInitialFocus()
		{
			//this.BtnBrowse.Focus();
			this.BtnResetInstallDir.Focus();
		}

		private void BtnBrowse_Click(object sender, EventArgs e)
		{
			try
			{
				string homeDir = Directory.GetCurrentDirectory();
				try
				{
					string initialSelectedPath = this.TxtInstallDir.Text;

					// アプリケーション名を除去する。
					initialSelectedPath = SCommon.ToParentPath(initialSelectedPath);

					// 未作成の場合は更に遡る。
					while (
						!SCommon.IsAbsRootDir(initialSelectedPath) &&
						!Directory.Exists(initialSelectedPath)
						)
						initialSelectedPath = SCommon.ToParentPath(initialSelectedPath);

					using (FolderBrowserDialog fbd = new FolderBrowserDialog())
					{
						const bool 新しいフォルダの作成を許可するか_Flag = true;

						fbd.Description =
							"インストール先フォルダを変更します。\r\n" +
							"※ルートディレクトリ・ネットワークディレクトリは選択できません。\r\n" +
							"※ルートディレクトリを選択すると、アプリケーション名を付与します。";
						fbd.RootFolder = Environment.SpecialFolder.Desktop;
						fbd.SelectedPath = initialSelectedPath;
						fbd.ShowNewFolderButton = 新しいフォルダの作成を許可するか_Flag;

						if (fbd.ShowDialog() == DialogResult.OK)
						{
							string installDir = SCommon.MakeFullPath(fbd.SelectedPath);

							// ルートディレクトリ回避
							if (SCommon.IsAbsRootDir(installDir))
								installDir = Path.Combine(installDir, Consts.APPLICATION_NAME);

							this.TxtInstallDir.Text = installDir;

							this.RefreshView();
						}
					}
				}
				finally
				{
					Directory.SetCurrentDirectory(homeDir);
				}
			}
			catch (Exception ex)
			{
				MessageDlg.Run(
					MessageDlg.Kind_e.Error,
					"エラー",
					"インストール先フォルダの変更に失敗しました。\r\n"
					+ "原因：" + GetInnermostException(ex).Message,
					ex,
					new string[] { "OK" }
					);
			}
		}

		private static Exception GetInnermostException(Exception ex)
		{
			while (ex.InnerException != null)
				ex = ex.InnerException;

			return ex;
		}

		private void BtnResetInstallDir_Click(object sender, EventArgs e)
		{
			this.TxtInstallDir.Text = Consts.DEFAULT_INSTALL_DIR;

			this.RefreshView();
		}

		private void TxtInstallDirMenu_コピー_Click(object sender, EventArgs e)
		{
			try
			{
				Clipboard.SetText(this.TxtInstallDir.Text);
			}
			catch
			{ }
		}

		#region インストール

		private void BtnInstall_Click(object sender, EventArgs e)
		{
			this.Visible = false;

			try
			{
				DoInstall();

				MessageDlg.Run(
					MessageDlg.Kind_e.Complete,
					"インストール完了",
					"インストールが完了しました。",
					null,
					new string[] { "OK" }
					);

				this.Close();
				return;
			}
			catch (Cancelled)
			{
				MessageDlg.Run(
					MessageDlg.Kind_e.Warning,
					"インストール中止",
					"インストールを中止しました。",
					null,
					new string[] { "OK" }
					);
			}
			catch (Exception ex)
			{
				MessageDlg.Run(
					MessageDlg.Kind_e.Error,
					"インストール失敗",
					"インストールに失敗しました。\r\n"
					+ "原因：" + GetInnermostException(ex).Message,
					ex,
					new string[] { "OK" }
					);
			}

			this.Visible = true;

			this.RefreshView();
			this.SetInitialFocus();
		}

		/// <summary>
		/// インストールを実行する。Ph-01
		/// </summary>
		private void DoInstall()
		{
			string installDir = this.TxtInstallDir.Text;
			string sigFile = Path.Combine(installDir, Consts.INSTALLED_SIGNATURE);
			bool createShortcutFlag = this.CBCreateShortcut.Checked;

			if (
				Directory.Exists(installDir) &&
				!File.Exists(sigFile)
				)
			{
				int ret = MessageDlg.Run(
					MessageDlg.Kind_e.Warning,
					"上書き確認",
					"インストール先フォルダは既に存在します。\r\n"
					+ "上書きしてよろしいですか？",
					null,
					new string[] { "はい", "いいえ", "キャンセル" }
					);

				if (ret != 1)
					throw new Cancelled();
			}

			if (createShortcutFlag)
			{
				string shortcutPath = ShortcutCreator.GetShortcutPath();

				if (File.Exists(shortcutPath))
				{
					int ret = MessageDlg.Run(
						MessageDlg.Kind_e.Warning,
						"ショートカットの上書き確認",
						$"ショートカット「{Path.GetFileNameWithoutExtension(shortcutPath)}」はデスクトップに既に作成されています。\r\n"
						+ "上書きしてよろしいですか？",
						null,
						new string[] { "はい", "いいえ", "キャンセル" }
						);

					if (ret != 1)
						throw new Cancelled();
				}
			}

			// 確認
			{
				int ret = MessageDlg.Run(
					MessageDlg.Kind_e.Question,
					"インストール実行の確認",
					$"「{Consts.APPLICATION_LONG_NAME}」のインストールを実行します。",
					null,
					new string[] { "OK", "キャンセル" }
					);

				if (ret != 1)
					throw new Cancelled();
			}

			ProcessingDlg.Run("インストール", () =>
			{
				DoInstallWkTh(
					installDir,
					sigFile,
					createShortcutFlag
					);
			});
		}

		/// <summary>
		/// インストールを実行する。Ph-02
		/// ワーカースレッド内で実行されることに注意！
		/// 例外を投げても良い。
		/// </summary>
		/// <param name="installDir">インストール先</param>
		/// <param name="sigFile">シグネチャファイル</param>
		/// <param name="createShortcutFlag">ショートカットを作成するか</param>
		private void DoInstallWkTh(
			string installDir,
			string sigFile,
			bool createShortcutFlag
			)
		{
			foreach (string fileName in Consts.CLUSTER_FILES)
			{
				string clusterFile = Path.Combine(ProcMain.SelfDir, fileName);
				string clusterHashFile = clusterFile + Consts.HASH_EXTENSION;

				if (!File.Exists(clusterFile))
					throw new Exception($"クラスタファイル「{fileName}」が見つかりません。");

				if (!File.Exists(clusterHashFile))
					throw new Exception($"クラスタファイル「{fileName}」のハッシュ値が見つかりません。");

				byte[] hash1 = SCommon.GetSHA512File(clusterFile);
				byte[] hash2 = SCommon.Hex.I.GetBytes(File.ReadAllText(clusterHashFile, Encoding.ASCII).Trim());

				if (SCommon.Comp(hash1, hash2, SCommon.Comp) != 0)
					throw new Exception($"クラスタファイル「{fileName}」が破損しています。");
			}

			if (Directory.Exists(installDir)) // インストール先フォルダのハンドル残存チェック
			{
				string escapeDir = SCommon.ToCreatablePath(installDir);

				try
				{
					SCommon.EnsureMoveDir(installDir, escapeDir);
				}
				catch
				{
					throw new Exception("インストール先のフォルダは現在使用されています。");
				}

				// memo:
				// インストール先にはデータベースなどが作成されているかもしれないので、
				// インストール先そのものの削除・再作成は行わないようにする！

				SCommon.EnsureMoveDir(escapeDir, installDir); // 元に戻す。
			}
			SCommon.CreateDir(installDir);

			// memo:
			// 既インストール先にデータベースなどが存在することを想定し、
			// 必要なフォルダ・ファイルのみ削除・再作成を行う！

			SCommon.DeletePath(sigFile);
			File.WriteAllBytes(sigFile, SCommon.EMPTY_BYTES);

			foreach (string fileName in Consts.CLUSTER_FILES)
			{
				string clusterFile = Path.Combine(ProcMain.SelfDir, fileName);
				string extractedDir = Path.Combine(installDir, P_EraseExtension(fileName));

				DirToClusterFileTools.ClusterFileToDir(clusterFile, extractedDir);
			}

			if (createShortcutFlag)
				ShortcutCreator.Run(Path.Combine(installDir, Consts.MAIN_PROGRAM));
		}

		private static string P_EraseExtension(string fileName)
		{
			int p = fileName.LastIndexOf('.');

			if (p == -1) // ? 見つからない。
				throw null; // never

			if (p == 0) // ? 拡張子で始まっている。
				throw null; // never

			return fileName.Substring(0, p);
		}

		#endregion

		#region アンインストール

		private void BtnUninstall_Click(object sender, EventArgs e)
		{
			this.Visible = false;

			try
			{
				DoUninstall();

				MessageDlg.Run(
					MessageDlg.Kind_e.Information,
					"アンインストール完了",
					"アンインストールが完了しました。",
					null,
					new string[] { "OK" }
					);

				this.Close();
				return;
			}
			catch (Cancelled)
			{
				MessageDlg.Run(
					MessageDlg.Kind_e.Warning,
					"アンインストール中止",
					"アンインストールを中止しました。",
					null,
					new string[] { "OK" }
					);
			}
			catch (Exception ex)
			{
				MessageDlg.Run(
					MessageDlg.Kind_e.Error,
					"アンインストール失敗",
					"アンインストールに失敗しました。\r\n"
					+ "原因：" + GetInnermostException(ex).Message,
					ex,
					new string[] { "OK" }
					);
			}

			this.Visible = true;

			this.RefreshView();
			this.SetInitialFocus();
		}

		/// <summary>
		/// アンインストールを実行する。Ph-01
		/// </summary>
		private void DoUninstall()
		{
			string installDir = this.TxtInstallDir.Text;
			bool removeShortcutFlag = false;

			// ショートカット削除の確認
			{
				string shortcutPath = ShortcutCreator.GetShortcutPath();

				if (File.Exists(shortcutPath))
				{
					int ret = MessageDlg.Run(
						MessageDlg.Kind_e.Question,
						"ショートカット削除の確認",
						$"デスクトップ上のショートカット「{Path.GetFileNameWithoutExtension(shortcutPath)}」を削除しますか？",
						null,
						new string[] { "はい", "いいえ", "キャンセル" }
						);

					if (ret == 1)
						removeShortcutFlag = true;
					else if (ret == 2)
						removeShortcutFlag = false;
					else
						throw new Cancelled();
				}
			}

			// 確認
			{
				int ret = MessageDlg.Run(
					MessageDlg.Kind_e.Warning,
					"アンインストール実行の確認",
					$"「{Consts.APPLICATION_LONG_NAME}」をアンインストールします。\r\n"
					+ "インストール先フォルダにある全てのファイルが削除されます。\r\n"
					+ "実行してよろしいですか？\r\n"
					+ $"( ショートカットの削除：{(removeShortcutFlag ? "する" : "しない")} )",
					null,
					new string[] { "OK", "キャンセル" }
					);

				if (ret != 1)
					throw new Cancelled();
			}

			ProcessingDlg.Run("アンインストール", () =>
			{
				DoUninstallWkTh(
					installDir,
					removeShortcutFlag
					);
			});
		}

		/// <summary>
		/// アンインストールを実行する。Ph-02
		/// ワーカースレッド内で実行されることに注意！
		/// 例外を投げても良い。
		/// </summary>
		/// <param name="installDir">インストール先</param>
		/// <param name="removeShortcutFlag">ショートカットを削除するか</param>
		private void DoUninstallWkTh(
			string installDir,
			bool removeShortcutFlag
			)
		{
			if (!Directory.Exists(installDir))
				throw new Exception("インストール先のフォルダが存在しません！");

			// インストール先フォルダのハンドル残存チェック -> 削除
			{
				string escapeDir = SCommon.ToCreatablePath(installDir);

				try
				{
					SCommon.EnsureMoveDir(installDir, escapeDir);
				}
				catch
				{
					throw new Exception("インストール先のフォルダは現在使用されています。");
				}

				SCommon.DeletePath(escapeDir);
			}

			if (removeShortcutFlag)
				SCommon.DeletePath(ShortcutCreator.GetShortcutPath());
		}

		#endregion

		private class Cancelled : Exception
		{ }
	}
}
