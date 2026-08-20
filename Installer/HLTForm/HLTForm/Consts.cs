using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using HLTStudio.Commons;

namespace HLTStudio
{
	public static class Consts
	{
		/// <summary>
		/// アプリケーション名
		/// 用途：表示・デフォルトのフォルダ名・ショートカット名
		/// </summary>
		public static readonly string APPLICATION_NAME = "ExcelCsvWeaver";

		/// <summary>
		/// アプリケーション名
		/// 用途：表示
		/// </summary>
		public static readonly string APPLICATION_LONG_NAME = "エクセル CSV 変換・加工プログラム";

		/// <summary>
		/// デフォルトのインストール先
		/// </summary>
		public static string DEFAULT_INSTALL_DIR
		{
			get
			{
				string appDataLocalDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

				if (
					string.IsNullOrEmpty(appDataLocalDir) ||
					!Directory.Exists(appDataLocalDir)
					)
					throw new Exception("Bad appDataLocalDir");

				return Path.Combine(appDataLocalDir, APPLICATION_NAME);
			}
		}

		/// <summary>
		/// クラスタファイルの一覧
		/// </summary>
		public static readonly string[] CLUSTER_FILES = new string[]
		{
			"ECWeaverGUI.cmp-gz",
			"ECWeaver.cmp-gz",
			"ECWeaver2.cmp-gz",
		};

		/// <summary>
		/// クラスタファイルのハッシュ値ファイル用の拡張子
		/// </summary>
		public static readonly string HASH_EXTENSION = ".hash";

		/// <summary>
		/// 起動プログラム
		/// インストール先からの相対パス
		/// 例："MainProgram\\MainProgram.exe"
		/// </summary>
		public static string MAIN_PROGRAM
		{
			get
			{
				string mainCluster = CLUSTER_FILES[0];
				string mainClusterName;

				{
					int p = mainCluster.IndexOf('.');

					if (p == -1)
						throw null; // never

					mainClusterName = mainCluster.Substring(0, p);
				}

				if (mainClusterName == "")
					throw null; // never

				return $"{mainClusterName}\\{mainClusterName}.exe";
			}
		}

		/// <summary>
		/// インストール先に配置するシグネチャ・ファイル名
		/// 既インストール先の判定に使用する。
		/// </summary>
		public static string INSTALLED_SIGNATURE
		{
			get
			{
				const string TRAILER_PATTERN = "_{857ac54a-2973-45d2-bcbd-bf5086b313d9}";

				string hash = SCommon.Base32.I.Encode(
					SCommon.GetPart(
						SCommon.GetSHA512(
							Encoding.UTF8.GetBytes(
								APPLICATION_NAME + TRAILER_PATTERN
								)
							)
							, 0
							, 20
						)
					);

				return $"HLT_{hash}";
			}
		}
	}
}
