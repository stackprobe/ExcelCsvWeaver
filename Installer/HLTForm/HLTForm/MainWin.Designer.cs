namespace HLTStudio
{
	partial class MainWin
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows フォーム デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWin));
			this.LMainMessage = new System.Windows.Forms.Label();
			this.TxtInstallDir = new System.Windows.Forms.TextBox();
			this.TxtInstallDirMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.TxtInstallDirMenu_コピー = new System.Windows.Forms.ToolStripMenuItem();
			this.BtnBrowse = new System.Windows.Forms.Button();
			this.BtnInstall = new System.Windows.Forms.Button();
			this.CBCreateShortcut = new System.Windows.Forms.CheckBox();
			this.MainPanel = new System.Windows.Forms.Panel();
			this.G001 = new System.Windows.Forms.GroupBox();
			this.BtnResetInstallDir = new System.Windows.Forms.Button();
			this.BtnUninstall = new System.Windows.Forms.Button();
			this.TxtInstallDirMenu.SuspendLayout();
			this.MainPanel.SuspendLayout();
			this.G001.SuspendLayout();
			this.SuspendLayout();
			// 
			// LMainMessage
			// 
			this.LMainMessage.AutoSize = true;
			this.LMainMessage.Location = new System.Drawing.Point(40, 50);
			this.LMainMessage.Name = "LMainMessage";
			this.LMainMessage.Size = new System.Drawing.Size(282, 20);
			this.LMainMessage.TabIndex = 0;
			this.LMainMessage.Text = "「アプリケーション」をインストールします。";
			// 
			// TxtInstallDir
			// 
			this.TxtInstallDir.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TxtInstallDir.ContextMenuStrip = this.TxtInstallDirMenu;
			this.TxtInstallDir.Location = new System.Drawing.Point(50, 50);
			this.TxtInstallDir.Name = "TxtInstallDir";
			this.TxtInstallDir.ReadOnly = true;
			this.TxtInstallDir.Size = new System.Drawing.Size(580, 27);
			this.TxtInstallDir.TabIndex = 0;
			this.TxtInstallDir.Text = "C:\\HLT\\Program";
			// 
			// TxtInstallDirMenu
			// 
			this.TxtInstallDirMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TxtInstallDirMenu_コピー});
			this.TxtInstallDirMenu.Name = "TxtInstallDirMenu";
			this.TxtInstallDirMenu.Size = new System.Drawing.Size(115, 26);
			// 
			// TxtInstallDirMenu_コピー
			// 
			this.TxtInstallDirMenu_コピー.Name = "TxtInstallDirMenu_コピー";
			this.TxtInstallDirMenu_コピー.Size = new System.Drawing.Size(114, 22);
			this.TxtInstallDirMenu_コピー.Text = "コピー(&C)";
			this.TxtInstallDirMenu_コピー.Click += new System.EventHandler(this.TxtInstallDirMenu_コピー_Click);
			// 
			// BtnBrowse
			// 
			this.BtnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnBrowse.Location = new System.Drawing.Point(550, 83);
			this.BtnBrowse.Name = "BtnBrowse";
			this.BtnBrowse.Size = new System.Drawing.Size(80, 30);
			this.BtnBrowse.TabIndex = 2;
			this.BtnBrowse.Text = "参照...";
			this.BtnBrowse.UseVisualStyleBackColor = true;
			this.BtnBrowse.Click += new System.EventHandler(this.BtnBrowse_Click);
			// 
			// BtnInstall
			// 
			this.BtnInstall.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnInstall.ForeColor = System.Drawing.Color.DarkBlue;
			this.BtnInstall.Location = new System.Drawing.Point(552, 386);
			this.BtnInstall.Name = "BtnInstall";
			this.BtnInstall.Size = new System.Drawing.Size(200, 93);
			this.BtnInstall.TabIndex = 2;
			this.BtnInstall.Text = "インストール";
			this.BtnInstall.Click += new System.EventHandler(this.BtnInstall_Click);
			// 
			// CBCreateShortcut
			// 
			this.CBCreateShortcut.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.CBCreateShortcut.AutoSize = true;
			this.CBCreateShortcut.Checked = true;
			this.CBCreateShortcut.CheckState = System.Windows.Forms.CheckState.Checked;
			this.CBCreateShortcut.Location = new System.Drawing.Point(44, 310);
			this.CBCreateShortcut.Name = "CBCreateShortcut";
			this.CBCreateShortcut.Size = new System.Drawing.Size(275, 24);
			this.CBCreateShortcut.TabIndex = 2;
			this.CBCreateShortcut.Text = "デスクトップにショートカットを作成する";
			this.CBCreateShortcut.UseVisualStyleBackColor = true;
			// 
			// MainPanel
			// 
			this.MainPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.MainPanel.BackColor = System.Drawing.Color.White;
			this.MainPanel.Controls.Add(this.LMainMessage);
			this.MainPanel.Controls.Add(this.G001);
			this.MainPanel.Controls.Add(this.CBCreateShortcut);
			this.MainPanel.Location = new System.Drawing.Point(0, 0);
			this.MainPanel.Name = "MainPanel";
			this.MainPanel.Size = new System.Drawing.Size(780, 380);
			this.MainPanel.TabIndex = 0;
			// 
			// G001
			// 
			this.G001.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.G001.Controls.Add(this.TxtInstallDir);
			this.G001.Controls.Add(this.BtnResetInstallDir);
			this.G001.Controls.Add(this.BtnBrowse);
			this.G001.Location = new System.Drawing.Point(44, 110);
			this.G001.Name = "G001";
			this.G001.Size = new System.Drawing.Size(680, 150);
			this.G001.TabIndex = 1;
			this.G001.TabStop = false;
			this.G001.Text = "インストール先";
			// 
			// BtnResetInstallDir
			// 
			this.BtnResetInstallDir.Location = new System.Drawing.Point(50, 83);
			this.BtnResetInstallDir.Name = "BtnResetInstallDir";
			this.BtnResetInstallDir.Size = new System.Drawing.Size(160, 30);
			this.BtnResetInstallDir.TabIndex = 1;
			this.BtnResetInstallDir.Text = "デフォルトに戻す";
			this.BtnResetInstallDir.UseVisualStyleBackColor = true;
			this.BtnResetInstallDir.Click += new System.EventHandler(this.BtnResetInstallDir_Click);
			// 
			// BtnUninstall
			// 
			this.BtnUninstall.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnUninstall.ForeColor = System.Drawing.Color.DarkRed;
			this.BtnUninstall.Location = new System.Drawing.Point(346, 386);
			this.BtnUninstall.Name = "BtnUninstall";
			this.BtnUninstall.Size = new System.Drawing.Size(200, 93);
			this.BtnUninstall.TabIndex = 1;
			this.BtnUninstall.Text = "アンインストール";
			this.BtnUninstall.Click += new System.EventHandler(this.BtnUninstall_Click);
			// 
			// MainWin
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(764, 491);
			this.Controls.Add(this.BtnUninstall);
			this.Controls.Add(this.BtnInstall);
			this.Controls.Add(this.MainPanel);
			this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.MinimizeBox = false;
			this.Name = "MainWin";
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "アプリケーション インストーラ　( ビルド番号：v999-999-99999 )";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.MainWin_Load);
			this.Shown += new System.EventHandler(this.MainWin_Shown);
			this.TxtInstallDirMenu.ResumeLayout(false);
			this.MainPanel.ResumeLayout(false);
			this.MainPanel.PerformLayout();
			this.G001.ResumeLayout(false);
			this.G001.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Label LMainMessage;
		private System.Windows.Forms.TextBox TxtInstallDir;
		private System.Windows.Forms.Button BtnBrowse;
		private System.Windows.Forms.Button BtnInstall;
		private System.Windows.Forms.CheckBox CBCreateShortcut;
		private System.Windows.Forms.Panel MainPanel;
		private System.Windows.Forms.GroupBox G001;
		private System.Windows.Forms.Button BtnResetInstallDir;
		private System.Windows.Forms.ContextMenuStrip TxtInstallDirMenu;
		private System.Windows.Forms.ToolStripMenuItem TxtInstallDirMenu_コピー;
		private System.Windows.Forms.Button BtnUninstall;
	}
}

