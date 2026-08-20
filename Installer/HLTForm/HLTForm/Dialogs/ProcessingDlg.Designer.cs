namespace HLTStudio.Dialogs
{
	partial class ProcessingDlg
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProcessingDlg));
			this.PB001 = new System.Windows.Forms.ProgressBar();
			this.MainMessage = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// PB001
			// 
			this.PB001.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.PB001.Location = new System.Drawing.Point(12, 130);
			this.PB001.Name = "PB001";
			this.PB001.Size = new System.Drawing.Size(560, 30);
			this.PB001.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
			this.PB001.TabIndex = 0;
			// 
			// LMainMessage
			// 
			this.MainMessage.AutoSize = true;
			this.MainMessage.Location = new System.Drawing.Point(40, 60);
			this.MainMessage.Name = "LMainMessage";
			this.MainMessage.Size = new System.Drawing.Size(310, 20);
			this.MainMessage.TabIndex = 1;
			this.MainMessage.Text = "「アプリケーション」をインストールしています...";
			// 
			// ProcessingDlg
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(584, 201);
			this.Controls.Add(this.MainMessage);
			this.Controls.Add(this.PB001);
			this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ProcessingDlg";
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "インストール中";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.ProcessingDlg_Load);
			this.Shown += new System.EventHandler(this.ProcessingDlg_Shown);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ProgressBar PB001;
		private System.Windows.Forms.Label MainMessage;
	}
}