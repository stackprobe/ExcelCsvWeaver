namespace HLTStudio.Dialogs
{
	partial class MessageDlg
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
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MessageDlg));
			this.MainPanel = new System.Windows.Forms.Panel();
			this.WarningIcon = new System.Windows.Forms.PictureBox();
			this.QuestionIcon = new System.Windows.Forms.PictureBox();
			this.InformationIcon = new System.Windows.Forms.PictureBox();
			this.TxtDetailMessage = new System.Windows.Forms.TextBox();
			this.TxtDetailMessageMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.TxtDetailMessageMenu_コピー = new System.Windows.Forms.ToolStripMenuItem();
			this.LMainMessage = new System.Windows.Forms.Label();
			this.ErrorIcon = new System.Windows.Forms.PictureBox();
			this.BtnOption1 = new System.Windows.Forms.Button();
			this.LShowDetails = new System.Windows.Forms.Label();
			this.BtnOption2 = new System.Windows.Forms.Button();
			this.BtnOption3 = new System.Windows.Forms.Button();
			this.CompleteIcon = new System.Windows.Forms.PictureBox();
			this.MainPanel.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.WarningIcon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.QuestionIcon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.InformationIcon)).BeginInit();
			this.TxtDetailMessageMenu.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ErrorIcon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CompleteIcon)).BeginInit();
			this.SuspendLayout();
			// 
			// MainPanel
			// 
			this.MainPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.MainPanel.BackColor = System.Drawing.Color.White;
			this.MainPanel.Controls.Add(this.CompleteIcon);
			this.MainPanel.Controls.Add(this.WarningIcon);
			this.MainPanel.Controls.Add(this.QuestionIcon);
			this.MainPanel.Controls.Add(this.InformationIcon);
			this.MainPanel.Controls.Add(this.TxtDetailMessage);
			this.MainPanel.Controls.Add(this.LMainMessage);
			this.MainPanel.Controls.Add(this.ErrorIcon);
			this.MainPanel.Location = new System.Drawing.Point(0, 0);
			this.MainPanel.Name = "MainPanel";
			this.MainPanel.Size = new System.Drawing.Size(600, 386);
			this.MainPanel.TabIndex = 0;
			this.MainPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.MainPanel_Paint);
			// 
			// WarningIcon
			// 
			this.WarningIcon.Image = ((System.Drawing.Image)(resources.GetObject("WarningIcon.Image")));
			this.WarningIcon.Location = new System.Drawing.Point(162, 20);
			this.WarningIcon.Name = "WarningIcon";
			this.WarningIcon.Size = new System.Drawing.Size(40, 40);
			this.WarningIcon.TabIndex = 5;
			this.WarningIcon.TabStop = false;
			this.WarningIcon.Visible = false;
			// 
			// QuestionIcon
			// 
			this.QuestionIcon.Image = ((System.Drawing.Image)(resources.GetObject("QuestionIcon.Image")));
			this.QuestionIcon.Location = new System.Drawing.Point(254, 20);
			this.QuestionIcon.Name = "QuestionIcon";
			this.QuestionIcon.Size = new System.Drawing.Size(40, 40);
			this.QuestionIcon.TabIndex = 4;
			this.QuestionIcon.TabStop = false;
			this.QuestionIcon.Visible = false;
			// 
			// InformationIcon
			// 
			this.InformationIcon.Image = ((System.Drawing.Image)(resources.GetObject("InformationIcon.Image")));
			this.InformationIcon.Location = new System.Drawing.Point(208, 20);
			this.InformationIcon.Name = "InformationIcon";
			this.InformationIcon.Size = new System.Drawing.Size(40, 40);
			this.InformationIcon.TabIndex = 3;
			this.InformationIcon.TabStop = false;
			this.InformationIcon.Visible = false;
			// 
			// TxtDetailMessage
			// 
			this.TxtDetailMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TxtDetailMessage.ContextMenuStrip = this.TxtDetailMessageMenu;
			this.TxtDetailMessage.Location = new System.Drawing.Point(12, 170);
			this.TxtDetailMessage.Multiline = true;
			this.TxtDetailMessage.Name = "TxtDetailMessage";
			this.TxtDetailMessage.ReadOnly = true;
			this.TxtDetailMessage.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.TxtDetailMessage.Size = new System.Drawing.Size(560, 203);
			this.TxtDetailMessage.TabIndex = 1;
			this.TxtDetailMessage.Text = resources.GetString("TxtDetailMessage.Text");
			this.TxtDetailMessage.Visible = false;
			this.TxtDetailMessage.WordWrap = false;
			// 
			// TxtDetailMessageMenu
			// 
			this.TxtDetailMessageMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TxtDetailMessageMenu_コピー});
			this.TxtDetailMessageMenu.Name = "TxtDetailMessageMenu";
			this.TxtDetailMessageMenu.Size = new System.Drawing.Size(115, 26);
			// 
			// TxtDetailMessageMenu_コピー
			// 
			this.TxtDetailMessageMenu_コピー.Name = "TxtDetailMessageMenu_コピー";
			this.TxtDetailMessageMenu_コピー.Size = new System.Drawing.Size(114, 22);
			this.TxtDetailMessageMenu_コピー.Text = "コピー(&C)";
			this.TxtDetailMessageMenu_コピー.Click += new System.EventHandler(this.TxtDetailMessageMenu_コピー_Click);
			// 
			// LMainMessage
			// 
			this.LMainMessage.AutoSize = true;
			this.LMainMessage.Location = new System.Drawing.Point(170, 80);
			this.LMainMessage.Name = "LMainMessage";
			this.LMainMessage.Size = new System.Drawing.Size(321, 20);
			this.LMainMessage.TabIndex = 0;
			this.LMainMessage.Text = "アプリケーションの処理中にエラーが発生しました。";
			// 
			// ErrorIcon
			// 
			this.ErrorIcon.Image = ((System.Drawing.Image)(resources.GetObject("ErrorIcon.Image")));
			this.ErrorIcon.Location = new System.Drawing.Point(20, 20);
			this.ErrorIcon.Name = "ErrorIcon";
			this.ErrorIcon.Size = new System.Drawing.Size(136, 136);
			this.ErrorIcon.TabIndex = 0;
			this.ErrorIcon.TabStop = false;
			// 
			// BtnOption1
			// 
			this.BtnOption1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnOption1.Location = new System.Drawing.Point(200, 392);
			this.BtnOption1.Name = "BtnOption1";
			this.BtnOption1.Size = new System.Drawing.Size(120, 57);
			this.BtnOption1.TabIndex = 2;
			this.BtnOption1.Text = "Option1";
			this.BtnOption1.UseVisualStyleBackColor = true;
			this.BtnOption1.Click += new System.EventHandler(this.BtnOption1_Click);
			// 
			// LShowDetails
			// 
			this.LShowDetails.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.LShowDetails.AutoSize = true;
			this.LShowDetails.Cursor = System.Windows.Forms.Cursors.Hand;
			this.LShowDetails.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.LShowDetails.ForeColor = System.Drawing.Color.Navy;
			this.LShowDetails.Location = new System.Drawing.Point(12, 432);
			this.LShowDetails.Name = "LShowDetails";
			this.LShowDetails.Size = new System.Drawing.Size(100, 20);
			this.LShowDetails.TabIndex = 1;
			this.LShowDetails.Text = "詳細を表示する";
			this.LShowDetails.Click += new System.EventHandler(this.LShowDetails_Click);
			// 
			// BtnOption2
			// 
			this.BtnOption2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnOption2.Location = new System.Drawing.Point(326, 392);
			this.BtnOption2.Name = "BtnOption2";
			this.BtnOption2.Size = new System.Drawing.Size(120, 57);
			this.BtnOption2.TabIndex = 3;
			this.BtnOption2.Text = "Option2";
			this.BtnOption2.UseVisualStyleBackColor = true;
			this.BtnOption2.Click += new System.EventHandler(this.BtnOption2_Click);
			// 
			// BtnOption3
			// 
			this.BtnOption3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnOption3.Location = new System.Drawing.Point(452, 392);
			this.BtnOption3.Name = "BtnOption3";
			this.BtnOption3.Size = new System.Drawing.Size(120, 57);
			this.BtnOption3.TabIndex = 4;
			this.BtnOption3.Text = "Option3";
			this.BtnOption3.UseVisualStyleBackColor = true;
			this.BtnOption3.Click += new System.EventHandler(this.BtnOption3_Click);
			// 
			// CompleteIcon
			// 
			this.CompleteIcon.Image = ((System.Drawing.Image)(resources.GetObject("CompleteIcon.Image")));
			this.CompleteIcon.Location = new System.Drawing.Point(300, 20);
			this.CompleteIcon.Name = "CompleteIcon";
			this.CompleteIcon.Size = new System.Drawing.Size(40, 40);
			this.CompleteIcon.TabIndex = 6;
			this.CompleteIcon.TabStop = false;
			this.CompleteIcon.Visible = false;
			// 
			// MessageDlg
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(584, 461);
			this.Controls.Add(this.BtnOption3);
			this.Controls.Add(this.BtnOption2);
			this.Controls.Add(this.LShowDetails);
			this.Controls.Add(this.BtnOption1);
			this.Controls.Add(this.MainPanel);
			this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "MessageDlg";
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "メッセージ";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.MessageDlg_Load);
			this.Shown += new System.EventHandler(this.MessageDlg_Shown);
			this.MainPanel.ResumeLayout(false);
			this.MainPanel.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.WarningIcon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.QuestionIcon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.InformationIcon)).EndInit();
			this.TxtDetailMessageMenu.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ErrorIcon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CompleteIcon)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Panel MainPanel;
		private System.Windows.Forms.PictureBox ErrorIcon;
		private System.Windows.Forms.Button BtnOption1;
		private System.Windows.Forms.Label LMainMessage;
		private System.Windows.Forms.Label LShowDetails;
		private System.Windows.Forms.Button BtnOption2;
		private System.Windows.Forms.Button BtnOption3;
		private System.Windows.Forms.TextBox TxtDetailMessage;
		private System.Windows.Forms.PictureBox InformationIcon;
		private System.Windows.Forms.PictureBox WarningIcon;
		private System.Windows.Forms.PictureBox QuestionIcon;
		private System.Windows.Forms.ContextMenuStrip TxtDetailMessageMenu;
		private System.Windows.Forms.ToolStripMenuItem TxtDetailMessageMenu_コピー;
		private System.Windows.Forms.PictureBox CompleteIcon;
	}
}