using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HLTStudio.Dialogs
{
	public partial class MessageDlg : Form
	{
		public enum Kind_e
		{
			Error = 1,
			Warning,
			Information,
			Question,
			Complete,
		}

		/// <summary>
		/// メッセージダイアログを表示する。
		/// </summary>
		/// <param name="kind">メッセージの種類(アイコンの種類)</param>
		/// <param name="title">タイトル</param>
		/// <param name="message">主メッセージ(改行可)</param>
		/// <param name="detailMessage">詳細メッセージ(改行可・null可)</param>
		/// <param name="options">選択肢(1～3個)</param>
		/// <returns>選択された番号(1～3・右上の[X]ボタンを押したときは最後の選択肢が選ばれる)</returns>
		public static int Run(Kind_e kind, string title, string message, object detailMessage, string[] options)
		{
			if (
				kind != Kind_e.Error &&
				kind != Kind_e.Warning &&
				kind != Kind_e.Information &&
				kind != Kind_e.Question &&
				kind != Kind_e.Complete
				)
				kind = Kind_e.Error;

			title = ("" + title).Trim();

			if (title == "")
				title = "(タイトルを取得できませんでした)";

			message = ("" + message).Trim();

			if (message == "")
				message = "(メッセージを取得できませんでした)";

			string strDetailMessage;
			try
			{
				strDetailMessage = "" + detailMessage;
			}
			catch
			{
				strDetailMessage = "";
			}

			if (
				options == null ||
				options.Length < 1 ||
				options.Length > 3
				)
				options = new string[] { "(選択1)" };

			options = options.Select((option, i) =>
			{
				option = ("" + option).Trim();

				if (option == "")
					option = $"(選択{i + 1})";

				return option;
			})
			.ToArray();

			using (MessageDlg f = new MessageDlg())
			{
				f.Kind = kind;
				f.TitleText = title;
				f.MainMessage = message;
				f.DetailMessage = strDetailMessage;
				f.Options = options;

				f.ShowDialog();

				return f.RetValue;
			}
		}

		private Kind_e Kind;
		private string TitleText;
		private string MainMessage;
		private string DetailMessage;
		private string[] Options;
		private int RetValue;

		private MessageDlg()
		{
			InitializeComponent();
		}

		private void MessageDlg_Load(object sender, EventArgs e)
		{
			switch (this.Kind)
			{
				case Kind_e.Error:
					// noop
					break;

				case Kind_e.Warning:
					this.ErrorIcon.Visible = false;
					this.WarningIcon.Location = this.ErrorIcon.Location;
					this.WarningIcon.Size = this.ErrorIcon.Size;
					this.WarningIcon.Visible = true;
					break;

				case Kind_e.Information:
					this.ErrorIcon.Visible = false;
					this.InformationIcon.Location = this.ErrorIcon.Location;
					this.InformationIcon.Size = this.ErrorIcon.Size;
					this.InformationIcon.Visible = true;
					break;

				case Kind_e.Question:
					this.ErrorIcon.Visible = false;
					this.QuestionIcon.Location = this.ErrorIcon.Location;
					this.QuestionIcon.Size = this.ErrorIcon.Size;
					this.QuestionIcon.Visible = true;
					break;

				case Kind_e.Complete:
					this.ErrorIcon.Visible = false;
					this.CompleteIcon.Location = this.ErrorIcon.Location;
					this.CompleteIcon.Size = this.ErrorIcon.Size;
					this.CompleteIcon.Visible = true;
					break;

				default:
					throw null; // never
			}

			this.Text = this.TitleText;

			this.LShowDetails.Visible = this.DetailMessage != "";

			this.LMainMessage.Text = this.MainMessage;
			this.TxtDetailMessage.Text = this.DetailMessage;

			this.LMainMessage.Top = this.ErrorIcon.Top + (this.ErrorIcon.Height - this.LMainMessage.Height) / 2;

			// ウィンドウの横幅調整
			{
				const int MARGIN_R = 30;

				int r = this.LMainMessage.Right + MARGIN_R;
				int w = this.Width;

				if (w < r)
				{
					int aw = r - w;

					this.Width += aw;
					this.Left -= aw / 2;
				}
			}

			switch (this.Options.Length)
			{
				case 1:
					this.BtnOption1.Text = this.Options[0];
					this.BtnOption2.Visible = false;
					this.BtnOption3.Visible = false;
					this.SetOptionButtonPositions(this.BtnOption1);
					this.ControlShownFocusSet = this.BtnOption1;
					this.RetValue = 1;
					break;

				case 2:
					this.BtnOption1.Text = this.Options[0];
					this.BtnOption2.Text = this.Options[1];
					this.BtnOption3.Visible = false;
					this.SetOptionButtonPositions(this.BtnOption1, this.BtnOption2);
					this.ControlShownFocusSet = this.BtnOption2;
					this.RetValue = 2;
					break;

				case 3:
					this.BtnOption1.Text = this.Options[0];
					this.BtnOption2.Text = this.Options[1];
					this.BtnOption3.Text = this.Options[2];
					this.SetOptionButtonPositions(this.BtnOption1, this.BtnOption2, this.BtnOption3);
					this.ControlShownFocusSet = this.BtnOption3;
					this.RetValue = 3;
					break;

				default:
					throw null; // never
			}

			this.MinimumSize = this.Size;

			this.SetShowDetailMessage(false);
			this.Top += DETAIL_MESSAGE_H_ADD / 2;
		}

		private void SetOptionButtonPositions(params Button[] buttons)
		{
			const int DEFAULT_W = 120;
			const int MARGIN_X = 30;
			const int GAP_X = 10;

			int w = Math.Max(buttons.Max(button => GetButtonTextPxWidth(button)) + MARGIN_X, DEFAULT_W);
			int r = this.BtnOption3.Right;

			for (int i = buttons.Length - 1; 0 <= i; i--)
			{
				r -= w;

				buttons[i].Width = w;
				buttons[i].Left = r;

				r -= GAP_X;
			}
		}

		private static int GetButtonTextPxWidth(Button button)
		{
			// テキスト描画に必要なサイズを取得
			Size textSize = TextRenderer.MeasureText(
				button.Text,
				button.Font,
				new Size(int.MaxValue, int.MaxValue),
				TextFormatFlags.SingleLine | TextFormatFlags.NoPadding
				);

			// ボタンの余白（パディングや枠線）を加味
			//int width = textSize.Width + 12; // 左右のマージン分を適当に加える
			//int height = textSize.Height + 8; // 上下の余白分を加える

			//return new Size(width, height);
			return textSize.Width;
		}

		private Control ControlShownFocusSet;

		private void MessageDlg_Shown(object sender, EventArgs e)
		{
			this.ControlShownFocusSet.Focus();
		}

		private void MainPanel_Paint(object sender, PaintEventArgs e)
		{
			// none
		}

		private void LShowDetails_Click(object sender, EventArgs e)
		{
			if (this.TxtDetailMessage.Visible) // ? 表示中 -> 隠す。
			{
				this.SetShowDetailMessage(false);
			}
			else // ? 非表示 -> 表示する。
			{
				this.SetShowDetailMessage(true);
			}
		}

		private const int DETAIL_MESSAGE_H_ADD = 200;

		private void SetShowDetailMessage(bool showMode)
		{
			if (showMode) // ? 表示する。
			{
				this.LShowDetails.Text = "詳細を隠す";
				this.TxtDetailMessage.Visible = true;
				this.MinimumSize = new Size(
					this.MinimumSize.Width,
					this.MinimumSize.Height + DETAIL_MESSAGE_H_ADD
					);
			}
			else // ? 隠す。
			{
				this.LShowDetails.Text = "詳細を表示する";
				this.TxtDetailMessage.Visible = false;
				this.MinimumSize = new Size(
					this.MinimumSize.Width,
					this.MinimumSize.Height - DETAIL_MESSAGE_H_ADD
					);
			}
			this.Size = this.MinimumSize;
		}

		private void BtnOption1_Click(object sender, EventArgs e)
		{
			this.RetValue = 1;
			this.Close();
		}

		private void BtnOption2_Click(object sender, EventArgs e)
		{
			this.RetValue = 2;
			this.Close();
		}

		private void BtnOption3_Click(object sender, EventArgs e)
		{
			this.RetValue = 3;
			this.Close();
		}

		private void TxtDetailMessageMenu_コピー_Click(object sender, EventArgs e)
		{
			try
			{
				string text = this.LMainMessage.Text + "\r\n" + this.TxtDetailMessage.Text;

				if (text != "")
				{
					Clipboard.SetText(text);
				}
			}
			catch
			{ }
		}
	}
}
