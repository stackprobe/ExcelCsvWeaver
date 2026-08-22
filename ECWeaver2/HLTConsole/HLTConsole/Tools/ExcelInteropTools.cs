using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using HLTStudio.Commons;

namespace HLTStudio.Tools
{
	// memo: プロジェクト右クリック -> NuGet パッケージの監理 -> 参照 -> Microsoft.Office.Interop.Excel 16.0.18925.20022

	public static class ExcelInteropTools
	{
		public class SheetData
		{
			public string Name;
			public string[][] Rows;

			public SheetData(string name, string[][] rows)
			{
				this.Name = name;
				this.Rows = rows ?? new string[0][];
			}
		}

		private static void ExecuteWorkbook(string excelFile, Action<Microsoft.Office.Interop.Excel.Application, Microsoft.Office.Interop.Excel.Workbook> routine)
		{
			Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
			try
			{
				app.Visible = false;

				Microsoft.Office.Interop.Excel.Workbooks workbooks = app.Workbooks;
				try
				{
					Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Open(excelFile);
					try
					{
						routine(app, workbook);
					}
					finally
					{
						workbook.Close(false);
						Marshal.ReleaseComObject(workbook);
						workbook = null;
					}
				}
				finally
				{
					//workbooks.Close(); // 不要
					Marshal.ReleaseComObject(workbooks);
					workbooks = null;
				}
			}
			finally
			{
				app.Quit();
				Marshal.ReleaseComObject(app);
				app = null;

				// 1
				GC.Collect();
				GC.WaitForPendingFinalizers();

				// 2
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
		}

		public static void ToPDF(string excelFile, string pdfFile)
		{
			ExecuteWorkbook(excelFile, (app, workbook) =>
			{
				workbook.ExportAsFixedFormat(
					Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
					pdfFile,
					Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
					IncludeDocProperties: true,
					IgnorePrintAreas: false,
					OpenAfterPublish: false
					);
			});
		}

		public static void WriteWorkbook(string excelFile, SheetData[] sheets)
		{
			if (sheets == null || sheets.Length == 0)
				throw new Exception("Sheet data not specified.");

			Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
			try
			{
				app.Visible = false;
				app.DisplayAlerts = false;

				Microsoft.Office.Interop.Excel.Workbooks workbooks = app.Workbooks;
				try
				{
					Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add();
					try
					{
						while (sheets.Length < workbook.Worksheets.Count)
						{
							Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count];
							try
							{
								worksheet.Delete();
							}
							finally
							{
								Marshal.ReleaseComObject(worksheet);
							}
						}

						for (int index = 0; index < sheets.Length; index++)
						{
							Microsoft.Office.Interop.Excel.Worksheet worksheet;

							if (index < workbook.Worksheets.Count)
							{
								worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[index + 1];
							}
							else
							{
								Microsoft.Office.Interop.Excel.Worksheet afterWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count];
								try
								{
									worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(After: afterWorksheet);
								}
								finally
								{
									Marshal.ReleaseComObject(afterWorksheet);
								}
							}
							try
							{
								WriteSheet(worksheet, sheets[index]);
							}
							finally
							{
								Marshal.ReleaseComObject(worksheet);
							}
						}

						workbook.SaveAs(excelFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
					}
					finally
					{
						workbook.Close(false);
						Marshal.ReleaseComObject(workbook);
					}
				}
				finally
				{
					Marshal.ReleaseComObject(workbooks);
				}
			}
			finally
			{
				app.Quit();
				Marshal.ReleaseComObject(app);

				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
		}

		private static void WriteSheet(Microsoft.Office.Interop.Excel.Worksheet worksheet, SheetData sheet)
		{
			worksheet.Name = MakeSafeSheetName(sheet.Name);

			int rowCount = sheet.Rows.Length;
			int columnCount = rowCount == 0 ? 0 : sheet.Rows.Max(row => row.Length);

			if (rowCount == 0 || columnCount == 0)
				return;

			object[,] values = new object[rowCount, columnCount];

			for (int row = 0; row < rowCount; row++)
			{
				for (int column = 0; column < sheet.Rows[row].Length; column++)
				{
					values[row, column] = sheet.Rows[row][column];
				}
			}

			Microsoft.Office.Interop.Excel.Range start = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 1];
			Microsoft.Office.Interop.Excel.Range end = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowCount, columnCount];
			Microsoft.Office.Interop.Excel.Range range = worksheet.Range[start, end];
			try
			{
				range.Value2 = values;
			}
			finally
			{
				Marshal.ReleaseComObject(range);
				Marshal.ReleaseComObject(end);
				Marshal.ReleaseComObject(start);
			}
		}

		public static string MakeSafeSheetName(string name)
		{
			if (string.IsNullOrEmpty(name))
				name = "Sheet";

			foreach (char chr in new char[] { '\\', '/', '?', '*', '[', ']', ':' })
			{
				name = name.Replace(chr, '_');
			}
			name = name.Trim();

			if (name.Length == 0)
				name = "Sheet";

			if (31 < name.Length)
				name = name.Substring(0, 31);

			return name;
		}

		public static string[] GetPrinterNames()
		{
			string[] printerNames = new string[PrinterSettings.InstalledPrinters.Count];

			PrinterSettings.InstalledPrinters.CopyTo(printerNames, 0);

			return printerNames;
		}

		public static void Print(string excelFile, string optionalPrinterName = null)
		{
			ExecuteWorkbook(excelFile, (app, workbook) =>
			{
				workbook.PrintOut(
					1,
					9999, // 終了ページ、全ページを印刷させるため十分大きな値にしておく
					1,
					false,
					optionalPrinterName ?? Type.Missing
					);

				Thread.Sleep(5000); // 印刷キュー投入後すぐにエクセルを閉じると失敗する場合があるので待機
			});
		}
	}
}
