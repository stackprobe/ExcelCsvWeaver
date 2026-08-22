using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using HLTStudio.ECWArguments;
using HLTStudio.Commons;
using HLTStudio.Tools;

namespace HLTStudio.ECWOperations
{
	public class ECWeaverProcessor
	{
		private ECWeaverArgs Args;

		public void Run(ECWeaverArgs args)
		{
			if (args == null)
				throw new ArgumentNullException("args");

			this.Args = args;

			if (this.Args.HasOption(ECWeaverArgConsts.Options.Help))
			{
				this.ShowHelp(this.Args.Operation);
				return;
			}
			if (this.Args.HasOption(ECWeaverArgConsts.Options.Version))
			{
				this.ShowVersion();
				return;
			}

			string command = this.Args.Operation;

			if (string.IsNullOrEmpty(command))
			{
				this.ShowHelp(null);
				return;
			}

			switch (command.ToLowerInvariant())
			{
				case ECWeaverArgConsts.Operations.Help:
					this.ShowHelp(this.Args.Arguments.FirstOrDefault());
					break;

				case ECWeaverArgConsts.Operations.Version:
					this.ShowVersion();
					break;

				case ECWeaverArgConsts.Operations.ExcelToCsv:
					this.ExcelToCsv(CsvFileWriter.DELIMITER_COMMA);
					break;

				case ECWeaverArgConsts.Operations.ExcelToTsv:
					this.ExcelToCsv(CsvFileWriter.DELIMITER_TAB);
					break;

				case ECWeaverArgConsts.Operations.ExcelToPdf:
					this.ExcelToPdf();
					break;

				case ECWeaverArgConsts.Operations.CsvInfo:
					this.CsvInfo();
					break;

				case ECWeaverArgConsts.Operations.CsvSelectColumns:
					this.CsvSelectColumns();
					break;

				case ECWeaverArgConsts.Operations.CsvFilterRows:
					this.CsvFilterRows();
					break;

				case ECWeaverArgConsts.Operations.CsvReplace:
					this.CsvReplace();
					break;

				case ECWeaverArgConsts.Operations.CsvMerge:
					this.CsvMerge();
					break;

				case ECWeaverArgConsts.Operations.CsvSort:
					this.CsvSort();
					break;

				case ECWeaverArgConsts.Operations.CsvUnique:
					this.CsvUnique();
					break;

				case ECWeaverArgConsts.Operations.ExcelListSheets:
					this.ExcelListSheets();
					break;

				case ECWeaverArgConsts.Operations.ExcelExtractPictures:
					this.ExcelExtractPictures();
					break;

				case ECWeaverArgConsts.Operations.ExcelReplacePicture:
					this.ExcelReplacePicture();
					break;

				case ECWeaverArgConsts.Operations.Printers:
					this.Printers();
					break;

				case ECWeaverArgConsts.Operations.Print:
					this.Print();
					break;

				case ECWeaverArgConsts.Operations.CsvToExcel:
				case ECWeaverArgConsts.Operations.CsvsToExcel:
				case ECWeaverArgConsts.Operations.Weave:
				case ECWeaverArgConsts.Operations.ExcelInfo:
				case ECWeaverArgConsts.Operations.ExcelReplaceText:
				case ECWeaverArgConsts.Operations.ExcelReplacePlaceholder:
				case ECWeaverArgConsts.Operations.CsvValidate:
				case ECWeaverArgConsts.Operations.ExcelValidate:
				case ECWeaverArgConsts.Operations.CsvDiff:
				case ECWeaverArgConsts.Operations.ExcelDiff:
				case ECWeaverArgConsts.Operations.RunScript:
					throw new Exception("This command is not implemented in ECWeaver yet: " + command);

				default:
					throw new Exception("Unknown command: " + command);
			}
		}

		private void ShowHelp(string command)
		{
			if (string.IsNullOrEmpty(command))
			{
				this.WriteOutput("Usage:");
				this.WriteOutput("  ECWeaver.exe <command> [options] [arguments]");
				this.WriteOutput("");
				this.WriteOutput("Commands:");
				this.WriteOutput("  help, version");
				this.WriteOutput("  excel-to-csv, excel-to-tsv, excel-to-pdf");
				this.WriteOutput("  excel-list-sheets, excel-extract-pictures, excel-replace-picture");
				this.WriteOutput("  csv-info, csv-select-columns, csv-filter-rows, csv-replace");
				this.WriteOutput("  csv-merge, csv-sort, csv-unique");
				this.WriteOutput("  printers, print");
				this.WriteOutput("");
				this.WriteOutput("Common options:");
				this.WriteOutput("  --overwrite");
				this.WriteOutput("  --encoding auto|sjis|utf8|utf8bom|utf16le");
				this.WriteOutput("  --delimiter comma|tab|space|<char>");
				this.WriteOutput("  --silent");
				return;
			}

			switch (command.ToLowerInvariant())
			{
				case ECWeaverArgConsts.Operations.ExcelToCsv:
					this.WriteOutput("Usage: ECWeaver.exe excel-to-csv [--sheet <name-or-index>] [--overwrite] <input-excel> <output-dir-or-csv>");
					break;

				case ECWeaverArgConsts.Operations.ExcelToTsv:
					this.WriteOutput("Usage: ECWeaver.exe excel-to-tsv [--sheet <name-or-index>] [--overwrite] <input-excel> <output-dir-or-tsv>");
					break;

				case ECWeaverArgConsts.Operations.ExcelToPdf:
					this.WriteOutput("Usage: ECWeaver.exe excel-to-pdf [--overwrite] <input-excel> <output-pdf>");
					break;

				case ECWeaverArgConsts.Operations.CsvInfo:
					this.WriteOutput("Usage: ECWeaver.exe csv-info [--encoding <encoding>] [--delimiter <delimiter>] <input-csv>");
					break;

				case ECWeaverArgConsts.Operations.CsvSelectColumns:
					this.WriteOutput("Usage: ECWeaver.exe csv-select-columns (--columns <indexes>|--headers <names>) [--overwrite] <input-csv> <output-csv>");
					break;

				case ECWeaverArgConsts.Operations.CsvFilterRows:
					this.WriteOutput("Usage: ECWeaver.exe csv-filter-rows (--column <index>|--header <name>) (--equals <text>|--contains <text>|--regex <pattern>) [--invert] [--overwrite] <input-csv> <output-csv>");
					break;

				case ECWeaverArgConsts.Operations.CsvReplace:
					this.WriteOutput("Usage: ECWeaver.exe csv-replace (--from <text>|--regex <pattern>) --to <text> [--column <index>|--header <name>] [--overwrite] <input-csv> <output-csv>");
					break;

				default:
					this.WriteOutput("No detailed help for command: " + command);
					break;
			}
		}

		private void ShowVersion()
		{
			Version version = Assembly.GetExecutingAssembly().GetName().Version;
			this.WriteOutput("ECWeaver " + version);
		}

		private void ExcelToCsv(char defaultDelimiter)
		{
			this.CheckArgCount(2);
			this.CheckAppEngine();

			string inputExcel = this.Args.Arguments[0];
			string outputPath = this.Args.Arguments[1];
			string sheet = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Sheet);
			string sheets = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Sheets);
			char delimiter = this.GetDelimiter(defaultDelimiter, outputPath);
			Encoding encoding = this.GetOutputEncoding();
			ExcelAppTools.Sheet[] loadedSheets = ExcelAppTools.LoadSheets(inputExcel);

			if (!string.IsNullOrEmpty(sheet))
			{
				if (!string.IsNullOrEmpty(sheets))
					throw new Exception("--sheet and --sheets cannot be used together.");

				ExcelAppTools.Sheet selectedSheet = this.GetSheet(loadedSheets, sheet);
				this.WriteRowsToNewFile(outputPath, selectedSheet.Rows, encoding, delimiter);
			}
			else
			{
				ExcelAppTools.Sheet[] selectedSheets = loadedSheets;

				if (!string.IsNullOrEmpty(sheets))
					selectedSheets = this.GetSheets(loadedSheets, sheets);

				this.PrepareOutputDir(outputPath);

				for (int index = 0; index < selectedSheets.Length; index++)
				{
					string outputFile = Path.Combine(outputPath, string.Format("{0:D4}{1}", index + 1, defaultDelimiter == CsvFileWriter.DELIMITER_TAB ? ".tsv" : ".csv"));
					CsvFileWriter.WriteRows(outputFile, false, encoding, delimiter, selectedSheets[index].Rows);
				}
				File.WriteAllLines(Path.Combine(outputPath, "sheet-names.txt"), selectedSheets.Select(v => v.Name).ToArray(), Encoding.UTF8);
			}
		}

		private void ExcelToPdf()
		{
			this.CheckArgCount(2);
			this.CheckAppEngine();

			string inputExcel = this.Args.Arguments[0];
			string outputPdf = this.Args.Arguments[1];

			this.PrepareOutputFile(outputPdf);
			ExcelAppTools.ToPDF(inputExcel, outputPdf);
		}

		private void CsvInfo()
		{
			this.CheckArgCount(1);
			this.CheckNoCsvEngine();

			string[][] rows = this.ReadCsv(this.Args.Arguments[0]);
			int rowCount = rows.Length;
			int minColumns = rowCount == 0 ? 0 : rows.Min(row => row.Length);
			int maxColumns = rowCount == 0 ? 0 : rows.Max(row => row.Length);
			int emptyRows = rows.Count(row => row.All(cell => cell == ""));

			this.WriteOutput("Rows: " + rowCount);
			this.WriteOutput("MinColumns: " + minColumns);
			this.WriteOutput("MaxColumns: " + maxColumns);
			this.WriteOutput("EmptyRows: " + emptyRows);
		}

		private void CsvSelectColumns()
		{
			this.CheckArgCount(2);
			this.CheckNoCsvEngine();

			bool hasColumns = this.Args.HasOption(ECWeaverArgConsts.Options.Columns);
			bool hasHeaders = this.Args.HasOption(ECWeaverArgConsts.Options.Headers);

			if (hasColumns == hasHeaders)
				throw new Exception("Specify either --columns or --headers.");

			string[][] rows = this.ReadCsv(this.Args.Arguments[0]);
			int[] indexes = hasColumns ?
				this.ParseColumnIndexes(this.Args.GetOptionValue(ECWeaverArgConsts.Options.Columns)) :
				this.GetHeaderIndexes(rows, this.ParseCsvList(this.Args.GetOptionValue(ECWeaverArgConsts.Options.Headers)));

			this.CheckColumnIndexes(rows, indexes);

			string[][] destRows = rows
				.Select(row => indexes.Select(index => index < row.Length ? row[index] : "").ToArray())
				.ToArray();

			this.WriteRowsToNewFile(this.Args.Arguments[1], destRows, this.GetOutputEncoding(), this.GetDelimiter(CsvFileWriter.DELIMITER_COMMA, this.Args.Arguments[1]));
		}

		private void CsvFilterRows()
		{
			this.CheckArgCount(2);
			this.CheckNoCsvEngine();

			string[][] rows = this.ReadCsv(this.Args.Arguments[0]);
			int columnIndex = this.GetSingleColumnIndex(rows);
			string equalsValue = this.Args.GetOptionValue(ECWeaverArgConsts.Options.EqualsCondition);
			string containsValue = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Contains);
			string regexValue = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Regex);
			int conditionCount = new string[] { equalsValue, containsValue, regexValue }.Count(value => value != null);

			if (conditionCount != 1)
				throw new Exception("Specify one of --equals, --contains, or --regex.");

			Regex regex = regexValue == null ? null : new Regex(regexValue);
			bool invert = this.Args.HasOption(ECWeaverArgConsts.Options.Invert);
			bool hasHeader = this.Args.HasOption(ECWeaverArgConsts.Options.HasHeader) || this.Args.HasOption(ECWeaverArgConsts.Options.Header);
			List<string[]> destRows = new List<string[]>();

			for (int index = 0; index < rows.Length; index++)
			{
				if (index == 0 && hasHeader)
				{
					destRows.Add(rows[index]);
					continue;
				}

				string cell = columnIndex < rows[index].Length ? rows[index][columnIndex] : "";
				bool match =
					equalsValue != null ? cell == equalsValue :
					containsValue != null ? cell.Contains(containsValue) :
					regex.IsMatch(cell);

				if (invert)
					match = !match;

				if (match)
					destRows.Add(rows[index]);
			}

			this.WriteRowsToNewFile(this.Args.Arguments[1], destRows.ToArray(), this.GetOutputEncoding(), this.GetDelimiter(CsvFileWriter.DELIMITER_COMMA, this.Args.Arguments[1]));
		}

		private void CsvReplace()
		{
			this.CheckArgCount(2);
			this.CheckNoCsvEngine();

			string from = this.Args.GetOptionValue(ECWeaverArgConsts.Options.From);
			string regexValue = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Regex);
			string to = this.Args.GetOptionValue(ECWeaverArgConsts.Options.To);

			if (to == null)
				throw new Exception("Missing --to.");

			if ((from == null) == (regexValue == null))
				throw new Exception("Specify either --from or --regex.");

			string[][] rows = this.ReadCsv(this.Args.Arguments[0]);
			int? columnIndex = null;

			if (this.Args.HasOption(ECWeaverArgConsts.Options.Column) || this.Args.HasOption(ECWeaverArgConsts.Options.Header))
				columnIndex = this.GetSingleColumnIndex(rows);

			Regex regex = regexValue == null ? null : new Regex(regexValue);

			string[][] destRows = rows
				.Select(row => row.Select((cell, index) =>
				{
					if (columnIndex != null && columnIndex.Value != index)
						return cell;

					return from != null ? cell.Replace(from, to) : regex.Replace(cell, to);
				}).ToArray())
				.ToArray();

			this.WriteRowsToNewFile(this.Args.Arguments[1], destRows, this.GetOutputEncoding(), this.GetDelimiter(CsvFileWriter.DELIMITER_COMMA, this.Args.Arguments[1]));
		}

		private void CsvMerge()
		{
			this.CheckArgCount(2);
			this.CheckNoCsvEngine();

			string inputDir = this.Args.Arguments[0];
			string outputCsv = this.Args.Arguments[1];
			string pattern = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Pattern, "*.csv");

			if (!Directory.Exists(inputDir))
				throw new Exception("Input directory not found: " + inputDir);

			string[] files = Directory.GetFiles(inputDir, pattern).OrderBy(SCommon.CompIgnoreCase).ToArray();
			List<string[]> destRows = new List<string[]>();
			bool skipHeader = this.Args.HasOption(ECWeaverArgConsts.Options.SkipHeader);

			for (int fileIndex = 0; fileIndex < files.Length; fileIndex++)
			{
				string[][] rows = this.ReadCsv(files[fileIndex]);

				if (skipHeader && 0 < fileIndex && 0 < rows.Length)
					rows = rows.Skip(1).ToArray();

				destRows.AddRange(rows);
			}

			this.WriteRowsToNewFile(outputCsv, destRows.ToArray(), this.GetOutputEncoding(), this.GetDelimiter(CsvFileWriter.DELIMITER_COMMA, outputCsv));
		}

		private void CsvSort()
		{
			this.CheckArgCount(2);
			this.CheckNoCsvEngine();

			string[][] rows = this.ReadCsv(this.Args.Arguments[0]);
			int columnIndex = this.GetSingleColumnIndex(rows);
			bool hasHeader = this.Args.HasOption(ECWeaverArgConsts.Options.HasHeader) || this.Args.HasOption(ECWeaverArgConsts.Options.Header);
			List<string[]> bodyRows = rows.Skip(hasHeader ? 1 : 0).ToList();
			bool numeric = this.Args.HasOption(ECWeaverArgConsts.Options.Numeric);
			bool desc = this.Args.HasOption(ECWeaverArgConsts.Options.Desc);

			bodyRows.Sort((a, b) =>
			{
				string aCell = columnIndex < a.Length ? a[columnIndex] : "";
				string bCell = columnIndex < b.Length ? b[columnIndex] : "";
				int ret;

				if (numeric)
				{
					double aValue;
					double bValue;

					if (!double.TryParse(aCell, out aValue))
						aValue = 0.0;

					if (!double.TryParse(bCell, out bValue))
						bValue = 0.0;

					ret = aValue.CompareTo(bValue);
				}
				else
				{
					ret = SCommon.Comp(aCell, bCell);
				}
				return desc ? -ret : ret;
			});

			string[][] destRows = hasHeader && 0 < rows.Length ?
				new string[][] { rows[0] }.Concat(bodyRows).ToArray() :
				bodyRows.ToArray();

			this.WriteRowsToNewFile(this.Args.Arguments[1], destRows, this.GetOutputEncoding(), this.GetDelimiter(CsvFileWriter.DELIMITER_COMMA, this.Args.Arguments[1]));
		}

		private void CsvUnique()
		{
			this.CheckArgCount(2);
			this.CheckNoCsvEngine();

			string[][] rows = this.ReadCsv(this.Args.Arguments[0]);
			bool hasColumns = this.Args.HasOption(ECWeaverArgConsts.Options.Columns);
			bool hasHeaders = this.Args.HasOption(ECWeaverArgConsts.Options.Headers);

			if (hasColumns && hasHeaders)
				throw new Exception("--columns and --headers cannot be used together.");

			int[] indexes = null;

			if (hasColumns)
				indexes = this.ParseColumnIndexes(this.Args.GetOptionValue(ECWeaverArgConsts.Options.Columns));
			else if (hasHeaders)
				indexes = this.GetHeaderIndexes(rows, this.ParseCsvList(this.Args.GetOptionValue(ECWeaverArgConsts.Options.Headers)));

			if (indexes != null)
				this.CheckColumnIndexes(rows, indexes);

			HashSet<string> seen = new HashSet<string>();
			List<string[]> destRows = new List<string[]>();

			foreach (string[] row in rows)
			{
				string key = indexes == null ?
					string.Join("\t", row.Select(EscapeKey)) :
					string.Join("\t", indexes.Select(index => EscapeKey(index < row.Length ? row[index] : "")));

				if (seen.Add(key))
					destRows.Add(row);
			}

			this.WriteRowsToNewFile(this.Args.Arguments[1], destRows.ToArray(), this.GetOutputEncoding(), this.GetDelimiter(CsvFileWriter.DELIMITER_COMMA, this.Args.Arguments[1]));
		}

		private void ExcelListSheets()
		{
			this.CheckArgCount(1);
			this.CheckAppEngine();

			string inputExcel = this.Args.Arguments[0];
			ExcelAppTools.Sheet[] sheets = ExcelAppTools.LoadSheets(inputExcel);
			string[] lines = sheets.Select((sheet, index) => (index + 1) + "\t" + sheet.Name).ToArray();
			string output = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Output);

			if (string.IsNullOrEmpty(output))
			{
				foreach (string line in lines)
					this.WriteOutput(line);
			}
			else
			{
				this.WriteLinesToNewFile(output, lines, Encoding.UTF8);
			}
		}

		private void ExcelExtractPictures()
		{
			this.CheckArgCount(2);
			this.CheckZipEngine();

			string inputExcel = this.Args.Arguments[0];
			string outputDir = this.Args.Arguments[1];
			int index = 0;

			this.PrepareOutputDir(outputDir);

			ExcelTools.CollectPicture(inputExcel, pictureFile =>
			{
				index++;

				string outputFile = Path.Combine(outputDir, string.Format("{0:D4}{1}", index, Path.GetExtension(pictureFile)));
				File.Copy(pictureFile, outputFile);
			});
		}

		private void ExcelReplacePicture()
		{
			this.CheckArgCount(3);
			this.CheckZipEngine();

			string inputExcel = this.Args.Arguments[0];
			string outputExcel = this.Args.Arguments[1];
			string pictureFile = this.Args.Arguments[2];
			string indexText = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Index);
			int? targetIndex = null;

			if (indexText != null)
				targetIndex = this.ParsePositiveInt(indexText, "--index");

			if (!File.Exists(pictureFile))
				throw new Exception("Picture file not found: " + pictureFile);

			this.PrepareOutputFile(outputExcel);

			int currentIndex = 0;
			ExcelTools.ReplacePicture(inputExcel, outputExcel, embeddedPictureFile =>
			{
				currentIndex++;

				if (targetIndex == null || targetIndex.Value == currentIndex)
					File.Copy(pictureFile, embeddedPictureFile, true);
			});
		}

		private void Printers()
		{
			this.CheckArgCount(0);
			this.CheckAppEngine();

			foreach (string printerName in ExcelAppTools.GetPrinterNames())
				this.WriteOutput(printerName);
		}

		private void Print()
		{
			this.CheckArgCount(1);
			this.CheckAppEngine();

			ExcelAppTools.Print(this.Args.Arguments[0], this.Args.GetOptionValue(ECWeaverArgConsts.Options.Printer));
		}

		private void CheckArgCount(int count)
		{
			if (this.Args.Arguments.Length != count)
				throw new Exception("Bad argument count. Expected: " + count + ", Actual: " + this.Args.Arguments.Length);
		}

		private void CheckNoCsvEngine()
		{
			if (this.Args.HasOption(ECWeaverArgConsts.Options.Engine))
				throw new Exception("--engine cannot be used for this command.");
		}

		private void CheckAppEngine()
		{
			string engine = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Engine, ECWeaverArgConsts.Values.Auto);

			if (!engine.EqualsIgnoreCase(ECWeaverArgConsts.Values.Auto) && !engine.EqualsIgnoreCase(ECWeaverArgConsts.Values.App))
				throw new Exception("Unsupported engine for ECWeaver: " + engine);
		}

		private void CheckZipEngine()
		{
			string engine = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Engine, ECWeaverArgConsts.Values.Auto);

			if (!engine.EqualsIgnoreCase(ECWeaverArgConsts.Values.Auto) && !engine.EqualsIgnoreCase(ECWeaverArgConsts.Values.Zip))
				throw new Exception("Unsupported engine for this command: " + engine);
		}

		private string[][] ReadCsv(string file)
		{
			if (!File.Exists(file))
				throw new Exception("Input file not found: " + file);

			Encoding encoding = this.GetInputEncoding();
			char delimiter = this.GetDelimiter(CsvFileReader.DELIMITER_COMMA, file);

			return encoding == null ?
				CsvFileReader.ReadToEnd(file, GetAutoDelimiter(file, delimiter)) :
				CsvFileReader.ReadToEnd(file, encoding, delimiter);
		}

		private void WriteRowsToNewFile(string file, string[][] rows, Encoding encoding, char delimiter)
		{
			this.WriteFileWithTemp(file, tempFile => CsvFileWriter.WriteRows(tempFile, false, encoding, delimiter, rows));
		}

		private void WriteLinesToNewFile(string file, string[] lines, Encoding encoding)
		{
			this.WriteFileWithTemp(file, tempFile => File.WriteAllLines(tempFile, lines, encoding));
		}

		private void WriteFileWithTemp(string file, Action<string> writer)
		{
			string parentDir = Path.GetDirectoryName(Path.GetFullPath(file));

			if (!Directory.Exists(parentDir))
				SCommon.CreateDir(parentDir);

			this.PrepareOutputFile(file);

			using (WorkingDir wd = new WorkingDir())
			{
				string tempFile = wd.MakePath() + Path.GetExtension(file);

				writer(tempFile);
				SCommon.EnsureMoveFile(tempFile, file);
			}
		}

		private void PrepareOutputFile(string file)
		{
			if (string.IsNullOrEmpty(file))
				throw new Exception("Bad output file.");

			if (!SCommon.IsExistsPath(file))
				return;

			if (!this.Args.HasOption(ECWeaverArgConsts.Options.Overwrite))
				throw new Exception("Output path already exists: " + file);

			SCommon.DeletePath(file);
		}

		private void PrepareOutputDir(string dir)
		{
			if (string.IsNullOrEmpty(dir))
				throw new Exception("Bad output directory.");

			if (SCommon.IsExistsPath(dir))
			{
				if (!this.Args.HasOption(ECWeaverArgConsts.Options.Overwrite))
					throw new Exception("Output path already exists: " + dir);

				SCommon.DeletePath(dir);
			}
			SCommon.CreateDir(dir);
		}

		private Encoding GetInputEncoding()
		{
			string value = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Encoding, ECWeaverArgConsts.Values.Auto);

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Auto))
				return null;

			return ParseEncoding(value);
		}

		private Encoding GetOutputEncoding()
		{
			string value = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Encoding, ECWeaverArgConsts.Values.Sjis);

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Auto))
				return SCommon.ENCODING_SJIS;

			return ParseEncoding(value);
		}

		private static Encoding ParseEncoding(string value)
		{
			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Sjis))
				return SCommon.ENCODING_SJIS;

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Utf8))
				return new UTF8Encoding(false);

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Utf8Bom))
				return new UTF8Encoding(true);

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Utf16Le))
				return Encoding.Unicode;

			throw new Exception("Unknown encoding: " + value);
		}

		private char GetDelimiter(char defaultDelimiter, string file)
		{
			string value = this.Args.GetOptionValue(ECWeaverArgConsts.Options.Delimiter);

			if (value == null)
				return GetDefaultDelimiter(file, defaultDelimiter);

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Comma))
				return CsvFileReader.DELIMITER_COMMA;

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Tab))
				return CsvFileReader.DELIMITER_TAB;

			if (value.EqualsIgnoreCase(ECWeaverArgConsts.Values.Space))
				return CsvFileReader.DELIMITER_SPACE;

			if (value.Length == 1)
				return value[0];

			throw new Exception("Bad delimiter: " + value);
		}

		private static char GetAutoDelimiter(string file, char fallbackDelimiter)
		{
			return GetDefaultDelimiter(file, fallbackDelimiter);
		}

		private static char GetDefaultDelimiter(string file, char fallbackDelimiter)
		{
			string ext = Path.GetExtension(file).ToLowerInvariant();

			if (ext == ".csv")
				return CsvFileReader.DELIMITER_COMMA;

			if (ext == ".tsv")
				return CsvFileReader.DELIMITER_TAB;

			if (ext == ".ssv")
				return CsvFileReader.DELIMITER_SPACE;

			return fallbackDelimiter;
		}

		private ExcelAppTools.Sheet GetSheet(ExcelAppTools.Sheet[] sheets, string selector)
		{
			if (Regex.IsMatch(selector, "^[0-9]+$"))
			{
				int index = this.ParsePositiveInt(selector, "--sheet") - 1;

				if (sheets.Length <= index)
					throw new Exception("Sheet index out of range: " + selector);

				return sheets[index];
			}

			ExcelAppTools.Sheet sheet = sheets.FirstOrDefault(v => v.Name == selector);

			if (sheet == null)
				throw new Exception("Sheet not found: " + selector);

			return sheet;
		}

		private ExcelAppTools.Sheet[] GetSheets(ExcelAppTools.Sheet[] sheets, string selectors)
		{
			return this.ParseCsvList(selectors)
				.Select(selector => this.GetSheet(sheets, selector))
				.ToArray();
		}

		private int GetSingleColumnIndex(string[][] rows)
		{
			bool hasColumn = this.Args.HasOption(ECWeaverArgConsts.Options.Column);
			bool hasHeader = this.Args.HasOption(ECWeaverArgConsts.Options.Header);

			if (hasColumn == hasHeader)
				throw new Exception("Specify either --column or --header.");

			if (hasColumn)
				return this.ParsePositiveInt(this.Args.GetOptionValue(ECWeaverArgConsts.Options.Column), "--column") - 1;

			return this.GetHeaderIndexes(rows, new string[] { this.Args.GetOptionValue(ECWeaverArgConsts.Options.Header) })[0];
		}

		private int[] GetHeaderIndexes(string[][] rows, string[] headers)
		{
			if (rows.Length == 0)
				throw new Exception("Header row not found.");

			return headers.Select(header =>
			{
				int index = Array.IndexOf(rows[0], header);

				if (index == -1)
					throw new Exception("Header not found: " + header);

				return index;
			}).ToArray();
		}

		private int[] ParseColumnIndexes(string text)
		{
			return this.ParseCsvList(text)
				.Select(value => this.ParsePositiveInt(value, "--columns") - 1)
				.ToArray();
		}

		private string[] ParseCsvList(string text)
		{
			if (string.IsNullOrEmpty(text))
				throw new Exception("Bad list value.");

			return text.Split(',')
				.Select(value => value.Trim())
				.Where(value => value != "")
				.ToArray();
		}

		private int ParsePositiveInt(string text, string name)
		{
			int value;

			if (!int.TryParse(text, out value) || value < 1)
				throw new Exception("Bad " + name + ": " + text);

			return value;
		}

		private void CheckColumnIndexes(string[][] rows, int[] indexes)
		{
			if (indexes.Length == 0)
				throw new Exception("No columns specified.");

			int maxColumnCount = rows.Length == 0 ? 0 : rows.Max(row => row.Length);

			foreach (int index in indexes)
				if (index < 0 || maxColumnCount <= index)
					throw new Exception("Column index out of range: " + (index + 1));
		}

		private static string EscapeKey(string value)
		{
			return value.Replace("\\", "\\\\").Replace("\t", "\\t");
		}

		private void WriteOutput(string message)
		{
			if (!this.Args.HasOption(ECWeaverArgConsts.Options.Silent))
				Console.WriteLine(message);
		}
	}
}
