using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using HLTStudio.Commons;
using HLTStudio.ECWArguments;
using HLTStudio.ECWOperations;
using HLTStudio.Tools;

namespace HLTStudio.Tests
{
	public class Test0001
	{
#if DEBUG
		private const string ResourceDir = @"C:\home\res\ExcelCsvWeaver\ECWeaver\Test0001";
		private const string BaseDir = @"C:\temp\ECWeaver";
		private const string InputDir = ResourceDir + @"\input";
		private const string OutputDir = BaseDir + @"\output";

		private readonly Encoding Sjis = Encoding.GetEncoding(932);
		private readonly Encoding Utf8NoBom = new UTF8Encoding(false);
		private readonly Encoding Utf8Bom = new UTF8Encoding(true);

		public void Test01()
		{
			try
			{
				// 人間が F5 で実行するときコンソールがすぐ閉じると結果を確認しづらいため、待機時間の強制 0 秒化は無効にしている。
				//SCommon.Pause_WaitSeconds = 0;

				SCommon.DeleteAndCreateDir(BaseDir);
				Directory.CreateDirectory(OutputDir);

				// C:\home\res\ 配下のテストデータは永続管理するため、初回生成後は毎回作り直さない。
				//this.PrepareTestData();
				if (!File.Exists(this.InputFile("text-replace.xlsx")) || !File.Exists(this.InputFile("placeholder-map.csv")))
					this.PrepareTestData();

				this.TestHelpVersionAndCommandErrors();
				this.TestResponseFileOption();
				this.TestCsvInfoCountsRowsAndColumns();
				this.TestCsvSelectColumnsByIndexesAndHeaders();
				this.TestCsvSelectColumnsRejectsBadOptions();
				this.TestCsvFilterRowsByEqualsContainsRegexAndInvert();
				this.TestCsvFilterRowsRejectsBadOptions();
				this.TestCsvReplaceAllCellsAndSelectedColumns();
				this.TestCsvReplaceRejectsBadOptions();
				this.TestCsvMergeOrdersFilesAndSkipsRepeatedHeaders();
				this.TestCsvSortStringAndNumericValues();
				this.TestCsvUniqueRemovesDuplicateRowsAndKeys();
				this.TestDelimiterAndEncodingOptions();
				this.TestOverwriteProtection();
				this.TestExcelReplaceTextAndPlaceholder();
				this.TestExcelExtractPicturesByZipEngine();
				this.TestExcelReplacePictureByZipEngine();
				this.TestZipEngineCommandsRejectBadOptions();

				this.ShowSuccessBanner();
			}
			catch (Exception ex)
			{
				this.ShowFailure(ex);
			}
		}

		/// <summary>
		/// help/version の標準出力と、未知コマンド・未実装コマンドの異常系を確認する。
		/// </summary>
		private void TestHelpVersionAndCommandErrors()
		{
			string helpOutput = this.CaptureOutput(new string[] { "help" });
			this.AssertContains(helpOutput, "Usage:", "help output");
			this.AssertContains(helpOutput, "csv-info", "help output");

			string commandHelpOutput = this.CaptureOutput(new string[] { "help", "csv-replace" });
			this.AssertContains(commandHelpOutput, "Usage: ECWeaver.exe csv-replace", "command help output");

			string versionOutput = this.CaptureOutput(new string[] { "version" });
			this.AssertContains(versionOutput, "ECWeaver ", "version output");

			this.AssertThrows(() => this.Run(new string[] { "not-a-command" }), "Unknown command");
			this.AssertThrows(() => this.Run(new string[] { "csv-to-excel", this.InputFile("basic.csv"), this.OutputFile("not-implemented.xlsx") }), "not implemented");
		}

		private void TestResponseFileOption()
		{
			string responseFile = this.OutputFile("response.txt");
			string responseEqualsFile = this.OutputFile("response-equals.txt");
			string output = this.OutputFile("response-select.csv");
			string equalsOutput = this.OutputFile("response-equals-select.csv");

			File.WriteAllLines(responseFile, new string[]
			{
				"csv-select-columns",
				"--columns",
				"1,3",
				this.InputFile("basic.csv"),
				output,
			}, this.Sjis);

			this.Run(new string[] { "--response", responseFile });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Price" },
					new string[] { "1", "100" },
					new string[] { "2", "200" },
					new string[] { "3", "150" },
					new string[] { "4", "200" },
				},
				this.ReadCsv(output),
				"response file"
				);

			File.WriteAllLines(responseEqualsFile, new string[]
			{
				"csv-select-columns",
				"--headers",
				"Name,Category",
				this.InputFile("basic.csv"),
				equalsOutput,
			}, this.Sjis);

			this.Run(new string[] { "--response=" + responseEqualsFile });
			this.AssertRows(
				new string[][]
				{
					new string[] { "Name", "Category" },
					new string[] { "Apple", "Fruit" },
					new string[] { "Carrot", "Vegetable" },
					new string[] { "Banana", "Fruit" },
					new string[] { "Apple", "Fruit" },
				},
				this.ReadCsv(equalsOutput),
				"response file equals"
				);

			this.AssertThrows(() => this.Run(new string[] { "--response" }), "Missing command line option value: --response");
			this.AssertThrows(() => this.Run(new string[] { "--response", this.OutputFile("missing-response.txt") }), "Response file not found");
		}

		/// <summary>
		/// csv-info が行数、最小列数、最大列数、空行数を正しく出力することを確認する。
		/// </summary>
		private void TestCsvInfoCountsRowsAndColumns()
		{
			string output = this.CaptureOutput(new string[] { "csv-info", this.InputFile("info.csv") });

			this.AssertContains(output, "Rows: 4", "csv-info output");
			this.AssertContains(output, "MinColumns: 1", "csv-info output");
			this.AssertContains(output, "MaxColumns: 4", "csv-info output");
			this.AssertContains(output, "EmptyRows: 1", "csv-info output");

			string silentOutput = this.CaptureOutput(new string[] { "csv-info", "--silent", this.InputFile("info.csv") });
			this.AssertEquals("", silentOutput, "--silent output");
		}

		/// <summary>
		/// csv-select-columns が列番号指定とヘッダー名指定の両方で指定列だけを出力することを確認する。
		/// </summary>
		private void TestCsvSelectColumnsByIndexesAndHeaders()
		{
			string outputByColumns = this.OutputFile("select-by-columns.csv");

			this.Run(new string[] { "csv-select-columns", "--columns", "1,3", this.InputFile("basic.csv"), outputByColumns });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Price" },
					new string[] { "1", "100" },
					new string[] { "2", "200" },
					new string[] { "3", "150" },
					new string[] { "4", "200" },
				},
				this.ReadCsv(outputByColumns),
				"select by columns"
				);

			string outputByHeaders = this.OutputFile("select-by-headers.csv");

			this.Run(new string[] { "csv-select-columns", "--headers", "Name,Category", this.InputFile("basic.csv"), outputByHeaders });
			this.AssertRows(
				new string[][]
				{
					new string[] { "Name", "Category" },
					new string[] { "Apple", "Fruit" },
					new string[] { "Carrot", "Vegetable" },
					new string[] { "Banana", "Fruit" },
					new string[] { "Apple", "Fruit" },
				},
				this.ReadCsv(outputByHeaders),
				"select by headers"
				);
		}

		/// <summary>
		/// csv-select-columns が必須オプション不足、排他オプション同時指定、範囲外列、存在しないヘッダーをエラーにすることを確認する。
		/// </summary>
		private void TestCsvSelectColumnsRejectsBadOptions()
		{
			this.AssertThrows(() => this.Run(new string[] { "csv-select-columns", this.InputFile("basic.csv"), this.OutputFile("select-no-option.csv") }), "Specify either --columns or --headers.");
			this.AssertThrows(() => this.Run(new string[] { "csv-select-columns", "--columns", "1", "--headers", "ID", this.InputFile("basic.csv"), this.OutputFile("select-both-options.csv") }), "Specify either --columns or --headers.");
			this.AssertThrows(() => this.Run(new string[] { "csv-select-columns", "--columns", "0", this.InputFile("basic.csv"), this.OutputFile("select-zero-column.csv") }), "Bad --columns");
			this.AssertThrows(() => this.Run(new string[] { "csv-select-columns", "--columns", "99", this.InputFile("basic.csv"), this.OutputFile("select-out-of-range.csv") }), "Column index out of range");
			this.AssertThrows(() => this.Run(new string[] { "csv-select-columns", "--headers", "Missing", this.InputFile("basic.csv"), this.OutputFile("select-missing-header.csv") }), "Header not found");
		}

		/// <summary>
		/// csv-filter-rows が equals、contains、regex、invert、ヘッダー行維持の条件で期待行だけを抽出することを確認する。
		/// </summary>
		private void TestCsvFilterRowsByEqualsContainsRegexAndInvert()
		{
			string equalsOutput = this.OutputFile("filter-equals.csv");

			this.Run(new string[] { "csv-filter-rows", "--header", "Category", "--equals", "Fruit", this.InputFile("basic.csv"), equalsOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "4", "Apple", "200", "Fruit" },
				},
				this.ReadCsv(equalsOutput),
				"filter equals"
				);

			string containsOutput = this.OutputFile("filter-contains.csv");

			this.Run(new string[] { "csv-filter-rows", "--column", "2", "--contains", "pp", this.InputFile("basic.csv"), containsOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "4", "Apple", "200", "Fruit" },
				},
				this.ReadCsv(containsOutput),
				"filter contains"
				);

			string regexOutput = this.OutputFile("filter-regex.csv");

			this.Run(new string[] { "csv-filter-rows", "--column", "3", "--regex", "^1[0-9]{2}$", "--has-header", this.InputFile("basic.csv"), regexOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "3", "Banana", "150", "Fruit" },
				},
				this.ReadCsv(regexOutput),
				"filter regex"
				);

			string invertOutput = this.OutputFile("filter-invert.csv");

			this.Run(new string[] { "csv-filter-rows", "--header", "Category", "--equals", "Fruit", "--invert", "--has-header", this.InputFile("basic.csv"), invertOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
				},
				this.ReadCsv(invertOutput),
				"filter invert"
				);
		}

		/// <summary>
		/// csv-filter-rows が列指定不足、条件不足、条件の複数指定、CSV 系への engine 指定をエラーにすることを確認する。
		/// </summary>
		private void TestCsvFilterRowsRejectsBadOptions()
		{
			this.AssertThrows(() => this.Run(new string[] { "csv-filter-rows", "--equals", "Fruit", this.InputFile("basic.csv"), this.OutputFile("filter-no-column.csv") }), "Specify either --column or --header.");
			this.AssertThrows(() => this.Run(new string[] { "csv-filter-rows", "--column", "1", this.InputFile("basic.csv"), this.OutputFile("filter-no-condition.csv") }), "Specify one of --equals, --contains, or --regex.");
			this.AssertThrows(() => this.Run(new string[] { "csv-filter-rows", "--column", "1", "--equals", "1", "--contains", "1", this.InputFile("basic.csv"), this.OutputFile("filter-multiple-conditions.csv") }), "Specify one of --equals, --contains, or --regex.");
			this.AssertThrows(() => this.Run(new string[] { "csv-filter-rows", "--engine", "zip", "--column", "1", "--equals", "1", this.InputFile("basic.csv"), this.OutputFile("filter-engine.csv") }), "--engine cannot be used");
		}

		/// <summary>
		/// csv-replace が全セル置換、列番号指定置換、ヘッダー名指定の正規表現置換を正しく行うことを確認する。
		/// </summary>
		private void TestCsvReplaceAllCellsAndSelectedColumns()
		{
			string allOutput = this.OutputFile("replace-all.csv");

			this.Run(new string[] { "csv-replace", "--from", "Apple", "--to", "Orange", this.InputFile("basic.csv"), allOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Orange", "100", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "4", "Orange", "200", "Fruit" },
				},
				this.ReadCsv(allOutput),
				"replace all cells"
				);

			string columnOutput = this.OutputFile("replace-column.csv");

			this.Run(new string[] { "csv-replace", "--from", "Fruit", "--to", "Food", "--column", "4", this.InputFile("basic.csv"), columnOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Food" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Food" },
					new string[] { "4", "Apple", "200", "Food" },
				},
				this.ReadCsv(columnOutput),
				"replace selected column"
				);

			string regexOutput = this.OutputFile("replace-regex-header.csv");

			this.Run(new string[] { "csv-replace", "--regex", "^[0-9]+$", "--to", "NUMBER", "--header", "Price", this.InputFile("basic.csv"), regexOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "NUMBER", "Fruit" },
					new string[] { "2", "Carrot", "NUMBER", "Vegetable" },
					new string[] { "3", "Banana", "NUMBER", "Fruit" },
					new string[] { "4", "Apple", "NUMBER", "Fruit" },
				},
				this.ReadCsv(regexOutput),
				"replace regex by header"
				);
		}

		/// <summary>
		/// csv-replace が置換先不足、置換元指定不足、通常置換と正規表現置換の同時指定をエラーにすることを確認する。
		/// </summary>
		private void TestCsvReplaceRejectsBadOptions()
		{
			this.AssertThrows(() => this.Run(new string[] { "csv-replace", "--from", "Apple", this.InputFile("basic.csv"), this.OutputFile("replace-no-to.csv") }), "Missing --to.");
			this.AssertThrows(() => this.Run(new string[] { "csv-replace", "--to", "Orange", this.InputFile("basic.csv"), this.OutputFile("replace-no-source.csv") }), "Specify either --from or --regex.");
			this.AssertThrows(() => this.Run(new string[] { "csv-replace", "--from", "Apple", "--regex", "Apple", "--to", "Orange", this.InputFile("basic.csv"), this.OutputFile("replace-both-source.csv") }), "Specify either --from or --regex.");
		}

		/// <summary>
		/// csv-merge がファイル名順に結合し、skip-header 指定時に 2 ファイル目以降のヘッダーを除外することを確認する。
		/// </summary>
		private void TestCsvMergeOrdersFilesAndSkipsRepeatedHeaders()
		{
			string output = this.OutputFile("merge.csv");

			this.Run(new string[] { "csv-merge", "--skip-header", this.InputFile("merge"), output });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "1", "Alpha" },
					new string[] { "2", "Beta" },
					new string[] { "3", "Gamma" },
					new string[] { "4", "Delta" },
				},
				this.ReadCsv(output),
				"merge skip header"
				);

			string patternOutput = this.OutputFile("merge-pattern.csv");

			this.Run(new string[] { "csv-merge", "--pattern", "*.part", this.InputFile("merge"), patternOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "9", "PartOnly" },
				},
				this.ReadCsv(patternOutput),
				"merge pattern"
				);

			this.AssertThrows(() => this.Run(new string[] { "csv-merge", this.InputFile("missing-dir"), this.OutputFile("merge-missing.csv") }), "Input directory not found");
		}

		/// <summary>
		/// csv-sort が文字列昇順、ヘッダー指定の数値昇順、数値降順で期待順に並べ替えることを確認する。
		/// </summary>
		private void TestCsvSortStringAndNumericValues()
		{
			string stringOutput = this.OutputFile("sort-string.csv");

			this.Run(new string[] { "csv-sort", "--column", "2", "--has-header", this.InputFile("basic.csv"), stringOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "4", "Apple", "200", "Fruit" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
				},
				this.ReadCsv(stringOutput),
				"sort string"
				);

			string numericOutput = this.OutputFile("sort-numeric.csv");

			this.Run(new string[] { "csv-sort", "--header", "Price", "--numeric", this.InputFile("basic.csv"), numericOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "4", "Apple", "200", "Fruit" },
				},
				this.ReadCsv(numericOutput),
				"sort numeric"
				);

			string descOutput = this.OutputFile("sort-desc.csv");

			this.Run(new string[] { "csv-sort", "--header", "Price", "--numeric", "--desc", this.InputFile("basic.csv"), descOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "4", "Apple", "200", "Fruit" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "1", "Apple", "100", "Fruit" },
				},
				this.ReadCsv(descOutput),
				"sort desc"
				);
		}

		/// <summary>
		/// csv-unique が行全体、列番号キー、ヘッダー名キーで最初に出現した行だけを残すことを確認する。
		/// </summary>
		private void TestCsvUniqueRemovesDuplicateRowsAndKeys()
		{
			string allOutput = this.OutputFile("unique-all.csv");

			this.Run(new string[] { "csv-unique", this.InputFile("duplicates.csv"), allOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "4", "Apple", "200", "Fruit" },
				},
				this.ReadCsv(allOutput),
				"unique all columns"
				);

			string columnsOutput = this.OutputFile("unique-columns.csv");

			this.Run(new string[] { "csv-unique", "--columns", "2,4", this.InputFile("duplicates.csv"), columnsOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Fruit" },
				},
				this.ReadCsv(columnsOutput),
				"unique by columns"
				);

			string headersOutput = this.OutputFile("unique-headers.csv");

			this.Run(new string[] { "csv-unique", "--headers", "Name,Category", this.InputFile("duplicates.csv"), headersOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Fruit" },
				},
				this.ReadCsv(headersOutput),
				"unique by headers"
				);

			this.AssertThrows(() => this.Run(new string[] { "csv-unique", "--columns", "1", "--headers", "ID", this.InputFile("duplicates.csv"), this.OutputFile("unique-both-options.csv") }), "--columns and --headers cannot be used together.");
		}

		/// <summary>
		/// delimiter と encoding の指定により TSV/SSV/UTF-8 BOM 入出力が正しく扱われることを確認する。
		/// </summary>
		private void TestDelimiterAndEncodingOptions()
		{
			string tsvOutput = this.OutputFile("select-tsv.tsv");

			this.Run(new string[] { "csv-select-columns", "--columns", "1,3", this.InputFile("tab.tsv"), tsvOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "Code", "Memo" },
					new string[] { "A", "Hello" },
					new string[] { "B", "Tabbed" },
				},
				CsvFileReader.ReadToEnd(tsvOutput, this.Sjis, '\t'),
				"tsv default delimiter"
				);

			string ssvOutput = this.OutputFile("select-ssv.csv");

			this.Run(new string[] { "csv-select-columns", "--delimiter", "space", "--columns", "1,3", this.InputFile("space.ssv"), ssvOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "Code", "Memo" },
					new string[] { "A", "Hello" },
					new string[] { "B", "World" },
				},
				CsvFileReader.ReadToEnd(ssvOutput, this.Sjis, ' '),
				"space delimiter"
				);

			string utf8BomOutput = this.OutputFile("utf8bom.csv");

			this.Run(new string[] { "csv-select-columns", "--encoding", "utf8bom", "--columns", "1,2", this.InputFile("utf8bom.csv"), utf8BomOutput });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "1", "桃" },
					new string[] { "2", "梨" },
				},
				CsvFileReader.ReadToEnd(utf8BomOutput, Encoding.UTF8, ','),
				"utf8bom output"
				);

			byte[] bytes = File.ReadAllBytes(utf8BomOutput);

			this.AssertTrue(3 <= bytes.Length && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf, "utf8bom output has BOM");
			this.AssertThrows(() => this.Run(new string[] { "csv-info", "--delimiter", "bad", this.InputFile("basic.csv") }), "Bad delimiter");
			this.AssertThrows(() => this.Run(new string[] { "csv-info", "--encoding", "unknown", this.InputFile("basic.csv") }), "Unknown encoding");
		}

		/// <summary>
		/// 出力先が存在する場合は --overwrite なしで失敗し、--overwrite ありで置き換えられることを確認する。
		/// </summary>
		private void TestOverwriteProtection()
		{
			string output = this.OutputFile("overwrite.csv");

			File.WriteAllText(output, "already exists", this.Sjis);
			this.AssertThrows(() => this.Run(new string[] { "csv-select-columns", "--columns", "1", this.InputFile("basic.csv"), output }), "Output path already exists");

			this.Run(new string[] { "csv-select-columns", "--overwrite", "--columns", "1", this.InputFile("basic.csv"), output });
			this.AssertRows(
				new string[][]
				{
					new string[] { "ID" },
					new string[] { "1" },
					new string[] { "2" },
					new string[] { "3" },
					new string[] { "4" },
				},
				this.ReadCsv(output),
				"overwrite output"
				);

			this.AssertThrows(() => this.Run(new string[] { "csv-info", this.InputFile("missing.csv") }), "Input file not found");
			this.AssertThrows(() => this.Run(new string[] { "csv-info", this.InputFile("basic.csv"), this.InputFile("extra.csv") }), "Bad argument count");
		}

		/// <summary>
		/// Excel 文字列置換とプレースホルダ置換の正常系・異常系を確認する。
		/// </summary>
		private void TestExcelReplaceTextAndPlaceholder()
		{
			string replaceOutput = this.OutputFile("excel-replace-text.xlsx");
			string regexOutput = this.OutputFile("excel-replace-regex.xlsx");
			string sheetOutput = this.OutputFile("excel-replace-sheet.xlsx");
			string placeholderOutput = this.OutputFile("excel-replace-placeholder.xlsx");
			string setFileOutput = this.OutputFile("excel-replace-placeholder-set-file.xlsx");
			string existingOutput = this.OutputFile("excel-replace-existing.xlsx");

			this.Run(new string[] { "excel-replace-text", "--from", "Apple", "--to", "Orange", this.InputFile("text-replace.xlsx"), replaceOutput });
			this.AssertWorkbookContains(replaceOutput, "Orange", "excel replace text");

			this.Run(new string[] { "excel-replace-text", "--regex", "Item[0-9]+", "--to", "ItemX", this.InputFile("text-replace.xlsx"), regexOutput });
			this.AssertWorkbookContains(regexOutput, "ItemX", "excel replace regex");

			this.Run(new string[] { "excel-replace-text", "--sheet", "Target", "--from", "SheetOnly", "--to", "ChangedSheet", this.InputFile("text-replace.xlsx"), sheetOutput });
			this.AssertWorkbookContains(sheetOutput, "ChangedSheet", "excel replace text sheet");
			this.AssertWorkbookContains(sheetOutput, "SheetOnly", "excel replace text other sheet");

			this.Run(new string[] { "excel-replace-placeholder", "--set", "**NAME**=Yamada", "--set", "**CITY**=Tokyo", this.InputFile("text-replace.xlsx"), placeholderOutput });
			this.AssertWorkbookContains(placeholderOutput, "Yamada", "excel replace placeholder name");
			this.AssertWorkbookContains(placeholderOutput, "Tokyo", "excel replace placeholder city");

			this.Run(new string[] { "excel-replace-placeholder", "--set-file", this.InputFile("placeholder-map.csv"), this.InputFile("text-replace.xlsx"), setFileOutput });
			this.AssertWorkbookContains(setFileOutput, "Suzuki", "excel replace placeholder set file name");
			this.AssertWorkbookContains(setFileOutput, "Osaka", "excel replace placeholder set file city");

			File.WriteAllText(existingOutput, "already exists", this.Sjis);

			this.AssertThrows(() => this.Run(new string[] { "excel-replace-text", "--from", "Apple", "--regex", "Apple", "--to", "Orange", this.InputFile("text-replace.xlsx"), this.OutputFile("excel-replace-both.xlsx") }), "Specify either --from or --regex.");
			this.AssertThrows(() => this.Run(new string[] { "excel-replace-text", "--from", "Apple", this.InputFile("text-replace.xlsx"), this.OutputFile("excel-replace-no-to.xlsx") }), "--to is required.");
			this.AssertThrows(() => this.Run(new string[] { "excel-replace-placeholder", this.InputFile("text-replace.xlsx"), this.OutputFile("excel-replace-no-set.xlsx") }), "Specify --set or --set-file.");
			this.AssertThrows(() => this.Run(new string[] { "excel-replace-text", "--from", "Apple", "--to", "Orange", this.InputFile("text-replace.xlsx"), existingOutput }), "Output path already exists");
		}

		/// <summary>
		/// excel-extract-pictures が xlsx ZIP 内の画像だけを抽出し、連番ファイルとして保存することを確認する。
		/// </summary>
		private void TestExcelExtractPicturesByZipEngine()
		{
			string outputDir = this.OutputFile("pictures");

			this.Run(new string[] { "excel-extract-pictures", "--engine", "zip", this.InputFile("pictures.xlsx"), outputDir });
			this.AssertTrue(File.Exists(Path.Combine(outputDir, "0001.png")), "first extracted picture exists");
			this.AssertTrue(File.Exists(Path.Combine(outputDir, "0002.jpg")), "second extracted picture exists");
			this.AssertBytes(new byte[] { 1, 2, 3, 4 }, File.ReadAllBytes(Path.Combine(outputDir, "0001.png")), "first extracted picture");
			this.AssertBytes(new byte[] { 5, 6, 7, 8 }, File.ReadAllBytes(Path.Combine(outputDir, "0002.jpg")), "second extracted picture");
			this.AssertEquals(2, Directory.GetFiles(outputDir).Length, "extracted picture count");
		}

		/// <summary>
		/// excel-replace-picture が xlsx ZIP 内の全画像置換と index 指定置換を行うことを確認する。
		/// </summary>
		private void TestExcelReplacePictureByZipEngine()
		{
			string allOutput = this.OutputFile("replace-picture-all.xlsx");
			string indexedOutput = this.OutputFile("replace-picture-index.xlsx");

			this.Run(new string[] { "excel-replace-picture", "--engine", "zip", this.InputFile("pictures.xlsx"), allOutput, this.InputFile("replacement.png") });
			this.AssertZipEntryBytes(new byte[] { 9, 8, 7, 6 }, allOutput, "xl/media/image1.png", "all replace image1");
			this.AssertZipEntryBytes(new byte[] { 9, 8, 7, 6 }, allOutput, "xl/media/image2.jpg", "all replace image2");

			this.Run(new string[] { "excel-replace-picture", "--engine", "zip", "--index", "2", this.InputFile("pictures.xlsx"), indexedOutput, this.InputFile("replacement.png") });
			this.AssertZipEntryBytes(new byte[] { 1, 2, 3, 4 }, indexedOutput, "xl/media/image1.png", "index replace image1");
			this.AssertZipEntryBytes(new byte[] { 9, 8, 7, 6 }, indexedOutput, "xl/media/image2.jpg", "index replace image2");
		}

		/// <summary>
		/// ZIP 直操作の Excel 画像コマンドが不正な engine、index、入力画像不足、出力先存在をエラーにすることを確認する。
		/// </summary>
		private void TestZipEngineCommandsRejectBadOptions()
		{
			string outputDir = this.OutputFile("pictures-existing");
			string outputXlsx = this.OutputFile("replace-picture-existing.xlsx");

			Directory.CreateDirectory(outputDir);
			File.WriteAllText(outputXlsx, "already exists", this.Sjis);

			this.AssertThrows(() => this.Run(new string[] { "excel-extract-pictures", "--engine", "app", this.InputFile("pictures.xlsx"), this.OutputFile("pictures-bad-engine") }), "Unsupported engine");
			this.AssertThrows(() => this.Run(new string[] { "excel-extract-pictures", this.InputFile("pictures.xlsx"), outputDir }), "Output path already exists");
			this.AssertThrows(() => this.Run(new string[] { "excel-replace-picture", "--index", "0", this.InputFile("pictures.xlsx"), this.OutputFile("replace-picture-bad-index.xlsx"), this.InputFile("replacement.png") }), "Bad --index");
			this.AssertThrows(() => this.Run(new string[] { "excel-replace-picture", this.InputFile("pictures.xlsx"), this.OutputFile("replace-picture-missing-image.xlsx"), this.InputFile("missing.png") }), "Picture file not found");
			this.AssertThrows(() => this.Run(new string[] { "excel-replace-picture", this.InputFile("pictures.xlsx"), outputXlsx, this.InputFile("replacement.png") }), "Output path already exists");
		}

		private void PrepareTestData()
		{
			this.DeleteIfExists(ResourceDir);
			Directory.CreateDirectory(InputDir);
			Directory.CreateDirectory(OutputDir);
			Directory.CreateDirectory(this.InputFile("merge"));

			this.WriteCsv(
				this.InputFile("basic.csv"),
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "4", "Apple", "200", "Fruit" },
				},
				this.Sjis,
				','
				);

			this.WriteCsv(
				this.InputFile("duplicates.csv"),
				new string[][]
				{
					new string[] { "ID", "Name", "Price", "Category" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "1", "Apple", "100", "Fruit" },
					new string[] { "2", "Carrot", "200", "Vegetable" },
					new string[] { "3", "Banana", "150", "Fruit" },
					new string[] { "4", "Apple", "200", "Fruit" },
				},
				this.Sjis,
				','
				);

			this.WriteCsv(
				this.InputFile("info.csv"),
				new string[][]
				{
					new string[] { "A", "B", "C" },
					new string[] { "1", "2", "3", "4" },
					new string[] { "" },
					new string[] { "x", "y" },
				},
				this.Sjis,
				','
				);

			this.WriteCsv(
				this.InputFile("tab.tsv"),
				new string[][]
				{
					new string[] { "Code", "Name", "Memo" },
					new string[] { "A", "Alpha", "Hello" },
					new string[] { "B", "Beta", "Tabbed" },
				},
				this.Sjis,
				'\t'
				);

			this.WriteCsv(
				this.InputFile("space.ssv"),
				new string[][]
				{
					new string[] { "Code", "Name", "Memo" },
					new string[] { "A", "Alpha", "Hello" },
					new string[] { "B", "Beta", "World" },
				},
				this.Sjis,
				' '
				);

			this.WriteCsv(
				this.InputFile("utf8bom.csv"),
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "1", "桃" },
					new string[] { "2", "梨" },
				},
				this.Utf8Bom,
				','
				);

			this.WriteCsv(
				this.InputFile(@"merge\001.csv"),
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "1", "Alpha" },
					new string[] { "2", "Beta" },
				},
				this.Sjis,
				','
				);

			this.WriteCsv(
				this.InputFile(@"merge\002.csv"),
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "3", "Gamma" },
					new string[] { "4", "Delta" },
				},
				this.Sjis,
				','
				);

			this.WriteCsv(
				this.InputFile(@"merge\ignored.tsv"),
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "8", "Ignored" },
				},
				this.Sjis,
				'\t'
				);

			this.WriteCsv(
				this.InputFile(@"merge\only.part"),
				new string[][]
				{
					new string[] { "ID", "Name" },
					new string[] { "9", "PartOnly" },
				},
				this.Sjis,
				','
				);

			File.WriteAllBytes(this.InputFile("replacement.png"), new byte[] { 9, 8, 7, 6 });
			this.CreatePictureWorkbook(this.InputFile("pictures.xlsx"));
			this.CreateTextWorkbook(this.InputFile("text-replace.xlsx"));
			this.WriteCsv(
				this.InputFile("placeholder-map.csv"),
				new string[][]
				{
					new string[] { "**NAME**", "Suzuki" },
					new string[] { "**CITY**", "Osaka" },
				},
				this.Sjis,
				','
				);
		}

		private void Run(string[] args)
		{
			ECWeaverArgs parsedArgs = ECWeaverArgs.Read(new ArgsReader(args));
			new ECWeaverProcessor().Run(parsedArgs);
		}

		private string CaptureOutput(string[] args)
		{
			TextWriter originalWriter = Console.Out;
			StringWriter writer = new StringWriter();

			try
			{
				Console.SetOut(writer);
				this.Run(args);
			}
			finally
			{
				Console.SetOut(originalWriter);
			}
			return writer.ToString();
		}

		private string InputFile(string localPath)
		{
			return Path.Combine(InputDir, localPath);
		}

		private string OutputFile(string localPath)
		{
			return Path.Combine(OutputDir, localPath);
		}

		private string[][] ReadCsv(string file)
		{
			return CsvFileReader.ReadToEnd(file, this.Sjis, ',');
		}

		private void WriteCsv(string file, string[][] rows, Encoding encoding, char delimiter)
		{
			string dir = Path.GetDirectoryName(file);

			if (!Directory.Exists(dir))
				Directory.CreateDirectory(dir);

			CsvFileWriter.WriteRows(file, false, encoding, delimiter, rows);
		}

		private void CreatePictureWorkbook(string file)
		{
			string dir = Path.GetDirectoryName(file);

			if (!Directory.Exists(dir))
				Directory.CreateDirectory(dir);

			if (File.Exists(file))
				File.Delete(file);

			using (ZipArchive archive = ZipFile.Open(file, ZipArchiveMode.Create))
			{
				this.WriteZipEntry(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />");
				this.WriteZipEntry(archive, "xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
				this.WriteZipEntry(archive, "xl/media/image1.png", new byte[] { 1, 2, 3, 4 });
				this.WriteZipEntry(archive, "xl/media/image2.jpg", new byte[] { 5, 6, 7, 8 });
				this.WriteZipEntry(archive, "xl/media/readme.txt", new byte[] { 0 });
			}
		}

		private void CreateTextWorkbook(string file)
		{
			string dir = Path.GetDirectoryName(file);

			if (!Directory.Exists(dir))
				Directory.CreateDirectory(dir);

			if (File.Exists(file))
				File.Delete(file);

			using (ZipArchive archive = ZipFile.Open(file, ZipArchiveMode.Create))
			{
				this.WriteZipEntry(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
				this.WriteZipEntry(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
				this.WriteZipEntry(archive, "xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/></Relationships>");
				this.WriteZipEntry(archive, "xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Main\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Target\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>");
				this.WriteZipEntry(archive, "xl/worksheets/sheet1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Apple</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Item123</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>**NAME**</t></is></c><c r=\"D1\" t=\"inlineStr\"><is><t>SheetOnly</t></is></c></row></sheetData></worksheet>");
				this.WriteZipEntry(archive, "xl/worksheets/sheet2.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Apple</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>**CITY**</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>SheetOnly</t></is></c></row></sheetData></worksheet>");
			}
		}

		private void AssertWorkbookContains(string file, string expectedPart, string message)
		{
			using (ZipArchive archive = ZipFile.OpenRead(file))
			{
				string text = string.Join("\n", archive.Entries
					.Where(entry => entry.FullName.StartsWith("xl/") && entry.FullName.EndsWith(".xml"))
					.Select(entry => this.ReadZipEntryText(entry))
					.ToArray());

				this.AssertContains(text, expectedPart, message);
			}
		}

		private string ReadZipEntryText(ZipArchiveEntry entry)
		{
			using (Stream stream = entry.Open())
			using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
			{
				return reader.ReadToEnd();
			}
		}

		private void WriteZipEntry(ZipArchive archive, string entryName, string text)
		{
			this.WriteZipEntry(archive, entryName, this.Utf8NoBom.GetBytes(text));
		}

		private void WriteZipEntry(ZipArchive archive, string entryName, byte[] bytes)
		{
			ZipArchiveEntry entry = archive.CreateEntry(entryName);

			using (Stream stream = entry.Open())
			{
				stream.Write(bytes, 0, bytes.Length);
			}
		}

		private void AssertZipEntryBytes(byte[] expected, string zipFile, string entryName, string message)
		{
			using (ZipArchive archive = ZipFile.OpenRead(zipFile))
			{
				ZipArchiveEntry entry = archive.GetEntry(entryName);

				if (entry == null)
					throw new Exception("ZIP entry not found: " + entryName);

				using (Stream stream = entry.Open())
				using (MemoryStream memory = new MemoryStream())
				{
					stream.CopyTo(memory);
					this.AssertBytes(expected, memory.ToArray(), message);
				}
			}
		}

		private void AssertRows(string[][] expected, string[][] actual, string message)
		{
			this.AssertEquals(expected.Length, actual.Length, message + " row count");

			for (int row = 0; row < expected.Length; row++)
			{
				this.AssertEquals(expected[row].Length, actual[row].Length, message + " column count at row " + row);

				for (int column = 0; column < expected[row].Length; column++)
					this.AssertEquals(expected[row][column], actual[row][column], message + " cell at row " + row + ", column " + column);
			}
		}

		private void AssertBytes(byte[] expected, byte[] actual, string message)
		{
			this.AssertEquals(expected.Length, actual.Length, message + " byte length");

			for (int index = 0; index < expected.Length; index++)
				this.AssertEquals(expected[index], actual[index], message + " byte at index " + index);
		}

		private void AssertThrows(Action routine, string expectedMessagePart)
		{
			try
			{
				routine();
			}
			catch (Exception ex)
			{
				if (ex.Message.Contains(expectedMessagePart))
					return;

				throw new Exception("Unexpected exception message. Expected part: " + expectedMessagePart + ", Actual: " + ex.Message, ex);
			}
			throw new Exception("Expected exception was not thrown. Expected part: " + expectedMessagePart);
		}

		private void AssertContains(string actual, string expectedPart, string message)
		{
			if (!actual.Contains(expectedPart))
				throw new Exception(message + " does not contain expected text. Expected part: " + expectedPart + ", Actual: " + actual);
		}

		private void AssertEquals<T>(T expected, T actual, string message)
		{
			if (!object.Equals(expected, actual))
				throw new Exception(message + " mismatch. Expected: " + expected + ", Actual: " + actual);
		}

		private void AssertTrue(bool condition, string message)
		{
			if (!condition)
				throw new Exception(message);
		}

		private void DeleteIfExists(string path)
		{
			if (Directory.Exists(path))
				Directory.Delete(path, true);
			else if (File.Exists(path))
				File.Delete(path);
		}

		private void ShowSuccessBanner()
		{
			Console.WriteLine("");
			Console.WriteLine("============================================================");
			Console.WriteLine("============================================================");
			Console.WriteLine("====                                                    ====");
			Console.WriteLine("====              ECWeaver Test0001 SUCCESS             ====");
			Console.WriteLine("====                                                    ====");
			Console.WriteLine("====              ALL TESTS PASSED                      ====");
			Console.WriteLine("====                                                    ====");
			Console.WriteLine("============================================================");
			Console.WriteLine("============================================================");
			Console.WriteLine("");
		}

		private void ShowFailure(Exception ex)
		{
			Console.WriteLine("");
			Console.WriteLine("============================================================");
			Console.WriteLine("ECWeaver Test0001 FAILED");
			Console.WriteLine("============================================================");
			Console.WriteLine(ex.ToString());
			Console.WriteLine("============================================================");
			Console.WriteLine("");
		}
#endif
	}
}
