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
		private const string ResourceDir = @"C:\home\res\ExcelCsvWeaver\ECWeaver2\Test0001";
		private const string BaseDir = @"C:\temp\ECWeaver2";
		private const string InputDir = ResourceDir + @"\input";
		private const string OutputDir = BaseDir + @"\output";

		private readonly Encoding Sjis = Encoding.GetEncoding(932);
		private readonly Encoding Utf8Bom = new UTF8Encoding(true);

		public void Test01()
		{
			SCommon.Pause_WaitSeconds = 0;
			SCommon.DeleteAndCreateDir(BaseDir);

			this.PrepareTestData();

			this.TestHelpVersionAndCommandErrors();
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
			this.TestCsvToExcel();
			this.TestCsvsToExcel();
			this.TestWeaveToExcel();
			this.TestExcelOutputCommandsRejectBadOptions();

			this.ShowSuccessBanner();
		}

		/// <summary>
		/// help/version の標準出力と、未知コマンド・ECWeaver2 未実装コマンドの異常系を確認する。
		/// </summary>
		private void TestHelpVersionAndCommandErrors()
		{
			string helpOutput = this.CaptureOutput(new string[] { "help" });
			this.AssertContains(helpOutput, "Usage:", "help output");
			this.AssertContains(helpOutput, "csv-to-excel", "help output");
			this.AssertContains(helpOutput, "csv-info", "help output");

			string commandHelpOutput = this.CaptureOutput(new string[] { "help", "csv-replace" });
			this.AssertContains(commandHelpOutput, "Usage: ECWeaver2.exe csv-replace", "command help output");

			string versionOutput = this.CaptureOutput(new string[] { "version" });
			this.AssertContains(versionOutput, "ECWeaver2 ", "version output");

			this.AssertThrows(() => this.Run(new string[] { "not-a-command" }), "Unknown command");
			this.AssertThrows(() => this.Run(new string[] { "excel-to-csv", this.InputFile("basic.csv"), this.OutputFile("not-implemented.csv") }), "not implemented in ECWeaver2");
			this.AssertThrows(() => this.Run(new string[] { "excel-extract-pictures", this.InputFile("basic.csv"), this.OutputFile("pictures") }), "not implemented in ECWeaver2");
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

			string ssvOutput = this.OutputFile("select-ssv.ssv");

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
		/// csv-to-excel が CSV から XLSX を生成し、sheet 指定が workbook.xml に反映されることを確認する。
		/// </summary>
		private void TestCsvToExcel()
		{
			string output = this.OutputFile("csv-to-excel.xlsx");

			this.Run(new string[] { "csv-to-excel", "--sheet", "CustomSheet", this.InputFile("basic.csv"), output });
			this.AssertWorkbookBasics(output, 1, "csv-to-excel");
			this.AssertZipEntryContains(output, "xl/workbook.xml", "CustomSheet", "csv-to-excel sheet name");
			this.AssertWorkbookContains(output, "Apple", "csv-to-excel string");
			this.AssertWorkbookContains(output, "Vegetable", "csv-to-excel string");
		}

		/// <summary>
		/// csvs-to-excel がディレクトリ内 CSV をファイル名順の複数シートとして XLSX に出力することを確認する。
		/// </summary>
		private void TestCsvsToExcel()
		{
			string output = this.OutputFile("csvs-to-excel.xlsx");
			string patternOutput = this.OutputFile("csvs-to-excel-pattern.xlsx");

			this.Run(new string[] { "csvs-to-excel", this.InputFile("excel-sheets"), output });
			this.AssertWorkbookBasics(output, 2, "csvs-to-excel");
			this.AssertZipEntryContains(output, "xl/workbook.xml", "alpha", "csvs-to-excel alpha sheet");
			this.AssertZipEntryContains(output, "xl/workbook.xml", "beta", "csvs-to-excel beta sheet");
			this.AssertWorkbookContains(output, "AlphaName", "csvs-to-excel alpha data");
			this.AssertWorkbookContains(output, "BetaName", "csvs-to-excel beta data");

			this.Run(new string[] { "csvs-to-excel", "--pattern", "*.part", this.InputFile("excel-sheets"), patternOutput });
			this.AssertWorkbookBasics(patternOutput, 1, "csvs-to-excel pattern");
			this.AssertZipEntryContains(patternOutput, "xl/workbook.xml", "only", "csvs-to-excel pattern sheet");
			this.AssertWorkbookContains(patternOutput, "PartName", "csvs-to-excel pattern data");
		}

		/// <summary>
		/// weave --to-excel が CSV / TSV / SSV 混在入力を複数シートの XLSX に出力することを確認する。
		/// </summary>
		private void TestWeaveToExcel()
		{
			string output = this.OutputFile("weave.xlsx");

			this.Run(new string[] { "weave", this.InputFile("basic.csv"), this.InputFile("tab.tsv"), this.InputFile("space.ssv"), "--to-excel", output });
			this.AssertWorkbookBasics(output, 3, "weave to excel");
			this.AssertZipEntryContains(output, "xl/workbook.xml", "basic", "weave basic sheet");
			this.AssertZipEntryContains(output, "xl/workbook.xml", "tab", "weave tab sheet");
			this.AssertZipEntryContains(output, "xl/workbook.xml", "space", "weave space sheet");
			this.AssertWorkbookContains(output, "Apple", "weave csv data");
			this.AssertWorkbookContains(output, "Tabbed", "weave tsv data");
			this.AssertWorkbookContains(output, "World", "weave ssv data");
		}

		/// <summary>
		/// ECWeaver2 固有の Excel 出力系コマンドが出力先存在、入力不足、不正 engine、不正 weave 指定をエラーにすることを確認する。
		/// </summary>
		private void TestExcelOutputCommandsRejectBadOptions()
		{
			string csvToExcelOutput = this.OutputFile("csv-to-excel-existing.xlsx");
			string weaveOutput = this.OutputFile("weave-existing.xlsx");

			File.WriteAllText(csvToExcelOutput, "already exists", this.Sjis);
			File.WriteAllText(weaveOutput, "already exists", this.Sjis);

			this.AssertThrows(() => this.Run(new string[] { "csv-to-excel", "--engine", "zip", this.InputFile("basic.csv"), this.OutputFile("csv-to-excel-bad-engine.xlsx") }), "Unsupported engine");
			this.AssertThrows(() => this.Run(new string[] { "csv-to-excel", this.InputFile("basic.csv"), csvToExcelOutput }), "Output path already exists");
			this.AssertThrows(() => this.Run(new string[] { "csvs-to-excel", this.InputFile("missing-dir"), this.OutputFile("csvs-missing.xlsx") }), "Input directory not found");
			this.AssertThrows(() => this.Run(new string[] { "csvs-to-excel", "--pattern", "*.missing", this.InputFile("excel-sheets"), this.OutputFile("csvs-empty.xlsx") }), "Input CSV file not found");
			this.AssertThrows(() => this.Run(new string[] { "weave", this.InputFile("basic.csv") }), "Specify one output mode");
			this.AssertThrows(() => this.Run(new string[] { "weave", this.InputFile("basic.csv"), "--to-excel", this.OutputFile("weave-both.xlsx"), "--to-csv-dir", this.OutputFile("csv-dir") }), "Specify one output mode");
			this.AssertThrows(() => this.Run(new string[] { "weave", "--to-excel", this.OutputFile("weave-no-input.xlsx") }), "Input file not specified");
			this.AssertThrows(() => this.Run(new string[] { "weave", this.InputFile("unsupported.txt"), "--to-excel", this.OutputFile("weave-unsupported.xlsx") }), "currently supports CSV/TSV/SSV");
			this.AssertThrows(() => this.Run(new string[] { "weave", this.InputFile("basic.csv"), "--to-excel", weaveOutput }), "Output path already exists");
			this.AssertThrows(() => this.Run(new string[] { "weave", this.InputFile("basic.csv"), "--to-same-dir" }), "not implemented in ECWeaver2");
		}

		private void PrepareTestData()
		{
			this.DeleteIfExists(ResourceDir);
			Directory.CreateDirectory(InputDir);
			Directory.CreateDirectory(OutputDir);
			Directory.CreateDirectory(this.InputFile("merge"));
			Directory.CreateDirectory(this.InputFile("excel-sheets"));

			this.WriteCsv(this.InputFile("basic.csv"), new string[][]
			{
				new string[] { "ID", "Name", "Price", "Category" },
				new string[] { "1", "Apple", "100", "Fruit" },
				new string[] { "2", "Carrot", "200", "Vegetable" },
				new string[] { "3", "Banana", "150", "Fruit" },
				new string[] { "4", "Apple", "200", "Fruit" },
			}, this.Sjis, ',');

			this.WriteCsv(this.InputFile("duplicates.csv"), new string[][]
			{
				new string[] { "ID", "Name", "Price", "Category" },
				new string[] { "1", "Apple", "100", "Fruit" },
				new string[] { "1", "Apple", "100", "Fruit" },
				new string[] { "2", "Carrot", "200", "Vegetable" },
				new string[] { "3", "Banana", "150", "Fruit" },
				new string[] { "4", "Apple", "200", "Fruit" },
			}, this.Sjis, ',');

			this.WriteCsv(this.InputFile("info.csv"), new string[][]
			{
				new string[] { "A", "B", "C" },
				new string[] { "1", "2", "3", "4" },
				new string[] { "" },
				new string[] { "x", "y" },
			}, this.Sjis, ',');

			this.WriteCsv(this.InputFile("tab.tsv"), new string[][]
			{
				new string[] { "Code", "Name", "Memo" },
				new string[] { "A", "Alpha", "Hello" },
				new string[] { "B", "Beta", "Tabbed" },
			}, this.Sjis, '\t');

			this.WriteCsv(this.InputFile("space.ssv"), new string[][]
			{
				new string[] { "Code", "Name", "Memo" },
				new string[] { "A", "Alpha", "Hello" },
				new string[] { "B", "Beta", "World" },
			}, this.Sjis, ' ');

			this.WriteCsv(this.InputFile("utf8bom.csv"), new string[][]
			{
				new string[] { "ID", "Name" },
				new string[] { "1", "桃" },
				new string[] { "2", "梨" },
			}, this.Utf8Bom, ',');

			this.WriteCsv(this.InputFile(@"merge\001.csv"), new string[][]
			{
				new string[] { "ID", "Name" },
				new string[] { "1", "Alpha" },
				new string[] { "2", "Beta" },
			}, this.Sjis, ',');

			this.WriteCsv(this.InputFile(@"merge\002.csv"), new string[][]
			{
				new string[] { "ID", "Name" },
				new string[] { "3", "Gamma" },
				new string[] { "4", "Delta" },
			}, this.Sjis, ',');

			this.WriteCsv(this.InputFile(@"merge\ignored.tsv"), new string[][]
			{
				new string[] { "ID", "Name" },
				new string[] { "8", "Ignored" },
			}, this.Sjis, '\t');

			this.WriteCsv(this.InputFile(@"merge\only.part"), new string[][]
			{
				new string[] { "ID", "Name" },
				new string[] { "9", "PartOnly" },
			}, this.Sjis, ',');

			this.WriteCsv(this.InputFile(@"excel-sheets\alpha.csv"), new string[][] { new string[] { "ID", "Name" }, new string[] { "1", "AlphaName" } }, this.Sjis, ',');
			this.WriteCsv(this.InputFile(@"excel-sheets\beta.csv"), new string[][] { new string[] { "ID", "Name" }, new string[] { "2", "BetaName" } }, this.Sjis, ',');
			this.WriteCsv(this.InputFile(@"excel-sheets\only.part"), new string[][] { new string[] { "ID", "Name" }, new string[] { "3", "PartName" } }, this.Sjis, ',');
			File.WriteAllText(this.InputFile("unsupported.txt"), "unsupported", this.Sjis);
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

		private void AssertWorkbookBasics(string file, int expectedSheetCount, string message)
		{
			this.AssertTrue(File.Exists(file), message + " output exists");

			using (ZipArchive archive = ZipFile.OpenRead(file))
			{
				this.AssertTrue(archive.GetEntry("[Content_Types].xml") != null, message + " content types exists");
				this.AssertTrue(archive.GetEntry("xl/workbook.xml") != null, message + " workbook exists");
				this.AssertEquals(expectedSheetCount, archive.Entries.Count(entry => entry.FullName.StartsWith("xl/worksheets/sheet") && entry.FullName.EndsWith(".xml")), message + " worksheet count");
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

		private void AssertZipEntryContains(string zipFile, string entryName, string expectedPart, string message)
		{
			using (ZipArchive archive = ZipFile.OpenRead(zipFile))
			{
				ZipArchiveEntry entry = archive.GetEntry(entryName);

				if (entry == null)
					throw new Exception("ZIP entry not found: " + entryName);

				this.AssertContains(this.ReadZipEntryText(entry), expectedPart, message);
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
			Console.WriteLine("====              ECWeaver2 Test0001 SUCCESS           ====");
			Console.WriteLine("====                                                    ====");
			Console.WriteLine("====              ALL TESTS PASSED                     ====");
			Console.WriteLine("====                                                    ====");
			Console.WriteLine("============================================================");
			Console.WriteLine("============================================================");
			Console.WriteLine("");
		}
#endif
	}
}
