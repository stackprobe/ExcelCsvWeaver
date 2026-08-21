namespace HLTStudio.CommandLine
{
	public static class CommandLineConsts
	{
		public const string OptionPrefix = "--";

		public static class Commands
		{
			public const string Help = "help";
			public const string Version = "version";
			public const string ExcelToCsv = "excel-to-csv";
			public const string ExcelToTsv = "excel-to-tsv";
			public const string CsvToExcel = "csv-to-excel";
			public const string CsvsToExcel = "csvs-to-excel";
			public const string ExcelToPdf = "excel-to-pdf";
			public const string Weave = "weave";
			public const string CsvInfo = "csv-info";
			public const string CsvSelectColumns = "csv-select-columns";
			public const string CsvFilterRows = "csv-filter-rows";
			public const string CsvReplace = "csv-replace";
			public const string CsvMerge = "csv-merge";
			public const string CsvSort = "csv-sort";
			public const string CsvUnique = "csv-unique";
			public const string ExcelListSheets = "excel-list-sheets";
			public const string ExcelInfo = "excel-info";
			public const string ExcelExtractPictures = "excel-extract-pictures";
			public const string ExcelReplacePicture = "excel-replace-picture";
			public const string ExcelReplaceText = "excel-replace-text";
			public const string ExcelReplacePlaceholder = "excel-replace-placeholder";
			public const string CsvValidate = "csv-validate";
			public const string ExcelValidate = "excel-validate";
			public const string CsvDiff = "csv-diff";
			public const string ExcelDiff = "excel-diff";
			public const string Printers = "printers";
			public const string Print = "print";
			public const string RunScript = "run-script";
		}

		public static class Options
		{
			public const string Help = "help";
			public const string Version = "version";
			public const string Engine = "engine";
			public const string Overwrite = "overwrite";
			public const string Encoding = "encoding";
			public const string Delimiter = "delimiter";
			public const string Newline = "newline";
			public const string InputList = "input-list";
			public const string Log = "log";
			public const string Silent = "silent";
			public const string Verbose = "verbose";
			public const string NoDialog = "no-dialog";
			public const string Sheet = "sheet";
			public const string Sheets = "sheets";
			public const string Range = "range";
			public const string Password = "password";
			public const string HasHeader = "has-header";
			public const string Columns = "columns";
			public const string Headers = "headers";
			public const string Pattern = "pattern";
			public const string ToExcel = "to-excel";
			public const string ToCsvDir = "to-csv-dir";
			public const string ToSameDir = "to-same-dir";
			public const string Output = "output";
			public const string Index = "index";
			public const string From = "from";
			public const string To = "to";
			public const string Regex = "regex";
			public const string Column = "column";
			public const string Header = "header";
			public const string EqualsCondition = "equals";
			public const string Contains = "contains";
			public const string Invert = "invert";
			public const string SkipHeader = "skip-header";
			public const string Numeric = "numeric";
			public const string Desc = "desc";
			public const string KeyColumns = "key-columns";
			public const string Printer = "printer";
			public const string Set = "set";
			public const string SetFile = "set-file";
			public const string StopOnError = "stop-on-error";
			public const string ContinueOnError = "continue-on-error";
		}

		public static class Values
		{
			public const string Auto = "auto";
			public const string App = "app";
			public const string Interop = "interop";
			public const string Zip = "zip";
			public const string Sjis = "sjis";
			public const string Utf8 = "utf8";
			public const string Utf8Bom = "utf8bom";
			public const string Utf16Le = "utf16le";
			public const string Comma = "comma";
			public const string Tab = "tab";
			public const string Space = "space";
			public const string CrLf = "crlf";
			public const string Lf = "lf";
		}
	}
}
