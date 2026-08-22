using System;
using System.Collections.Generic;
using System.Linq;
using HLTStudio.Commons;

namespace HLTStudio.ECWArguments
{
	public class ECWeaverArgs
	{
		private readonly List<string> _rawArgs = new List<string>();
		private readonly List<string> _arguments = new List<string>();
		private readonly Dictionary<string, List<string>> _options = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

		public string Operation { get; private set; }

		public string[] RawArgs
		{
			get
			{
				return this._rawArgs.ToArray();
			}
		}

		public string[] Arguments
		{
			get
			{
				return this._arguments.ToArray();
			}
		}

		public string[] OptionNames
		{
			get
			{
				return this._options.Keys.ToArray();
			}
		}

		public static ECWeaverArgs Read(ArgsReader ar)
		{
			ECWeaverArgs args = new ECWeaverArgs();

			foreach (string arg in ar.TrailArgs())
			{
				args._rawArgs.Add(arg);
			}
			ar.End();

			args.ReadRawArgs();
			args.NormalizePathArgs();
			return args;
		}

		public bool HasOption(string name)
		{
			return this._options.ContainsKey(ECWeaverArgConsts.NormalizeOptionName(name));
		}

		public string GetOptionValue(string name, string defaultValue = null)
		{
			string[] values = this.GetOptionValues(name);

			if (values.Length == 0)
				return defaultValue;

			return values[values.Length - 1];
		}

		public string[] GetOptionValues(string name)
		{
			List<string> values;

			if (this._options.TryGetValue(ECWeaverArgConsts.NormalizeOptionName(name), out values))
				return values.ToArray();

			return new string[0];
		}

		private void ReadRawArgs()
		{
			for (int index = 0; index < this._rawArgs.Count; index++)
			{
				string arg = this._rawArgs[index];

				if (IsOptionToken(arg))
				{
					string name;
					string value;
					bool hasValue;

					SplitOptionToken(arg, out name, out value, out hasValue);

					if (!hasValue && ECWeaverArgConsts.IsValueOption(name))
					{
						if (this._rawArgs.Count <= index + 1)
							throw new Exception("Missing command line option value: " + ECWeaverArgConsts.OptionPrefix + name);

						value = this._rawArgs[++index];
						hasValue = true;
					}
					else if (!hasValue && !ECWeaverArgConsts.IsFlagOption(name))
					{
						if (index + 1 < this._rawArgs.Count && !IsOptionToken(this._rawArgs[index + 1]))
						{
							value = this._rawArgs[++index];
							hasValue = true;
						}
					}

					this.AddOption(name, hasValue ? value : null);
				}
				else if (this.Operation == null)
				{
					this.Operation = arg;
				}
				else
				{
					this._arguments.Add(arg);
				}
			}
		}

		private void AddOption(string name, string value)
		{
			name = ECWeaverArgConsts.NormalizeOptionName(name);

			List<string> values;

			if (!this._options.TryGetValue(name, out values))
			{
				values = new List<string>();
				this._options.Add(name, values);
			}
			values.Add(value);
		}

		private void NormalizePathArgs()
		{
			this.NormalizePathOptions();

			if (this.Operation == null)
				return;

			switch (this.Operation.ToLowerInvariant())
			{
				case ECWeaverArgConsts.Operations.ExcelToCsv:
				case ECWeaverArgConsts.Operations.ExcelToTsv:
				case ECWeaverArgConsts.Operations.CsvToExcel:
				case ECWeaverArgConsts.Operations.CsvsToExcel:
				case ECWeaverArgConsts.Operations.ExcelToPdf:
				case ECWeaverArgConsts.Operations.CsvSelectColumns:
				case ECWeaverArgConsts.Operations.CsvFilterRows:
				case ECWeaverArgConsts.Operations.CsvReplace:
				case ECWeaverArgConsts.Operations.CsvMerge:
				case ECWeaverArgConsts.Operations.CsvSort:
				case ECWeaverArgConsts.Operations.CsvUnique:
				case ECWeaverArgConsts.Operations.ExcelExtractPictures:
				case ECWeaverArgConsts.Operations.ExcelReplaceText:
				case ECWeaverArgConsts.Operations.ExcelReplacePlaceholder:
					this.NormalizeArgumentPaths(0, 1);
					break;

				case ECWeaverArgConsts.Operations.CsvInfo:
				case ECWeaverArgConsts.Operations.ExcelListSheets:
				case ECWeaverArgConsts.Operations.ExcelInfo:
				case ECWeaverArgConsts.Operations.CsvValidate:
				case ECWeaverArgConsts.Operations.ExcelValidate:
				case ECWeaverArgConsts.Operations.Print:
				case ECWeaverArgConsts.Operations.RunScript:
					this.NormalizeArgumentPaths(0);
					break;

				case ECWeaverArgConsts.Operations.ExcelReplacePicture:
				case ECWeaverArgConsts.Operations.CsvDiff:
				case ECWeaverArgConsts.Operations.ExcelDiff:
					this.NormalizeArgumentPaths(0, 1, 2);
					break;

				case ECWeaverArgConsts.Operations.Weave:
					this.NormalizeAllArgumentPaths();
					break;
			}
		}

		private void NormalizePathOptions()
		{
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.InputList);
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.Log);
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.ToExcel);
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.ToCsvDir);
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.ToSameDir);
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.Output);
			this.NormalizeOptionValues(ECWeaverArgConsts.Options.SetFile);
		}

		private void NormalizeOptionValues(string name)
		{
			List<string> values;

			if (!this._options.TryGetValue(ECWeaverArgConsts.NormalizeOptionName(name), out values))
				return;

			for (int index = 0; index < values.Count; index++)
			{
				if (values[index] != null)
					values[index] = SCommon.MakeFullPath(values[index]);
			}
		}

		private void NormalizeArgumentPaths(params int[] indexes)
		{
			foreach (int index in indexes)
				if (index < this._arguments.Count)
					this._arguments[index] = SCommon.MakeFullPath(this._arguments[index]);
		}

		private void NormalizeAllArgumentPaths()
		{
			for (int index = 0; index < this._arguments.Count; index++)
				this._arguments[index] = SCommon.MakeFullPath(this._arguments[index]);
		}

		private static bool IsOptionToken(string token)
		{
			return token != null && token.StartsWith(ECWeaverArgConsts.OptionPrefix) && ECWeaverArgConsts.OptionPrefix.Length < token.Length;
		}

		private static void SplitOptionToken(string token, out string name, out string value, out bool hasValue)
		{
			string option = token.Substring(ECWeaverArgConsts.OptionPrefix.Length);
			int delimiter = option.IndexOf('=');

			if (delimiter == -1)
			{
				name = option;
				value = null;
				hasValue = false;
			}
			else
			{
				name = option.Substring(0, delimiter);
				value = option.Substring(delimiter + 1);
				hasValue = true;
			}

			name = ECWeaverArgConsts.NormalizeOptionName(name);

			if (name.Length == 0)
				throw new Exception("Bad command line option: " + token);
		}

	}
}
