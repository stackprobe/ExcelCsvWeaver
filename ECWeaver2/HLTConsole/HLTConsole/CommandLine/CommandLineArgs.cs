using System;
using System.Collections.Generic;
using System.Linq;
using HLTStudio.Commons;

namespace HLTStudio.CommandLine
{
	public class CommandLineArgs
	{
		private readonly List<string> _rawArgs = new List<string>();
		private readonly List<string> _arguments = new List<string>();
		private readonly Dictionary<string, List<string>> _options = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

		public string Command { get; private set; }

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

		public static CommandLineArgs Read(ArgsReader ar)
		{
			CommandLineArgs args = new CommandLineArgs();

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
			return this._options.ContainsKey(CommandLineConsts.NormalizeOptionName(name));
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

			if (this._options.TryGetValue(CommandLineConsts.NormalizeOptionName(name), out values))
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

					if (!hasValue && CommandLineConsts.IsValueOption(name))
					{
						if (this._rawArgs.Count <= index + 1)
							throw new Exception("Missing command line option value: " + CommandLineConsts.OptionPrefix + name);

						value = this._rawArgs[++index];
						hasValue = true;
					}
					else if (!hasValue && !CommandLineConsts.IsFlagOption(name))
					{
						if (index + 1 < this._rawArgs.Count && !IsOptionToken(this._rawArgs[index + 1]))
						{
							value = this._rawArgs[++index];
							hasValue = true;
						}
					}

					this.AddOption(name, hasValue ? value : null);
				}
				else if (this.Command == null)
				{
					this.Command = arg;
				}
				else
				{
					this._arguments.Add(arg);
				}
			}
		}

		private void AddOption(string name, string value)
		{
			name = CommandLineConsts.NormalizeOptionName(name);

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

			if (this.Command == null)
				return;

			switch (this.Command.ToLowerInvariant())
			{
				case CommandLineConsts.Commands.ExcelToCsv:
				case CommandLineConsts.Commands.ExcelToTsv:
				case CommandLineConsts.Commands.CsvToExcel:
				case CommandLineConsts.Commands.CsvsToExcel:
				case CommandLineConsts.Commands.ExcelToPdf:
				case CommandLineConsts.Commands.CsvSelectColumns:
				case CommandLineConsts.Commands.CsvFilterRows:
				case CommandLineConsts.Commands.CsvReplace:
				case CommandLineConsts.Commands.CsvMerge:
				case CommandLineConsts.Commands.CsvSort:
				case CommandLineConsts.Commands.CsvUnique:
				case CommandLineConsts.Commands.ExcelExtractPictures:
				case CommandLineConsts.Commands.ExcelReplaceText:
				case CommandLineConsts.Commands.ExcelReplacePlaceholder:
					this.NormalizeArgumentPaths(0, 1);
					break;

				case CommandLineConsts.Commands.CsvInfo:
				case CommandLineConsts.Commands.ExcelListSheets:
				case CommandLineConsts.Commands.ExcelInfo:
				case CommandLineConsts.Commands.CsvValidate:
				case CommandLineConsts.Commands.ExcelValidate:
				case CommandLineConsts.Commands.Print:
				case CommandLineConsts.Commands.RunScript:
					this.NormalizeArgumentPaths(0);
					break;

				case CommandLineConsts.Commands.ExcelReplacePicture:
				case CommandLineConsts.Commands.CsvDiff:
				case CommandLineConsts.Commands.ExcelDiff:
					this.NormalizeArgumentPaths(0, 1, 2);
					break;

				case CommandLineConsts.Commands.Weave:
					this.NormalizeAllArgumentPaths();
					break;
			}
		}

		private void NormalizePathOptions()
		{
			this.NormalizeOptionValues(CommandLineConsts.Options.InputList);
			this.NormalizeOptionValues(CommandLineConsts.Options.Log);
			this.NormalizeOptionValues(CommandLineConsts.Options.ToExcel);
			this.NormalizeOptionValues(CommandLineConsts.Options.ToCsvDir);
			this.NormalizeOptionValues(CommandLineConsts.Options.ToSameDir);
			this.NormalizeOptionValues(CommandLineConsts.Options.Output);
			this.NormalizeOptionValues(CommandLineConsts.Options.SetFile);
		}

		private void NormalizeOptionValues(string name)
		{
			List<string> values;

			if (!this._options.TryGetValue(CommandLineConsts.NormalizeOptionName(name), out values))
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
			return token != null && token.StartsWith(CommandLineConsts.OptionPrefix) && CommandLineConsts.OptionPrefix.Length < token.Length;
		}

		private static void SplitOptionToken(string token, out string name, out string value, out bool hasValue)
		{
			string option = token.Substring(CommandLineConsts.OptionPrefix.Length);
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

			name = CommandLineConsts.NormalizeOptionName(name);

			if (name.Length == 0)
				throw new Exception("Bad command line option: " + token);
		}

	}
}
