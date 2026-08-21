using System;
using System.Collections.Generic;
using System.Linq;
using HLTStudio.Commons;

namespace HLTStudio.CommandLine
{
	public class CommandLineArgs
	{
		private static readonly string[] ValueOptionNames = new string[]
		{
			"engine",
			"encoding",
			"delimiter",
			"newline",
			"input-list",
			"log",
			"sheet",
			"sheets",
			"range",
			"password",
			"columns",
			"headers",
			"pattern",
			"to-excel",
			"to-csv-dir",
			"to-same-dir",
			"output",
			"index",
			"from",
			"to",
			"regex",
			"column",
			"header",
			"equals",
			"contains",
			"key-columns",
			"printer",
			"set",
			"set-file",
		};

		private static readonly string[] FlagOptionNames = new string[]
		{
			"help",
			"version",
			"overwrite",
			"silent",
			"verbose",
			"no-dialog",
			"has-header",
			"invert",
			"skip-header",
			"numeric",
			"desc",
			"stop-on-error",
			"continue-on-error",
		};

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
			return args;
		}

		public bool HasOption(string name)
		{
			return this._options.ContainsKey(NormalizeOptionName(name));
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

			if (this._options.TryGetValue(NormalizeOptionName(name), out values))
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

					if (!hasValue && IsValueOption(name))
					{
						if (this._rawArgs.Count <= index + 1)
							throw new Exception("Missing command line option value: --" + name);

						value = this._rawArgs[++index];
						hasValue = true;
					}
					else if (!hasValue && !IsFlagOption(name))
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
			name = NormalizeOptionName(name);

			List<string> values;

			if (!this._options.TryGetValue(name, out values))
			{
				values = new List<string>();
				this._options.Add(name, values);
			}
			values.Add(value);
		}

		private static bool IsOptionToken(string token)
		{
			return token != null && token.StartsWith("--") && 2 < token.Length;
		}

		private static void SplitOptionToken(string token, out string name, out string value, out bool hasValue)
		{
			string option = token.Substring(2);
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

			name = NormalizeOptionName(name);

			if (name.Length == 0)
				throw new Exception("Bad command line option: " + token);
		}

		private static string NormalizeOptionName(string name)
		{
			if (name == null)
				throw new ArgumentNullException("name");

			while (name.StartsWith("--"))
				name = name.Substring(2);

			return name.ToLowerInvariant();
		}

		private static bool IsValueOption(string name)
		{
			return ValueOptionNames.Contains(NormalizeOptionName(name), StringComparer.OrdinalIgnoreCase);
		}

		private static bool IsFlagOption(string name)
		{
			return FlagOptionNames.Contains(NormalizeOptionName(name), StringComparer.OrdinalIgnoreCase);
		}
	}
}
