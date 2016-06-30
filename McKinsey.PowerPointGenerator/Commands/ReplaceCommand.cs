using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Elements;
using System.Data;
using McKinsey.PowerPointGenerator.Core.Data;

namespace McKinsey.PowerPointGenerator.Commands
{
    /// <summary>
    /// REPLACE{value=replacement}
    /// Will replace values
    /// Examples:
    /// REPLACE{true=ü}
    /// REPLACE{"true"="ü", "false"="", "1"="one", "client"="IBM"}
    /// </summary>
    public class ReplaceCommand : Command
    {
        private static Regex regex = new Regex(@"\""(?<from>.*?)\""\s*=\s*\""(?<to>.*?)\""(?:,?\s*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly string Name = "REPLACE";
        public Dictionary<string, string> Substitutions { get; private set; }

        public ReplaceCommand()
        {
            Substitutions = new Dictionary<string, string>();
        }

        public override void ParseArguments()
        {
            Substitutions.Clear();
            if (regex.Match(ArgumentsString).Success)
            {
                foreach (Match match in regex.Matches(ArgumentsString))
                {
                    string from = match.Groups["from"].Value;
                    string to = match.Groups["to"].Value;
                    if (!Substitutions.Any(s => s.Key == from))
                    {
                        Substitutions.Add(from, to);
                    }
                }
            }
        }

        public override void ApplyToData(DataElement data)
        {
            for (int rowIdx = 0; rowIdx < data.Rows.Count; rowIdx++)
            {
                for (int columnIdx = 0; columnIdx < data.Rows[rowIdx].Data.Count; columnIdx++)
                {
                    string valueToCheck = data.Rows[rowIdx].Data[columnIdx] == null ? "" : data.Rows[rowIdx].Data[columnIdx].ToString();
                    if (Substitutions.Any(s => s.Key.Equals(valueToCheck, StringComparison.OrdinalIgnoreCase)))
                    {
                        string newValue = Substitutions.First(s => s.Key.Equals(valueToCheck, StringComparison.OrdinalIgnoreCase)).Value;
                        data.Rows[rowIdx].Data[columnIdx] = newValue;
                        data.Columns[columnIdx].Data[rowIdx] = newValue;
                    }
                }
            }
        }

    }
}
