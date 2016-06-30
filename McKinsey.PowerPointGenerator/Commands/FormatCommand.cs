using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Core.Data;

namespace McKinsey.PowerPointGenerator.Commands
{
    /// <summary>
    /// FORMAT{format_string, culture}                      
    /// Formats the text using .NET format string.  When culture is not specified default culture of the operating system will be used.
    /// Examples:
    /// FORMAT{"YYYY mmmm"}   
    /// FORMAT{"##,#", "en-GB"}   
    /// FORMAT{"##,#", "de-DE"}
    /// </summary>
    public class FormatCommand : Command
    {
        private static Regex fommatRegex = new Regex(@"^(?:\[(?<index>(?:\"".*?\"")|(?:\d*?))\])?,?\s*\""(?<format>.*?)\""(?:,\s*\""(?<culture>.*?)\"")?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        

        public static readonly string Name = "FORMAT";
        public string FormatString { get; set; }
        public CultureInfo Culture { get; set; }
        public Index Index { get; set; }

        public override void ParseArguments()
        {
            Match match = fommatRegex.Match(ArgumentsString);
            if (match.Success)
            {
                if (!string.IsNullOrEmpty(match.Groups["index"].Value))
                {
                    Index = new Index(match.Groups["index"].Value);
                }
                FormatString = match.Groups["format"].Value;
                var culture = match.Groups["culture"].Value;
                if (string.IsNullOrEmpty(culture))
                {
                    Culture = CultureInfo.CurrentUICulture;
                }
                else
                {
                    try
                    {
                        Culture = new CultureInfo(culture);
                    }
                    catch (CultureNotFoundException)
                    {
                        Culture = CultureInfo.CurrentUICulture;
                    }
                }
            }
        }

        public override void ApplyToData(DataElement data)
        {
            if (string.IsNullOrEmpty(FormatString))
            {
                return;
            }

            if (Index == null)
            {
                for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++)
                {
                    FormatColumn(data, columnIndex);
                }
            }
            else
            {
                Column column = data.Column(Index);
                int columnIndex = data.Columns.IndexOf(column);
                FormatColumn(data, columnIndex);
            }
        }

        private void FormatColumn(DataElement data, int columnIndex)
        {
            for (int rowIndex = 0; rowIndex < data.Columns[columnIndex].Data.Count; rowIndex++)
            {
                if (data.Columns[columnIndex].Data[rowIndex] != null)
                {
                    data.Columns[columnIndex].Data[rowIndex] = string.Format(Culture, "{0:" + FormatString + "}", data.Columns[columnIndex].Data[rowIndex]);
                    data.Rows[rowIndex].Data[columnIndex] = data.Columns[columnIndex].Data[rowIndex];
                }
            }
        }
    }
}
