using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Extensions;
using NLog;

namespace McKinsey.PowerPointGenerator.Elements
{
    public class DataElementDescriptor
    {
        private static Regex indexParseRegex = new Regex(@"(?<range>(?<from>\d+)\-(?<to>\d+))|(?<number>\d+)|(?:\""(?<name>.*?)\"")(?:\,\s*)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        internal string RowIndexesString { get; set; }
        internal string ColumnIndexesString { get; set; }
        public string Name { get; set; }

        public void ExtractColumnIndexes(List<Index> columnIndexes)
        {
            ExtractIndexes(columnIndexes, ColumnIndexesString);
        }

        public void ExtractRowIndexes(List<Index> rowIndexes)
        {
            ExtractIndexes(rowIndexes, RowIndexesString);
        }

        private void ExtractIndexes(List<Index> columnIndexes, string indexesString)
        {
            if (!string.IsNullOrEmpty(indexesString))
            {
                foreach (Match match in indexParseRegex.Matches(indexesString))
                {
                    Index newIndex = null;
                    if (!string.IsNullOrEmpty(match.Groups["range"].Value))
                    {
                        int fromIndexValue = int.Parse(match.Groups["from"].Value);
                        int toIndexValue = int.Parse(match.Groups["to"].Value);
                        for (int index = fromIndexValue; index <= toIndexValue; index++)
                        {
                            newIndex = new Index(index) { IsCore = true };
                            if (!columnIndexes.Any(i => i == newIndex))
                            {
                                columnIndexes.Add(newIndex);
                            }
                        }
                        continue;
                    }
                    else
                        if (!string.IsNullOrEmpty(match.Groups["number"].Value))
                        {
                            int indexValue = int.Parse(match.Groups["number"].Value);
                            newIndex = new Index(indexValue) { IsCore = true };
                        }
                        else
                        {
                            newIndex = new Index(match.Groups["name"].Value) { IsCore = true };
                        }
                    if (newIndex != null && !columnIndexes.Any(i => i == newIndex))
                    {
                        columnIndexes.Add(newIndex);
                    }
                }
            }
            else
            {
                columnIndexes.Add(new Index("*") { IsAll = true, IsCore = true });
            }
        }
    }
}
