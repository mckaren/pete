using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Core.Data;

namespace McKinsey.PowerPointGenerator.Commands
{
    /// <summary>
    /// SORT{index ASC|DESC}
    /// Sort by column
    /// Examples
    /// SORT{4}  (ASC is assumed)
    /// SORT{2 ASC}
    /// SORT{"Column 2" DESC}
    /// </summary>
    public class SortCommand : Command, IUseIndexes
    {
        private static Regex regex = new Regex(@"^(?<index>(?:\"".*?\"")|(?:\d*?))\s*(?<order>(?:ASC)|(?:DESC))?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static readonly string Name = "SORT";
        public Index Index { get; set; }
        public SortOrder SortOrder { get; set; }
        public List<Index> UsedIndexes { get; set; }

        public SortCommand()
        {
            UsedIndexes = new List<Index>();
        }

        public override void ParseArguments()
        {
            Match match = regex.Match(ArgumentsString);
            if (match.Success)
            {
                Index = new Index(match.Groups["index"].Value);
                UsedIndexes.Add(Index);
                SortOrder = match.Groups["order"].Value.Equals("DESC", StringComparison.OrdinalIgnoreCase) ? SortOrder.Descending : SortOrder.Ascending;
            }
        }

        public override void ApplyToData(DataElement data)
        {
            Dictionary<int, object> dataToSort = new Dictionary<int, object>();
            int columnNo = data.Columns.IndexOf(data.Column(Index));
            for (int i = 0; i < data.Rows.Count; i++)
            {
                dataToSort.Add(i, data.Rows[i].Data[columnNo]);
            }
            DataElement temp = data.Clone();
            data.Rows.Clear();
            foreach (Column column in data.Columns)
            {
                column.Data.Clear();
            }
            List<KeyValuePair<int, object>> sortedData;
            if (SortOrder == Commands.SortOrder.Ascending)
            {
                sortedData = dataToSort.OrderBy(k => k.Value).ToList();
            }
            else
            {
                sortedData = dataToSort.OrderByDescending(k => k.Value).ToList();
            }
            foreach (KeyValuePair<int, object> kvp in sortedData)
            {
                data.Rows.Add(temp.Rows[kvp.Key]);
            }

            data.Normalize();
        }
    }
}