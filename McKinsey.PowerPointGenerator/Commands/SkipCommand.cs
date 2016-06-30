using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;

namespace McKinsey.PowerPointGenerator.Commands
{
    /// <summary>
    /// SKIP{n}
    /// Skip n rows from the top
    /// Example:
    /// SKIP{10}
    /// </summary>
    public class SkipCommand : Command
    {
        public static readonly string Name = "SKIP";
        public int RowsToSkip { get; set; }

        public override void ParseArguments()
        {
            RowsToSkip = 0;
            int rows = 0;
            if (int.TryParse(ArgumentsString, out rows))
            {
                RowsToSkip = rows;
            }
        }

        public override void ApplyToData(DataElement data)
        {
            var tmpRows = data.Rows.Skip(RowsToSkip).ToList();
            data.Rows.Clear();
            data.Rows.AddRange(tmpRows);
            foreach (Column column in data.Columns)
            {
                var tmpData = column.Data.Skip(RowsToSkip).ToList();
                var tmpLegends = column.Legends.Skip(RowsToSkip).ToList();
                column.Data.Clear();
                column.Data.AddRange(tmpData);
                column.Legends.Clear();
                column.Legends.AddRange(tmpLegends);
            }
        }
    }
}
