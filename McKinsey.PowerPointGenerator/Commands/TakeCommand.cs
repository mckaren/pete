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
    /// TAKE{n}
    /// Take top n nows
    /// Example:
    /// TOP{5}
    /// </summary>
    public class TakeCommand : Command
    {
        public static readonly string Name = "TAKE";
        public int RowsToTake { get; set; }

        public override void ParseArguments()
        {
            RowsToTake = 0;
            int rows = 0;
            if (int.TryParse(ArgumentsString, out rows))
            {
                RowsToTake = rows;
            }
        }

        public override void ApplyToData(DataElement data)
        {
            var tmpRows = data.Rows.Take(RowsToTake).ToList();
            data.Rows.Clear();
            data.Rows.AddRange(tmpRows);
            foreach (Column column in data.Columns)
            {
                var tmpData = column.Data.Take(RowsToTake).ToList();
                var tmpLegends = column.Legends.Take(RowsToTake).ToList();
                column.Data.Clear();
                column.Data.AddRange(tmpData);
                column.Legends.Clear();
                column.Legends.AddRange(tmpLegends);
            }
        }
    }
}
