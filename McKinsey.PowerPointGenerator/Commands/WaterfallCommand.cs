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
    /// WATERFALL
    /// Will calculate waterfall spread. The sum is indicated by "e" in the data range.
    /// </summary>
    public class WaterfallCommand : Command
    {
        public static readonly string Name = "WATERFALL";

        public override void ParseArguments()
        {
        }

        public override void ApplyToData(DataElement data)
        {
            Column initialData = data.Columns[0].Clone();
            initialData.Data[0] = 0;
            data.Columns.Clear();
            int columnsToInsert = data.Rows.Count - 1;
            for (int i = 0; i < columnsToInsert; i++)
            {
                Column column = new Column();
                column.Header = "column " + i;
                column.Data.AddRange(Enumerable.Repeat<object>(null, data.Rows.Count).ToList());
                column.Legends.AddRange(Enumerable.Repeat<object>(null, data.Rows.Count).ToList());
                double testValue = Convert.ToDouble(initialData.Data[i]);
                double sum = 0;
                if (testValue < 0)
                {
                    sum = initialData.Data.Skip(i).Select(v => Convert.ToDouble(v)).Where(v => v > 0).Sum();
                }
                else
                {
                    sum = initialData.Data.Skip(i + 1).Select(v => Convert.ToDouble(v)).Where(v => v > 0).Sum();
                }
                column.Data[i] = sum;
                column.Data[i + 1] = Math.Abs(Convert.ToDouble(initialData.Data[i + 1]));
                column.Legends[i + 1] = initialData.Legends[i + 1];
                if (i == 0)
                {
                    column.Legends[i] = initialData.Legends[i + 1];
                }
                data.Columns.Insert(0, column);
            }
            foreach (var row in data.Rows)
            {
                row.Data.Clear();
                row.Legends.Clear();
            }
            data.Normalize();
        }
    }
}
