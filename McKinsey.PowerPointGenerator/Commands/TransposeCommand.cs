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
    /// TRANSPOSE
    /// Transposes data
    /// </summary>
    public class TransposeCommand : Command
    {
        public static readonly string Name = "TRANSPOSE";

        public override void ParseArguments()
        {
        }

        public override void ApplyToData(DataElement data)
        {
            DataElement temp = data.Clone();
            data.Rows.Clear();
            data.Columns.Clear();
            foreach (Column column in temp.Columns)
            {
                Row row = new Row { Header = column.Header, MappedHeader = column.MappedHeader, IsHidden = column.IsHidden, IsCore = column.IsCore };
                row.Data.AddRange(column.Data);
                row.Legends.AddRange(column.Legends);
                data.Rows.Add(row);
            }
            foreach (Row row in temp.Rows)
            {
                data.Columns.Add(new Column { Header = row.Header, MappedHeader = row.MappedHeader, IsHidden = row.IsHidden, IsCore = row.IsCore });
            }
            data.Normalize();
            data.IsTransposed = !data.IsTransposed;
        }
    }
}
