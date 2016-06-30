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
    /// COLUMN_HEADER
    /// Indicates that the specified data element has no column headers. Row 0 will be the first row of the range, otherwise the first row will be interpreted as the column headers.
    /// </summary>
    public class ColumnHeaderCommand : Command
    {
        public static readonly string Name = "COLUMN_HEADER";

        public override void ParseArguments()
        {
        }

        public override void ApplyToData(DataElement data)
        {
            data.HasColumnHeaders = false;
        }
    }
}
