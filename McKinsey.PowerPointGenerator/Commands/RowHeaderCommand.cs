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
    /// ROW_HEADER
    /// Indicates that the specified data element has no row headers. Column 0 will be the first colummn of the range, otherwise the first column will be interpreted as the row headers.
    /// </summary>
    public class RowHeaderCommand : Command
    {
        public static readonly string Name = "ROW_HEADER";

        public override void ParseArguments()
        {
        }

        public override void ApplyToData(DataElement data)
        {
            data.HasRowHeaders = false;
        }
    }
}
