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
    /// FIXED
    /// Will preserve all properties and sizes of tables or charts. Without FIXED tables and charts will automatically adapt to changed number of rows and columns by resizing them. FIXED will try to keep them the same size. This may result in the whole size of object changing if there is more data than originally predicted.
    /// </summary>
    public class FixedCommand : Command
    {
        public static readonly string Name = "FIXED";

        public override void ParseArguments()
        {
        }

        public override void ApplyToData(DataElement data)
        {
        }
    }
}
