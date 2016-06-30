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
    public abstract class Command
    {
        public string ArgumentsString { get; set; }
        public ShapeElementBase TargetElement { get; set; }

        public abstract void ParseArguments();

        /// <summary>
        /// Applies the command to the data element. The command may modify the data element. If the command will create a new data element (i.e. Page) then new element(s) are returned.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        public abstract void ApplyToData(DataElement data);
    }
}
