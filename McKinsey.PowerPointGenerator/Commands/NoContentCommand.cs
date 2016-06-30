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
    /// NO_CONTENT
    /// The content of the object should not be replaced. Can be used to change properties of the object using legend but without afecting the content.
    /// </summary>
    public class NoContentCommand : Command
    {
        public static readonly string Name = "NO_CONTENT";

        public override void ParseArguments()
        {
        }

        public override void ApplyToData(DataElement data)
        {
        }
    }
}
