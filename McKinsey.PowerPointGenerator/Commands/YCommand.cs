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
    public class YCommand : Command
    {
        public static readonly string Name = "Y";
        private static Regex regex = new Regex(@"^(?<index>(?:\"".*?\"")|(?:\d*?))?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public Index Index { get; set; }

        public override void ParseArguments()
        {
            Match match = regex.Match(ArgumentsString);
            if (match.Success)
            {
                Index = new Index(match.Groups["index"].Value);
            }
        }

        public override void ApplyToData(DataElement data)
        {
        }
    }
}
