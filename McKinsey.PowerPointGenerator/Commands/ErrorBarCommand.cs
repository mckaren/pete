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
    public class ErrorBarCommand : Command
    {
        public static readonly string Name = "ERROR_BAR";
        private static Regex regex = new Regex(@"^(?<minusindex>(?:\"".*?\"")|(?:\d*?))(?:\s*\,\s*(?<plusindex>(?:\"".*?\"")|(?:\d*?)))?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public Index PlusIndex { get; set; }
        public Index MinusIndex { get; set; }

        public override void ParseArguments()
        {
            Match match = regex.Match(ArgumentsString);
            if (match.Success)
            {
                MinusIndex = new Index(match.Groups["minusindex"].Value);
                if (!string.IsNullOrEmpty(match.Groups["plusindex"].Value))
                {
                    PlusIndex = new Index(match.Groups["plusindex"].Value);
                }
            }
        }

        public override void ApplyToData(DataElement data)
        {
        }
    }
}
