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
    /// VISIBLE{formula                      
    /// Evaluates the formula. If false then object is removed
    /// Examples:
    /// VISIBLE{"'column' > 5"}
    /// VISIBLE{"'column 1' > 5 OR 'column 2' <- 5"}
    /// </summary>
    public class VisibleCommand : Command, IUseIndexes
    {
        public static readonly string Name = "VISIBLE";
        public List<Index> UsedIndexes { get; set; }
        public Index Index { get; set; }
        public string Formula { get; set; }
        public bool IsVisible { get; set; }

        public VisibleCommand()
        {
            UsedIndexes = new List<Index>();
        }

        public override void ParseArguments()
        {
            Formula = FormulaHelper.ParseFormulaIndexes(ArgumentsString, UsedIndexes);
        }

        public override void ApplyToData(DataElement data)
        {
            IsVisible = true;
            if (!FormulaHelper.Evaluate<bool>(data, data.Rows[0], UsedIndexes, Formula))
            {
                IsVisible = false;
            }
        }
    }
}
