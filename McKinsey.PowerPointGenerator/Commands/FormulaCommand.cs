using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using NCalc;
using NLog;

namespace McKinsey.PowerPointGenerator.Commands
{
    /// <summary>
    /// FORMULA{formula} 
    /// Evaluates the formula per row and adds result as an additional virtual column. The new column will have no header so it will have to be addressed by index.
    /// Examples:
    /// FORMULA{"c2 - c1"}
    /// FORMULA{"'column 2' - 'column 1'"}
    /// </summary>
    public class FormulaCommand : Command, IUseIndexes
    {
        private static Regex formulaRegex = new Regex(@"(?<variable>\[.*?\])", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly string Name = "FORMULA";
        public string Formula { get; set; }
        public bool IsInPlaceFormula { get; set; }
        public List<Index> UsedIndexes { get; set; }

        public FormulaCommand()
        {
            UsedIndexes = new List<Index>();
        }

        public override void ParseArguments()
        {
            UsedIndexes.Clear();
            Formula = FormulaHelper.ParseFormulaIndexes(ArgumentsString, UsedIndexes);
        }

        public override void ApplyToData(DataElement data)
        {
            try
            {
                string expressionText = Formula.ToLower();
                if (IsInPlaceFormula)
                {
                    CalculateInPlace(data, expressionText);
                }
                else
                {
                    CalculateWithColumnNames(data, expressionText);
                }
            }
            catch (Exception ex)
            {
                Logger logger = LogManager.GetLogger("Generator");
                logger.Debug("Unable to evaluate formula {0} in object {1} on slide {2}", Formula, TargetElement.FullName, TargetElement.Slide.Number);
            }
        }

        private void CalculateWithColumnNames(DataElement data, string expressionText)
        {
            int calculatedColumnsCount = data.Columns.Count(c => c.IsCalculated) + 1;
            Column calculated = new Column() { IsCalculated = true, Header = "Calc" + calculatedColumnsCount, ParentElement = data };
            foreach (Row row in data.Rows)
            {
                object result = FormulaHelper.Evaluate<object>(data, row, UsedIndexes, expressionText);
                calculated.Data.Add(result);
                row.Data.Add(result);
            }
            data.Columns.Add(calculated);
            return;
        }

        private void CalculateInPlace(DataElement data, string expressionText)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                object result = FormulaHelper.Evaluate<object>(data, data.Rows[i], UsedIndexes, expressionText);
                data.Rows[i].Data[0] = result;
                data.Columns[0].Data[i] = result;
            }
        }
    }
}
