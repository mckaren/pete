using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Extensions;
using NCalc;

namespace McKinsey.PowerPointGenerator.Commands
{
    public static class FormulaHelper
    {
        private static Regex formulaRegex = new Regex(@"(?<variable>\[.*?\])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static T Evaluate<T>(DataElement data, Row row, List<Index> usedIndexes, string formula, int columnIndex = 0)
        {
            IEnumerable<Index> actualUsedIndexes = usedIndexes.Where(i => data.HasColumn(i));
            string expressionText = formula.ToLower();
            Expression expression = new Expression(expressionText);
            foreach (Index usedColumnIndex in actualUsedIndexes)
            {
                object dataItemValue = null;
                string name = usedColumnIndex.Name;
                if (usedColumnIndex.Number.HasValue)
                {
                    dataItemValue = row[usedColumnIndex.Number.Value];
                    name = "c#" + usedColumnIndex.Number.Value;
                }
                else
                {
                    dataItemValue = row[usedColumnIndex.Name];
                }
                AddParameter(expression, name.ToLower(), dataItemValue);
            }
            if (expressionText.Contains("[value]"))
            {
                AddParameter(expression, "value", row[columnIndex]);
            }
            T result = (T)expression.Evaluate();
            return result;
        }

        public static string ParseFormulaIndexes(string argumentsString, List<Index> usedIndexes)
        {
            var formula = argumentsString.Trim('"');
            var matches = formulaRegex.Matches(formula.ToLower());
            foreach (Match m in matches)
            {
                if (!m.Groups["variable"].Value.Equals("[value]", StringComparison.OrdinalIgnoreCase))
                {
                    Index newIndex = null;
                    string var = m.Groups["variable"].Value.Trim('[', ']');
                    int columnIndex = 0;
                    if (int.TryParse(var, out columnIndex))
                    {
                        newIndex = new Index(columnIndex);
                        formula = formula.Replace(m.Groups["variable"].Value, "[c#" + columnIndex + "]");
                    }
                    else
                    {
                        newIndex = new Index(var);
                    }
                    if (newIndex != null && !usedIndexes.Any(i => i == newIndex))
                    {
                        usedIndexes.Add(newIndex);
                    }
                }
            }
            return formula;
        }

        private static void AddParameter(Expression expression, string name, object dataItemValue)
        {
            object parameterValue = null;
            if (dataItemValue == null)
            {
                parameterValue = string.Empty;
            }
            if (dataItemValue is string)
            {
                parameterValue = dataItemValue.ToString().ToLower();
            }
            if (dataItemValue is int)
            {
                parameterValue = (int)dataItemValue;
            }
            if (dataItemValue is long)
            {
                parameterValue = (long)dataItemValue;
            }
            if (dataItemValue is double)
            {
                parameterValue = (double)dataItemValue;
            }
            if (dataItemValue is decimal)
            {
                parameterValue = (decimal)dataItemValue;
            }
            if (dataItemValue is DateTime)
            {
                parameterValue = (DateTime)dataItemValue;
            }
            if (dataItemValue is bool)
            {
                parameterValue = (bool)dataItemValue;
            }
            expression.Parameters.Add(name, parameterValue);
        }
    }
}
