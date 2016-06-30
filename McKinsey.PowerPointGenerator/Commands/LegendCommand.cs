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
    /// <summary>
    /// LEGEND_FROM{columnName, formula, legendObject}
    /// Specifies index at which value will be used to find the legend on the slide. Works only when element is mapped to range with more than one column
    /// Examples:
    /// LEGEND_FROM{"column 1", "value = 'Q1'", "Rectangle 1"}
    /// </summary>
    public class LegendCommand : Command, IUseIndexes
    {
        private static Regex regex = new Regex(@"^(?<index>(?:\"".*?\"")|(?:\d*?)),\s*\""(?<formula>.*?)\"",\s*\""(?<target>.*?)\""$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static Regex formulaRegex = new Regex(@"(?<variable>'.*?')", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static readonly string Name = "LEGEND";
        public List<Index> UsedIndexes { get; set; }
        public Index Index { get; set; }
        public string Formula { get; set; }
        public string LegendObjectName { get; set; }
        public ShapeElement LegendObject { get; set; }

        public LegendCommand()
        {
            UsedIndexes = new List<Index>();
        }

        public override void ParseArguments()
        {
            Match match = regex.Match(ArgumentsString);
            if (match.Success)
            {
                Index = new Index(match.Groups["index"].Value);
                Formula = match.Groups["formula"].Value;
                LegendObjectName = match.Groups["target"].Value;
            }
            if (TargetElement.Slide != null && TargetElement.Slide.Slide != null)
            {
                var shape = TargetElement.Slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().FirstOrDefault(s => s.ElementName().Equals(LegendObjectName, StringComparison.OrdinalIgnoreCase));
                if (shape != null)
                {
                    LegendObject = ShapeElement.Create(LegendObjectName, shape, TargetElement.Slide);
                }
            }
            Formula = FormulaHelper.ParseFormulaIndexes(Formula, UsedIndexes);
        }

        public override void ApplyToData(DataElement data)
        {
            if (Index.IsAll)
            {
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    var column = data.Columns[i];
                    ApplyToColumn(data, column);
                }
            }
            else
            {
                var column = data.Column(Index);
                ApplyToColumn(data, column);
            }
        }

        private void ApplyToColumn(DataElement data, Column column)
        {
            int columnIdx = data.Columns.IndexOf(column);
            for (int rowInd = 0; rowInd < column.Data.Count; rowInd++)
            {
                if (FormulaHelper.Evaluate<bool>(data, data.Rows[rowInd], UsedIndexes, Formula, columnIdx))
                {
                    if (column.Legends.Count <= rowInd)
                    {
                        column.Legends.Add(LegendObject);
                    }
                    else
                    {
                        column.Legends[rowInd] = LegendObject;
                    }

                    if (data.Rows[rowInd].Legends.Count <= columnIdx)
                    {
                        column.Legends.Add(LegendObject);
                    }
                    else
                    {
                        data.Rows[rowInd].Legends[columnIdx] = LegendObject;
                    }
                }
            }
        }
    }
}