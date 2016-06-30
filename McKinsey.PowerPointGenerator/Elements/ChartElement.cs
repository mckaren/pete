using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using McKinsey.PowerPointGenerator.Extensions;

namespace McKinsey.PowerPointGenerator.Elements
{
    [DebuggerDisplay("{Name}, type: chart")]
    public class ChartElement : ShapeElementBase
    {
        public static string TypeNameId = "ChartElement";
        public GraphicFrame ChartFrame { get; set; }
        public bool IsWaterfall { get; set; }
        public bool IsFixed { get; set; }
        public bool IsPaged { get; set; }
        public override string TypeName { get { return TypeNameId; } }

        public static ChartElement Create(string name, GraphicFrame element, SlideElement slide)
        {
            ChartElement chart = new ChartElement();
            if (!chart.Parse(name, slide))
            {
                return null;
            }
            chart.ChartFrame = element;
            return chart;
        }

        public override IEnumerable<Commands.Command> PreprocessSwitchCommands(IEnumerable<Commands.Command> discoveredCommands)
        {
            var processedCommands = new List<Commands.Command>(discoveredCommands);
            if (processedCommands.Any(c => (c as Commands.WaterfallCommand) != null))
            {
                IsWaterfall = true;
            }
            else
            {
                IsWaterfall = false;
            }
            if (processedCommands.Any(c => (c as Commands.FixedCommand) != null))
            {
                IsFixed = true;
                processedCommands.Remove(processedCommands.First(c => (c as Commands.FixedCommand) != null));
            }
            else
            {
                IsFixed = false;
            }
            if (processedCommands.Any(c => (c as Commands.PageCommand) != null))
            {
                //we don't remove paging from the list as it will have to be processed later
                IsPaged = true;
            }
            else
            {
                IsPaged = false;
            }
            return processedCommands;
        }
    }
}
