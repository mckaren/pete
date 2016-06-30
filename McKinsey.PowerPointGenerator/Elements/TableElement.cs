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
    [DebuggerDisplay("{Name}, type: table")]
    public class TableElement : ShapeElementBase
    {
        public static string TypeNameId = "TableElement";
        public GraphicFrame TableFrame { get; set; }
        public bool IsFixed { get; set; }
        public bool IsPaged { get; set; }
        public bool IsDynamic { get; set; }
        public override string TypeName { get { return TypeNameId; } }

        public static TableElement Create(string name, GraphicFrame element, SlideElement slide)
        {
            TableElement table = new TableElement();
            if (!table.Parse(name, slide))
            {
                return null;
            }
            table.TableFrame = element;
            return table;
        }

        public override IEnumerable<Commands.Command> PreprocessSwitchCommands(IEnumerable<Commands.Command> discoveredCommands)
        {
            var processedCommands = new List<Commands.Command>(discoveredCommands);
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
            //if (processedCommands.Any(c => (c as Commands.DynamicCommand) != null))
            //{
            //    IsDynamic = true;
            //    processedCommands.Remove(processedCommands.First(c => (c as Commands.DynamicCommand) != null));
            //}
            //else
            //{
            //    IsDynamic = false;
            //}
            return processedCommands;
        }
    }
}
