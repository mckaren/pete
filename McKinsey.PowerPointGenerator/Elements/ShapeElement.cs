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
    [DebuggerDisplay("{Name}, type: shape")]
    public class ShapeElement : ShapeElementBase
    {
        public static string TypeNameId = "ShapeElement";
        public Shape Element { get; set; }
        public bool IsContentProtected { get; set; }
        public override string TypeName { get { return TypeNameId; } }

        public static ShapeElement Create(string name, Shape element, SlideElement slide)
        {
            ShapeElement shape = new ShapeElement();
            if (!shape.Parse(name, slide))
            {
                return null;
            }
            shape.Element = element;
            return shape;
        }

        public override IEnumerable<Commands.Command> PreprocessSwitchCommands(IEnumerable<Commands.Command> discoveredCommands)
        {
            var processedCommands = new List<Commands.Command>(discoveredCommands);
            if (processedCommands.Any(c => (c as Commands.NoContentCommand) != null))
            {
                IsContentProtected = true;
                processedCommands.Remove(processedCommands.First(c => (c as Commands.NoContentCommand) != null));
            }
            else
            {
                IsContentProtected = false;
            }
            return processedCommands;
        }
    }
}