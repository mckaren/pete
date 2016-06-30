using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;

namespace McKinsey.PowerPointGenerator.Elements
{
    [DebuggerDisplay("{Name}, type: text, format: {Format}")]
    public class TextElement : ShapeElementBase
    {
        public static string TypeNameId = "TextElement";
        private static Regex regex = new Regex(@"^(?<name>[\w_]+)(?:\[(?<columns>.*?)\])?(?:\[(?<rows>.*?)\])?(?:(?:\:+)(?<format>.*))?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public Paragraph Paragraph { get; set; }
        public override string TypeName { get { return TypeNameId; } }

        public static TextElement Create(string name, Paragraph element, SlideElement slide)
        {
            TextElement text = new TextElement();
            text.Paragraph = element;
            text.FullName = name.Trim();
            Match match = regex.Match(text.FullName.Replace('“', '"').Replace('”', '"'));
            if (match.Success)
            {
                text.Name = match.Groups["name"].Value;
                text.DataDescriptor.RowIndexesString = match.Groups["rows"].Value;
                text.DataDescriptor.ColumnIndexesString = match.Groups["columns"].Value;
                string format = match.Groups["format"].Value;
                if (!string.IsNullOrEmpty(format))
                {
                    text.CommandString = "FORMAT" + format;
                }
                text.ExtractIndexes(text.DataDescriptor);
                return text;
            }
            return null;
        }

        public override IEnumerable<Commands.Command> PreprocessSwitchCommands(IEnumerable<Commands.Command> discoveredCommands)
        {
            return discoveredCommands;
        }
    }
}
