using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;

namespace McKinsey.PowerPointGenerator.Extensions
{
    public static class ParagraphExtensions
    {
        public static void Replace(this Paragraph paragraph, string oldValue, string newValue, SolidFill fill = null)
        {
            int index = 0;
            Run run = null;
            for (; index < paragraph.Elements<Run>().Count(); index++)
            {
                run = paragraph.Elements<Run>().ElementAt(index);
                if (run.Text.Text.Contains("#"))
                {
                    break;
                }
            }
            while (!run.Text.Text.Contains(oldValue))
            {
                run.Text.Text = run.Text.Text + paragraph.Elements<Run>().ElementAt(index + 1).Text.Text;
                paragraph.Elements<Run>().ElementAt(index + 1).Remove();
            }
            run.Replace(oldValue, newValue, fill);
            paragraph.RemoveAllChildren<Field>();
        }

        public static void Replace(this Paragraph paragraph, string newValue, SolidFill fill = null)
        {
            Run run = null;
            if (paragraph.Elements<Run>().Count() == 0)
            {
                run = CreateNewRun(paragraph);
                paragraph.InsertAt<Run>(run, 0);
            }
            else
            {
                run = paragraph.Elements<Run>().First();
            }
            while (paragraph.Elements<Run>().Count() > 1)
            {
                paragraph.RemoveChild<Run>(paragraph.Elements<Run>().Last());
            }
            run.Replace(newValue, fill);
            paragraph.RemoveAllChildren<Field>();
        }

        private static Run CreateNewRun(Paragraph paragraph)
        {
            Run run = new Run() { Text = new Text() };
            EndParagraphRunProperties runProperties = paragraph.Elements<EndParagraphRunProperties>().FirstOrDefault();
            if (runProperties != null)
            {
                RunProperties newRunProperties = new RunProperties() { Language = runProperties.Language, FontSize = runProperties.FontSize, Dirty = runProperties.Dirty };
                foreach (var prop in runProperties.Elements())
                {
                    newRunProperties.Append(prop.CloneNode(true));
                }
                run.InsertAt(newRunProperties, 0);
            }
            return run;
        }
    }
}
