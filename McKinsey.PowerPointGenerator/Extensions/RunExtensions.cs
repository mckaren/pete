using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace McKinsey.PowerPointGenerator.Extensions
{
    public static class RunExtensions
    {
        /// <summary>
        /// Replaces the text.
        /// If the text content will have ## inside then value will be inserted in this place, otherwise the whole content will be replaced.
        /// </summary>
        /// <param name="run">The Run.</param>
        /// <param name="newValue">The new value.</param>
        /// <param name="oldValue">The old value.</param>
        /// <param name="fill">The fill.</param>
        public static void Replace(this Run run, string oldValue, string newValue, SolidFill fill = null)
        {
            if (fill != null)
            {
                run.ReplaceOrAppendProperty<SolidFill>(fill);
            }
            if (string.IsNullOrEmpty(oldValue))
            {
                run.Text.Text = newValue;
            }
            else
            {
                run.Text.Text = run.Text.Text.Replace(oldValue, newValue);
            }
        }

        /// <summary>
        /// Replaces the text.
        /// </summary>
        /// <param name="run">The Run.</param>
        /// <param name="newValue">The new value.</param>
        /// <param name="oldValue">The old value.</param>
        /// <param name="fill">The fill.</param>
        public static void Replace(this Run run, string newValue, SolidFill fill = null)
        {
            if (fill != null)
            {
                run.ReplaceOrAppendProperty<SolidFill>(fill);
            }
            run.Text.Text = newValue;
        }

        /// <summary>
        /// Replaces or appends property.
        /// </summary>
        /// <typeparam name="T">Type of the property</typeparam>
        /// <param name="run">The run.</param>
        /// <param name="newProperty">The new property.</param>
        public static void ReplaceOrAppendProperty<T>(this Run run, T newProperty) where T : OpenXmlElement
        {
            T replacedProperty = run.RunProperties.Elements<T>().FirstOrDefault();
            if (replacedProperty != null)
            {
                run.RunProperties.ReplaceChild(newProperty.CloneNode(true), replacedProperty);
            }
            else
            {
                run.RunProperties.InsertAt(newProperty.CloneNode(true), 0);
            }
        }
    }
}
