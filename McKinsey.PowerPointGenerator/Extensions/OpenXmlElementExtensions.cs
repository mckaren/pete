using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace McKinsey.PowerPointGenerator.Extensions
{
    public static class OpenXmlElementExtensions
    {
        /// <summary>
        /// Gets the first element of the specified type.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="container">The container.</param>
        /// <returns>The element, null if not found</returns>
        public static T FirstElement<T>(this OpenXmlElement container) where T : OpenXmlElement
        {
            if (container.Elements<T>().Count() > 0)
            {
                return container.Elements<T>().First();
            }
            return null;
        }
    }
}
