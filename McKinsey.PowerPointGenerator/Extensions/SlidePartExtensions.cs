using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace McKinsey.PowerPointGenerator.Extensions
{
    public static class SlidePartExtensions
    {
        public static bool IsHidden(this SlidePart slidePart)
        {
            if (slidePart.Slide.Show != null && slidePart.Slide.Show.HasValue && !slidePart.Slide.Show.Value)
            {
                return true;
            }
            return false;
        }
    }
}
