using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace McKinsey.PowerPointGenerator.Processing
{
    public static class SlideProcessor
    {
        public static int ErrorsCount { get; set; }

        public static void ProcessSlide(SlideElement slide)
        {
            ShapeProcessor.ErrorsCount = 0;
            ShapeProcessor.Process(slide.Shapes);
            ErrorsCount += ShapeProcessor.ErrorsCount;
            slide.Slide.Save();
        }
    }
}
