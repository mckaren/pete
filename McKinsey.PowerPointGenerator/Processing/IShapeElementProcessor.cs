using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;

namespace McKinsey.PowerPointGenerator.Processing
{
    public interface IShapeElementProcessor
    {
        void Process(ShapeElementBase shape);
    }
}
