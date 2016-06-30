using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.Practices.Unity;
using NLog;

namespace McKinsey.PowerPointGenerator.Processing
{
    public class ShapeProcessor
    {
        public static int ErrorsCount { get; set; }
        private static IUnityContainer container = new UnityContainer();
        private static Logger logger;

        static ShapeProcessor()
        {
            container.RegisterType<IShapeElementProcessor, TextElementProcessor>(TextElement.TypeNameId);
            container.RegisterType<IShapeElementProcessor, ChartElementProcessor>(ChartElement.TypeNameId);
            container.RegisterType<IShapeElementProcessor, TableElementProcessor>(TableElement.TypeNameId);
            container.RegisterType<IShapeElementProcessor, ShapeElementProcessor>(ShapeElement.TypeNameId);
        }

        public static void Process(List<ShapeElementBase> shapes)
        {
            logger = LogManager.GetLogger("Generator");
            foreach (ShapeElementBase shape in shapes)
            {
                if (container.IsRegistered<IShapeElementProcessor>(shape.TypeName))
                {
                    try
                    {
                        IShapeElementProcessor processor = container.Resolve<IShapeElementProcessor>(shape.TypeName);
                        processor.Process(shape);
                    }
                    catch (Exception ex)
                    {
                        logger.Debug(string.Format("Error while processing shape {0} on slide {1}. See below for details.", shape.FullName, shape.Slide.Number));
                        logger.Debug(ex.ToString());
                        ErrorsCount++;
                    }
                }
            }
        }
    }
}
