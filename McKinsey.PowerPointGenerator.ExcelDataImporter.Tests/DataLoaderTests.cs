using System;
using McKinsey.PowerPointGenerator.ExcelDataImporter.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.ExcelDataImporter.Tests
{
    [TestClass]
    public class DataLoaderTests
    {
        [TestMethod]
        public void ImportLoadsFileAndGetsAllDataElements()
        {
            var s = Helpers.GetTemplateStreamFromResources(Resources.No_macros);
            var loader = new DataLoader();
            loader.Import(s);
            var list = loader.LoadData();
        }
    }
}
