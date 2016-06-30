using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace McKinsey.PowerPointGenerator.ExcelDataImporter.Tests
{
    internal static class Helpers
    {
        internal static Stream GetTemplateStreamFromResources(byte[] resourceData)
        {
            string path = Path.GetTempFileName();
            File.WriteAllBytes(path, resourceData);
            Stream s = File.Open(path, FileMode.Open, FileAccess.ReadWrite);
            return s;
        }
    }
}