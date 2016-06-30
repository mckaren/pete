using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;

namespace McKinsey.PowerPointGenerator.Tests
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

        internal static DataElement CreateTestDataElement()
        {
            DataElement da = new DataElement();
            da.Name = "test";
            Row row1 = new Row { Header = "row 1", MappedHeader = "Totals" };
            Row row2 = new Row { Header = "row 2", IsHidden = true };
            Column column1 = new Column { Header = "Column 1" };
            Column column2 = new Column { Header = "Column 2", MappedHeader = "IBM" };
            Column column3 = new Column { Header = "Column 3", IsHidden = true };
            //row1.Data.AddRange(new object[] { 1.0, "client 1", false });
            //row2.Data.AddRange(new object[] { 5.05, "client 2", true });
            column1.Data.AddRange(new object[] { 1.0, 5.05 });
            column2.Data.AddRange(new object[] { "client 1", "client 2" });
            column3.Data.AddRange(new object[] { false, true });
            da.Rows.Add(row1);
            da.Rows.Add(row2);
            da.Columns.Add(column1);
            da.Columns.Add(column2);
            da.Columns.Add(column3);
            da.Normalize();
            da.HasColumnHeaders = true;
            da.HasRowHeaders = true;
            return da;
        }

        internal static DataElement CreateTestDataElementWithTwoNumericColumns()
        {
            DataElement da = new DataElement();
            da.Name = "test";
            Row row1 = new Row { Header = "row 1" };
            Row row2 = new Row { Header = "row 2" };
            Column column1 = new Column { Header = "Column 1" };
            Column column2 = new Column { Header = "Column 2" };
            Column column3 = new Column { Header = "Column 3" };
            Column column4 = new Column { Header = "Column 4" };
            //row1.Data.AddRange(new object[] { 1.0, "client 1", false });
            //row2.Data.AddRange(new object[] { 5.05, "client 2", true });
            column1.Data.AddRange(new object[] { 1.0, 5.05 });
            column2.Data.AddRange(new object[] { "client 1", "client 2" });
            column3.Data.AddRange(new object[] { false, true });
            column4.Data.AddRange(new object[] { 2.5, 4.05 });
            da.Rows.Add(row1);
            da.Rows.Add(row2);
            da.Columns.Add(column1);
            da.Columns.Add(column2);
            da.Columns.Add(column3);
            da.Columns.Add(column4);
            da.Normalize();
            da.HasColumnHeaders = true;
            da.HasRowHeaders = true;
            return da;
        }

        internal static DataElement CreateSingleValueElement(string name, object value)
        {
            DataElement da = new DataElement { Name = name };
            da.Rows.Add(new Row());
            da.Columns.Add(new Column());
            da.Rows[0].Data.Add(value);
            da.Normalize();
            return da;
        }

        internal static DataElement CreateTwoValueElement(string name, object value1, object value2)
        {
            DataElement da = new DataElement { Name = name };
            da.Rows.Add(new Row());
            da.Columns.Add(new Column());
            da.Columns.Add(new Column());
            da.Rows[0].Data.Add(value1);
            da.Rows[0].Data.Add(value2);
            da.Normalize();
            return da;
        }

        internal static string CreateJsonForDataElementWithDataInRows()
        {
            return 
@"{
  ""name"": ""test_element"",
  ""columns"": [ 
                { ""header"": ""Column 1"" },
                { ""header"": ""Column 2"" },
                { ""header"": ""Column 3"" },
              ],
  ""rows"": [
    {
      ""header"": ""row 1"",
      ""data"": [ 1.0, ""client 1"", false]
    },
    {
      ""header"": ""row 2"",
      ""data"": [ 5.05, ""client 2"", true]
    }
  ]
}";
        }

        internal static string CreateJsonForDataElementWithDataInColumns()
        {
            return
@"{
  ""name"": ""test_element"",
  ""columns"": [ 
    { 
      ""header"": ""Column 1"",
      ""data"": [ 1.0, 5.05]
    },
    {
      ""header"": ""Column 2"",
      ""data"": [ ""client 1"", ""client 2""]
    },
    {
      ""header"": ""Column 3"", 
      ""data"": [ false, true]
    },
  ],
  ""rows"": [
    { ""header"": ""row 1""},
    { ""header"": ""row 2"" }
  ]
}";
        }

        internal static string CreateJsonForDataElementForChartReplaceTest()
        {
            return
@"{
  ""name"": ""SpendByRevenuesC"",
  ""columns"": [ 
    { 
      ""header"": ""Column 1"",
      ""data"": [ 0.3, 0.3, 0.2, 0.2]
    },
    {
      ""header"": ""Column 2"",
      ""data"": [ 0.2, 0.2, 0.4, 0.2]
    },
    {
      ""header"": ""Column 3"", 
      ""data"": [ 0.3, 0.4, 0.1, 0.2]
    },
    {
      ""header"": ""Column 4"", 
      ""data"": [ 0.7, 0.1, 0.1, 0.1]
    }
  ],
  ""rows"": [
    { ""header"": ""row 1"" },
    { ""header"": ""row 2"" },
    { ""header"": ""row 3"" },
    { ""header"": ""row 4"" }
  ]
}";
        }
    }

    public class ShapeElementBaseTest : ShapeElementBase
    {
        public override string TypeName
        {
            get { return ""; }
        }

        public override IEnumerable<Command> PreprocessSwitchCommands(IEnumerable<Command> discoveredCommands)
        {
            return discoveredCommands;
        }

        public static ShapeElement Create(string name, object element, SlideElement slide)
        {
            ShapeElement shape = new ShapeElement();
            if (!shape.Parse(name, slide))
            {
                return null;
            }
            return shape;
        }

        public static ShapeElement Create(string name, SlideElement slide)
        {
            ShapeElement shape = new ShapeElement();
            if (!shape.Parse(name, slide))
            {
                return null;
            }
            return shape;
        }

        public static ShapeElement Create(string name)
        {
            ShapeElement shape = new ShapeElement();
            shape.Name = name;
            return shape;
        }

        public static ShapeElement Create()
        {
            ShapeElement shape = new ShapeElement();
            return shape;
        }
    }
}
