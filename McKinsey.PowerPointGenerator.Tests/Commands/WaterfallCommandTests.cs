using System;
using System.Globalization;
using System.Linq;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class WaterfallCommandTests
    {
        [TestMethod]
        public void ApplyToDataBuildsWaterfallStructure()
        {
            WaterfallCommand cmd = new WaterfallCommand();
            DataElement data = GetData();
            cmd.ApplyToData(data);
            Assert.AreEqual("77.24887193", data.Columns[0].Data[5].ToString());
            Assert.AreEqual("77.24887193", data.Columns[0].Data[6].ToString());
            Assert.AreEqual("77.24887193", data.Columns[1].Data[4].ToString());
            Assert.AreEqual("104.9621822", data.Columns[1].Data[5].ToString());
            Assert.AreEqual("77.24887193", data.Columns[2].Data[3].ToString());
            Assert.AreEqual("20.10254469", data.Columns[2].Data[4].ToString());
            Assert.AreEqual("145.3517002", data.Columns[3].Data[2].ToString());
            Assert.AreEqual("68.10282827", data.Columns[3].Data[3].ToString());
            Assert.AreEqual("174.16395352", data.Columns[4].Data[1].ToString());
            Assert.AreEqual("28.81225332", data.Columns[4].Data[2].ToString());
            Assert.AreEqual("194.80999056", data.Columns[5].Data[0].ToString());
            Assert.AreEqual("20.64603704", data.Columns[5].Data[1].ToString());
        }

        [TestMethod]
        public void ApplyDataPreservesLegendsButOnlyInValuePlaces()
        {
            WaterfallCommand cmd = new WaterfallCommand();
            DataElement data = GetData();
            data.Columns[0].Legends[4] = "negative";
            data.Columns[0].Legends[5] = "negative";
            data.Rows[4].Legends[0] = "negative";
            data.Rows[5].Legends[0] = "negative";
            cmd.ApplyToData(data);

            Assert.IsNull(data.Columns[0].Legends[5]);
            Assert.IsNull(data.Columns[0].Legends[6]);
            Assert.IsNull(data.Columns[1].Legends[4]);
            Assert.AreEqual("negative", data.Columns[1].Legends[5]);
            Assert.IsNull(data.Columns[2].Legends[3]);
            Assert.AreEqual("negative", data.Columns[2].Legends[4]);
            Assert.IsNull(data.Columns[3].Legends[2]);
            Assert.IsNull(data.Columns[3].Legends[3]);
            Assert.IsNull(data.Columns[4].Legends[1]);
            Assert.IsNull(data.Columns[4].Legends[2]);
            Assert.IsNull(data.Columns[5].Legends[0]);
            Assert.IsNull(data.Columns[5].Legends[1]);
        }

        private static DataElement GetData()
        {
            DataElement data = new DataElement();
            data.Rows.Add(new Row { Header = "IT" });
            data.Rows.Add(new Row { Header = "AD" });
            data.Rows.Add(new Row { Header = "AM" });
            data.Rows.Add(new Row { Header = "Servers" });
            data.Rows.Add(new Row { Header = "EUS" });
            data.Rows.Add(new Row { Header = "NTS" });
            data.Rows.Add(new Row { Header = "Mgmt" });
            data.Columns.Add(new Column());
            data.Columns[0].Data.AddRange(new object[] { "e", 20.64603704, 28.81225332, 68.10282827, -20.10254469, -104.9621822, 77.24887193 });
            data.Normalize();
            return data;
        }
    }
}
