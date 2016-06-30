using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace McKinsey.PowerPointGenerator.Processing
{
    public class ChartElementProcessor : IShapeElementProcessor
    {
        private ChartElement element;
        DataElement fullData;

        public void Process(ShapeElementBase shape)
        {
            element = shape as ChartElement;
            fullData = element.Data.Clone();
            element.Data = fullData.GetFragmentByIndexes(element.RowIndexes, element.ColumnIndexes);
            element.ProcessCommands(element.Data);
            if (element.Data == null)
            {
                return;
            }

            //get chart reference
            A.GraphicData graphicData = element.ChartFrame.Graphic.GraphicData;
            ChartReference chartReference = graphicData.FirstElement<ChartReference>();
            if (chartReference == null)
            {
                return;
            }

            //various chart structure elements
            ChartPart chartPart = element.Slide.Slide.SlidePart.GetPartById(chartReference.Id.Value) as ChartPart;
            Chart chart = chartPart.ChartSpace.FirstElement<Chart>();

            //get external data and update it
            DataElement dataToInsert = element.Data.Clone();
            foreach (var item in element.ChildShapes)
            {
                var childDataElement = fullData.GetFragmentByIndexes(item.RowIndexes, item.ColumnIndexes);
                element.ProcessCommands(childDataElement);
                dataToInsert.MergeWith(childDataElement);
            }
            ExternalData externalData = chartPart.ChartSpace.FirstElement<ExternalData>();
            EmbeddedPackagePart xlsPackagePart = chartPart.GetPartById(externalData.Id.Value) as EmbeddedPackagePart;
            Stream sourceStream = xlsPackagePart.GetStream();
            Stream outputStream = new MemoryStream();
            dataToInsert.TrimHiddenRowsAndColumns();
            List<ChartSeriesElement> newSeries = SpreadsheetProcessor.InsertData(dataToInsert, sourceStream, outputStream);
            outputStream.Seek(0, SeekOrigin.Begin);
            xlsPackagePart.FeedData(outputStream);


            ChartType type = ChartType.None;
            Tuple<int, int> dataRange = null;
            var charts = chart.PlotArea.Elements().ToList();

            OpenXmlElement mainChart = null;
            int chartIndex = 0;
            for (; chartIndex < charts.Count; chartIndex++)
            {
                GetChartTypeAndDataRange(ref type, ref dataRange, charts[chartIndex]);
                if (type != ChartType.None)
                {
                    mainChart = charts[chartIndex];
                    chartIndex++;
                    break;
                }
            }


            int seriesIndex = 0;
            foreach (ErrorBarCommand errorBarCommand in element.CommandsOf<ErrorBarCommand>())
            {
                ChartSeriesElement chartSeriesMinus = newSeries.FirstOrDefault(s => s.ColumnIndex == errorBarCommand.MinusIndex && !s.ColumnIndex.IsCore);
                ChartSeriesElement chartSeriesPlus = null;
                if (errorBarCommand.PlusIndex != null)
                {
                    chartSeriesPlus = newSeries.FirstOrDefault(s => s.ColumnIndex == errorBarCommand.PlusIndex && !s.ColumnIndex.IsCore);
                }
                if (seriesIndex < newSeries.Count)
                {
                    newSeries[seriesIndex].MinusErrorBar = chartSeriesMinus;
                    newSeries[seriesIndex].PlusErrorBar = chartSeriesPlus;
                }
                seriesIndex++;
            }
            foreach (ErrorBarCommand errorBarCommand in element.CommandsOf<ErrorBarCommand>())
            {
                ChartSeriesElement chartSeriesMinus = newSeries.FirstOrDefault(s => s.ColumnIndex == errorBarCommand.MinusIndex && !s.ColumnIndex.IsCore);
                ChartSeriesElement chartSeriesPlus = null;
                if (errorBarCommand.PlusIndex != null)
                {
                    chartSeriesPlus = newSeries.FirstOrDefault(s => s.ColumnIndex == errorBarCommand.PlusIndex && !s.ColumnIndex.IsCore);
                }
                if (chartSeriesMinus != null)
                {
                    newSeries.Remove(chartSeriesMinus);
                }
                if (chartSeriesPlus != null)
                {
                    newSeries.Remove(chartSeriesPlus);
                }
            }
            seriesIndex = 0;
            foreach (YCommand yCommand in element.CommandsOf<YCommand>())
            {
                ChartSeriesElement yChartSeries = newSeries.FirstOrDefault(s => s.ColumnIndex == yCommand.Index && !s.ColumnIndex.IsCore);
                if (seriesIndex < newSeries.Count)
                {
                    newSeries[seriesIndex].YValues = yChartSeries;
                }
                seriesIndex++;
            }
            foreach (YCommand yCommand in element.CommandsOf<YCommand>())
            {
                ChartSeriesElement yChartSeries = newSeries.FirstOrDefault(s => s.ColumnIndex == yCommand.Index && !s.ColumnIndex.IsCore);
                if (yChartSeries != null)
                {
                    newSeries.Remove(yChartSeries);
                }
            }


            switch (type)
            {
                case ChartType.Waterfall:
                case ChartType.Bar:
                    ReplaceBarChart(newSeries, mainChart as BarChart, element, element.IsWaterfall);
                    break;
                case ChartType.Scatter:
                    ReplaceScatterChart(newSeries, mainChart as ScatterChart, element);
                    break;
                case ChartType.Line:
                    ReplaceLineChart(newSeries, mainChart as LineChart, element);
                    break;
            }
            int childShapeIndex = 0;
            for (; chartIndex < charts.Count; chartIndex++)
            {
                var childChart = charts[chartIndex];
                if (element.ChildShapes.Count > childShapeIndex)
                {
                    GetChartTypeAndDataRange(ref type, ref dataRange, childChart);
                    if (type != ChartType.None)
                    {
                        ProcessChildShapes(element, element.ChildShapes[childShapeIndex], type, childChart, newSeries);
                        childShapeIndex++;
                    }                   
                }
            }
        }

        private void GetChartTypeAndDataRange(ref ChartType type, ref Tuple<int, int> dataRange, OpenXmlElement chart)
        {
            BarChart barchart = chart as BarChart;
            ScatterChart scatterChart = chart as ScatterChart;
            LineChart lineChart = chart as LineChart;
            PieChart pieChart = chart as PieChart;
            type = ChartType.None;
            if (barchart != null)
            {
                type = ChartType.Bar;
                if (element.IsFixed)
                {
                    dataRange = GetFixedDataRange<BarChartSeries>(barchart);
                    element.Data.TrimOrExpand(dataRange.Item1, dataRange.Item2);
                }
            }
            else if (scatterChart != null)
            {
                {
                    type = ChartType.Scatter;
                }
            }
            else if (lineChart != null)
            {
                type = ChartType.Line;
                if (element.IsFixed)
                {
                    dataRange = GetFixedDataRange<LineChartSeries>(barchart);
                    element.Data.TrimOrExpand(dataRange.Item1, dataRange.Item2);
                }
            }
            if (element.IsWaterfall)
            {
                type = ChartType.Waterfall;
            }
        }

        public void ProcessChildShapes(ShapeElementBase shape, ChildShapeElement childShape, ChartType type, OpenXmlElement childChart, List<ChartSeriesElement> newSeries)
        {
            childShape.Data = fullData.GetFragmentByIndexes(childShape.RowIndexes, childShape.ColumnIndexes);
            element.ProcessCommands(childShape.Data);
            switch (type)
            {
                case ChartType.Waterfall:
                case ChartType.Bar:
                    ReplaceBarChart(newSeries, childChart as BarChart, childShape, false);
                    break;
                case ChartType.Scatter:
                    ReplaceScatterChart(newSeries, childChart as ScatterChart, childShape);
                    break;
                case ChartType.Line:
                    ReplaceLineChart(newSeries, childChart as LineChart, childShape);
                    break;
            }
        }

        private void ReplaceLineChart(List<ChartSeriesElement> newSeries, LineChart chart, ShapeElementBase element)
        {
            int newSeriesCount = element.Data.Columns.Count;// newSeries.Count;
            var seriesList = SetChartSeries<LineChartSeries>(chart, newSeriesCount, true);
            int index = 0;
            for (int i = 0; i < newSeries.Count; i++)
            {
                ChartSeriesElement newSeriesItem = newSeries[i];
                if (element.RowIndexes.Any(idx => idx == newSeriesItem.ColumnIndex))
                {
                    var seriesItem = seriesList.ElementAt(index);
                    Column dataColumn = element.Data.Column(newSeriesItem.ColumnIndex);
                    var categoryAxisData = seriesItem.FirstElement<CategoryAxisData>();
                    FillCategoryAxisData(categoryAxisData, newSeriesItem, dataColumn);
                    SetSeriesText(seriesItem, newSeriesItem, dataColumn.GetHeader());
                    var values = seriesItem.FirstElement<Values>();
                    FillNumberReference(values.NumberReference, newSeriesItem, dataColumn);
                    FillSeriesDataPoints(seriesItem, dataColumn);
                    FillSeriesLabels(seriesItem, dataColumn);
                    SetPropertiesFromLegend(seriesItem, dataColumn);
                    var errorBars = seriesItem.FirstElement<ErrorBars>();
                    if (errorBars != null)
                    {
                        FillErrorBars(errorBars, newSeries[i].MinusErrorBar, newSeries[i].PlusErrorBar);
                    }
                    index++;
                }
            }
        }

        private void ReplaceScatterChart(List<ChartSeriesElement> newSeries, ScatterChart chart, ShapeElementBase element)
        {
            int newSeriesCount = element.Data.Columns.Count;// newSeries.Count;
            ChartSeriesElement yValuesSeries = null;
            var seriesList = SetChartSeries<ScatterChartSeries>(chart, newSeriesCount, false);
            int index = 0;

            for (int i = 0; i < newSeries.Count; i++)
            {
                if (newSeries[i].YValues != null)
                {
                    yValuesSeries = newSeries[i].YValues;
                }
                if (yValuesSeries == null)
                {
                    throw new Exception("At least on Y series required for scatter chart");
                }
                ChartSeriesElement newSeriesItem = newSeries.ElementAt(i);
                if (element.RowIndexes.Any(idx => idx == newSeriesItem.ColumnIndex))
                {
                    var seriesItem = seriesList.ElementAt(index);
                    Column dataColumn = element.Data.Column(newSeriesItem.ColumnIndex);

                    SetSeriesText(seriesItem, newSeriesItem, dataColumn.GetHeader());

                    var xvalues = seriesItem.FirstElement<XValues>();
                    FillNumberReference(xvalues.NumberReference, newSeriesItem, dataColumn);

                    var yvalues = seriesItem.FirstElement<YValues>();
                    FillNumberReference(yvalues.NumberReference, yValuesSeries, element.Data.Column(yValuesSeries.ColumnIndex));

                    var errorBars = seriesItem.FirstElement<ErrorBars>();
                    if (errorBars != null)
                    {
                        FillErrorBars(errorBars, newSeries[i].MinusErrorBar, newSeries[i].PlusErrorBar);
                    }
                    index++;
                }
            }
        }

        private void ReplaceBarChart(List<ChartSeriesElement> newSeries, BarChart chart, ShapeElementBase element, bool isWaterfall)
        {
            int newSeriesCount = element.Data.Columns.Count;// newSeries.Count;
            var seriesList = SetChartSeries<BarChartSeries>(chart, newSeriesCount, true);
            int index = 0;
            for (int i = 0; i < newSeries.Count; i++)
            {
                ChartSeriesElement newSeriesItem = newSeries[i];
                if (element.RowIndexes.Any(idx => idx == newSeriesItem.ColumnIndex))
                {
                    var seriesItem = seriesList.ElementAt(index);
                    Column dataColumn = element.Data.Column(newSeriesItem.ColumnIndex);
                    var categoryAxisData = seriesItem.FirstElement<CategoryAxisData>();
                    FillCategoryAxisData(categoryAxisData, newSeriesItem, dataColumn);
                    SetSeriesText(seriesItem, newSeriesItem, dataColumn.GetHeader());
                    var values = seriesItem.FirstElement<Values>();
                    FillNumberReference(values.NumberReference, newSeriesItem, dataColumn);
                    FillSeriesDataPoints(seriesItem, dataColumn);
                    FillSeriesLabels(seriesItem, dataColumn);
                    SetPropertiesFromLegend(seriesItem, dataColumn);
                    if (isWaterfall)
                    {
                        SetWaterfallStructure(seriesItem, dataColumn);
                    }

                    var errorBars = seriesItem.FirstElement<ErrorBars>();
                    if (errorBars != null)
                    {
                        FillErrorBars(errorBars, newSeries[i].MinusErrorBar, newSeries[i].PlusErrorBar);
                    }
                    index++;
                }
            }
        }

        private void FillErrorBars(ErrorBars errorBars, ChartSeriesElement minusErrorBarData, ChartSeriesElement plusErrorBarData)
        {
            var plus = errorBars.FirstElement<Plus>();
            if (plus != null && plusErrorBarData != null)
            {
                FillNumberReference(plus.NumberReference, plusErrorBarData, element.Data.Column(plusErrorBarData.ColumnIndex));
            }
            var minus = errorBars.FirstElement<Minus>();
            if (minus != null && minusErrorBarData != null)
            {
                FillNumberReference(minus.NumberReference, minusErrorBarData, element.Data.Column(minusErrorBarData.ColumnIndex));
            }
        }

        private void FillCategoryAxisData(CategoryAxisData categoryAxisData, ChartSeriesElement newSeriesItem, Column dataColumn)
        {
            int dataCount = dataColumn.Data.Count;
            UInt32Value pointCount = new UInt32Value((uint)dataCount);
            if (categoryAxisData.StringReference != null)
            {
                categoryAxisData.StringReference.StringCache.RemoveAllChildren<StringPoint>();
                categoryAxisData.StringReference.StringCache.PointCount.Val = pointCount;
                categoryAxisData.StringReference.Formula.Text = newSeriesItem.CategoryAxisDataAddress;
                for (int rowNo = 0; rowNo < element.Data.Rows.Count; rowNo++)
                {
                    StringPoint categoryAxisDataStringPoint = new StringPoint() { Index = (UInt32Value)((uint)rowNo) };
                    NumericValue categoryAxisDataNumericValue = new NumericValue() { Text = element.Data.Rows[rowNo].GetHeader() };
                    categoryAxisDataStringPoint.Append(categoryAxisDataNumericValue);
                    categoryAxisData.StringReference.StringCache.Append(categoryAxisDataStringPoint);
                }
            }
        }

        private void FillSeriesDataPoints(OpenXmlCompositeElement seriesItem, Column dataColumn)
        {
            var seriesChartShapeProperties = seriesItem.FirstElement<ChartShapeProperties>();

            for (int rowNo = 0; rowNo < dataColumn.Data.Count; rowNo++)
            {
                if (dataColumn.Data[rowNo] != null)
                {
                    DataPoint dp = seriesItem.Elements<DataPoint>().FirstOrDefault(p => p.Index != null && p.Index.Val != null && p.Index.Val.Value == rowNo);
                    if (dp == null)
                    {
                        var dataPoint = new DataPoint();
                        DocumentFormat.OpenXml.Drawing.Charts.Index index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = new UInt32Value((uint)rowNo) };
                        InvertIfNegative invertIfNegative = new InvertIfNegative() { Val = false };
                        Bubble3D bubble3D = new Bubble3D() { Val = false };
                        ChartShapeProperties chartShapeProperties = seriesChartShapeProperties == null ? new ChartShapeProperties() : (ChartShapeProperties)seriesChartShapeProperties.CloneNode(true);
                        dataPoint.Append(index);
                        dataPoint.Append(invertIfNegative);
                        dataPoint.Append(bubble3D);
                        dataPoint.Append(chartShapeProperties);
                        DataPoint lastDp = seriesItem.Elements<DataPoint>().LastOrDefault(p => p.Index != null && p.Index.Val != null && p.Index.Val.Value < rowNo);
                        if (lastDp != null)
                        {
                            seriesItem.InsertAfter(dataPoint, lastDp);
                        }
                        else
                        {
                            seriesItem.Append(dataPoint);
                        }
                    }
                }
            }
        }

        private void FillSeriesLabels(OpenXmlCompositeElement seriesItem, Column dataColumn)
        {
            // This seems to be redundant but not 100% sure.

            //var labels = seriesItem.FirstElement<DataLabels>();
            //if (labels == null)
            //{
            //    return;
            //}
            //TextProperties defaultTextProperties = labels.FirstElement<TextProperties>();
            //ChartShapeProperties defaultChartShapeProperties = labels.FirstElement<ChartShapeProperties>();
            //NumberingFormat defaultNumberingFormat = labels.FirstElement<NumberingFormat>();
            //ShowLegendKey defaultShowLegendKey = labels.FirstElement<ShowLegendKey>();
            //ShowValue defaultShowValue = labels.FirstElement<ShowValue>();
            //ShowCategoryName defaultShowCategoryName = labels.FirstElement<ShowCategoryName>();
            //ShowSeriesName defaultShowSeriesName = labels.FirstElement<ShowSeriesName>();
            //ShowPercent defaultShowPercent = labels.FirstElement<ShowPercent>();
            //ShowBubbleSize defaultShowBubbleSize = labels.FirstElement<ShowBubbleSize>();
            //ShowLeaderLines defaultShowLeaderLines = labels.FirstElement<ShowLeaderLines>();
            //DLblsExtension defaultDLblsExtension = labels.FirstElement<DLblsExtension>();
            //for (int rowNo = 0; rowNo < dataColumn.Data.Count; rowNo++)
            //{
            //    if (dataColumn.Data[rowNo] != null)
            //    {
            //        DataLabel dl = labels.Elements<DataLabel>().FirstOrDefault(l => l.Index != null && l.Index.Val != null && l.Index.Val.Value == rowNo);
            //        //if (dl != null)
            //        //{
            //        //    Delete delete = dl.FirstElement<Delete>();
            //        //    if (delete == null || !delete.Val)
            //        //    {
            //        //        var tp = dl.FirstElement<TextProperties>();
            //        //        if (tp != null)
            //        //        {
            //        //            textProperties = tp.CloneNode(true) as TextProperties;
            //        //        }
            //        //        labels.RemoveChild<DataLabel>(dl);
            //        //        dl = null;
            //        //    }
            //        //}
            //        if (dl == null)
            //        {
            //            var newDataLabel = new DataLabel();
            //            DocumentFormat.OpenXml.Drawing.Charts.Index index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = new UInt32Value((uint)rowNo) };
            //            newDataLabel.Index = index;
            //            if (defaultNumberingFormat != null)
            //            {
            //                newDataLabel.Append(defaultNumberingFormat.CloneNode(true));
            //            }
            //            if (defaultChartShapeProperties != null)
            //            {
            //                newDataLabel.Append(labels.FirstElement<ChartShapeProperties>().CloneNode(true));
            //            }
            //            if (defaultTextProperties != null)
            //            {
            //                newDataLabel.Append(defaultTextProperties.CloneNode(true));
            //            }
            //            newDataLabel.Append(defaultShowLegendKey.CloneNode(true));
            //            newDataLabel.Append(defaultShowValue.CloneNode(true));
            //            newDataLabel.Append(defaultShowCategoryName.CloneNode(true));
            //            newDataLabel.Append(defaultShowSeriesName.CloneNode(true));
            //            newDataLabel.Append(defaultShowPercent.CloneNode(true));
            //            newDataLabel.Append(defaultShowBubbleSize.CloneNode(true));
            //            if (defaultShowLeaderLines != null)
            //            {
            //                newDataLabel.Append(defaultShowLeaderLines.CloneNode(true));
            //            }
            //            if (defaultDLblsExtension != null)
            //            {
            //                newDataLabel.Append(defaultDLblsExtension.CloneNode(true));
            //            }
            //            DataLabel lastDataLabel = labels.Elements<DataLabel>().LastOrDefault(l => l.Index != null && l.Index.Val != null && l.Index.Val.Value < rowNo);
            //            if (lastDataLabel != null)
            //            {
            //                labels.InsertAfter(newDataLabel, lastDataLabel);
            //            }
            //            else
            //            {
            //                labels.InsertAt(newDataLabel, 0);
            //            }
            //        }
            //    }
            //}
        }

        private void SetWaterfallStructure(BarChartSeries seriesItem, Column dataColumn)
        {
            var dataPoints = seriesItem.Elements<DataPoint>();
            var labels = seriesItem.FirstElement<DataLabels>();
            bool isFirstBar = dataColumn.Data[0] != null;
            if (!isFirstBar)
            {
                int valueToHideInd = dataColumn.Data.IndexOf(dataColumn.Data.First(v => v != null));
                DataPoint dataPoint = dataPoints.FirstOrDefault(p => p.Index != null && p.Index.Val != null && p.Index.Val.Value == valueToHideInd);
                dataPoint.RemoveAllChildren<ChartShapeProperties>();
                ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
                chartShapeProperties.Append(new A.NoFill());
                dataPoint.Append(chartShapeProperties);
                DataLabel label = labels.Elements<DataLabel>().FirstOrDefault(p => p.Index != null && p.Index.Val != null && p.Index.Val.Value == valueToHideInd);
                label.FirstElement<ShowValue>().Val.Value = false;
            }
            labels.FirstElement<ShowValue>().Val.Value = false;
        }

        private void SetPropertiesFromLegend(OpenXmlCompositeElement seriesItem, Column dataColumn)
        {
            if (dataColumn.Legends.All(l => l == null))
            {
                return;
            }
            var dataPoints = seriesItem.Elements<DataPoint>();
            var labels = seriesItem.FirstElement<DataLabels>();
            for (int rowNo = 0; rowNo < dataColumn.Data.Count; rowNo++)
            {
                if (dataColumn.Legends[rowNo] != null && dataColumn.Data[rowNo] != null)
                {
                    ShapeElement legend = dataColumn.Legends[rowNo] as ShapeElement;
                    DataPoint dataPoint = dataPoints.FirstOrDefault(p => p.Index != null && p.Index.Val != null && p.Index.Val.Value == rowNo);
                    if (legend != null)
                    {
                        if (dataPoint != null)
                        {
                            dataPoint.RemoveAllChildren<ChartShapeProperties>();
                            A.SolidFill legendFill = legend.Element.GetFill();
                            A.Outline legendOutline = legend.Element.GetOutline();
                            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
                            chartShapeProperties.Append(legendFill);
                            chartShapeProperties.Append(legendOutline);
                            dataPoint.Append(chartShapeProperties);
                        }
                        if (labels != null)
                        {
                            DataLabel label = labels.Elements<DataLabel>().FirstOrDefault(p => p.Index != null && p.Index.Val != null && p.Index.Val.Value == rowNo);
                            if (label != null)
                            {
                                TextProperties labelTextProperties = label.FirstElement<TextProperties>();
                                if (labelTextProperties == null)
                                {
                                    labelTextProperties = labels.FirstElement<TextProperties>().CloneNode(true) as TextProperties;
                                    label.Append(labelTextProperties);
                                }
                                A.Paragraph labelParagraph = labelTextProperties.FirstElement<A.Paragraph>();
                                var legendRunProperties = legend.Element.GetRunProperties();
                                labelParagraph.ParagraphProperties.RemoveAllChildren<A.DefaultRunProperties>();
                                List<OpenXmlElement> list = new List<OpenXmlElement>();
                                foreach (var item in legendRunProperties.ChildElements)
                                {
                                    list.Add(item.CloneNode(true));
                                }
                                var newLabelRunProperties = new A.DefaultRunProperties(list);
                                labelParagraph.ParagraphProperties.Append(newLabelRunProperties);
                                var labelShapeProperties = label.FirstElement<ChartShapeProperties>();
                                if (labelShapeProperties != null && labelShapeProperties.FirstElement<A.NoFill>() == null)
                                {
                                    label.RemoveAllChildren<ChartShapeProperties>();
                                    A.SolidFill legendFill = legend.Element.GetFill();
                                    ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
                                    chartShapeProperties.Append(legendFill);
                                    label.Append(chartShapeProperties);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void FillNumberReference(NumberReference valuesNumberReference, ChartSeriesElement newSeriesItem, Column dataColumn)
        {
            int dataCount = dataColumn.Data.Count;
            UInt32Value pointCount = new UInt32Value((uint)dataCount);
            valuesNumberReference.Formula.Text = newSeriesItem.ValuesAddress;
            valuesNumberReference.NumberingCache.RemoveAllChildren<NumericPoint>();
            valuesNumberReference.NumberingCache.PointCount.Val = pointCount;
            for (int rowNo = 0; rowNo < dataColumn.Data.Count; rowNo++)
            {
                if (dataColumn.Data[rowNo] != null)
                {
                    var point = new NumericPoint() { Index = new UInt32Value((uint)rowNo) };
                    point.NumericValue = new NumericValue(dataColumn.Data[rowNo] == null ? "0" : dataColumn.Data[rowNo].ToString());
                    valuesNumberReference.NumberingCache.Append(point);
                }
            }
        }

        private void SetSeriesText(OpenXmlCompositeElement seriesItem, ChartSeriesElement newSeriesItem, string seriesHeader)
        {
            SeriesText seriesText = seriesItem.Elements<SeriesText>().First();
            seriesText.StringReference.Formula.Text = newSeriesItem.SeriesTextAddress;
            seriesText.StringReference.StringCache.RemoveAllChildren<StringPoint>();
            StringPoint stringReferencePoint = new StringPoint() { Index = (UInt32Value)((uint)0) };
            NumericValue stringReferenceNumericValue = new NumericValue() { Text = seriesHeader };
            stringReferencePoint.Append(stringReferenceNumericValue);
            seriesText.StringReference.StringCache.Append(stringReferencePoint);
        }

        private IEnumerable<T> SetChartSeries<T>(OpenXmlCompositeElement chart, int newSeriesCount, bool addNewIfRequired) where T : OpenXmlCompositeElement
        {
            var seriesList = chart.Elements<T>();
            int currentSeriesCount = seriesList.Count();
            if (currentSeriesCount > newSeriesCount)
            {
                //chart on template has more series than in the data, remove last x series
                int seriesToRemove = currentSeriesCount - newSeriesCount;
                for (int i = 0; i < seriesToRemove; i++)
                {
                    chart.RemoveChild<T>(seriesList.Last());
                }
                seriesList = chart.Elements<T>();
            }
            else
                if (addNewIfRequired && currentSeriesCount < newSeriesCount)
            {
                //chart on the template has fewer series so we need to add some by clonning the last one
                for (int i = 0; i < newSeriesCount - currentSeriesCount; i++)
                {
                    var lastSeries = chart.Elements<T>().Last();
                    var seriesClone = (T)lastSeries.CloneNode(true);
                    var lastSeriesIndex = lastSeries.FirstElement<DocumentFormat.OpenXml.Drawing.Charts.Index>();
                    var seriesCloneIndex = seriesClone.FirstElement<DocumentFormat.OpenXml.Drawing.Charts.Index>();
                    seriesCloneIndex.Val = new UInt32Value(lastSeriesIndex.Val.Value + 1);
                    var lastSeriesOrder = lastSeries.FirstElement<DocumentFormat.OpenXml.Drawing.Charts.Order>();
                    var seriesCloneOrder = seriesClone.FirstElement<DocumentFormat.OpenXml.Drawing.Charts.Order>();
                    seriesCloneOrder.Val = new UInt32Value(lastSeriesOrder.Val.Value + 1);
                    chart.InsertAfter(seriesClone, lastSeries);
                }
                seriesList = chart.Elements<T>();
            }
            return chart.Elements<T>();
        }


        private Tuple<int, int> GetFixedDataRange<T>(OpenXmlCompositeElement chart) where T : OpenXmlCompositeElement
        {
            int columns = chart.Elements<T>().Count();
            var values = chart.Elements<T>().First().FirstElement<Values>();
            int rows = (int)values.NumberReference.NumberingCache.PointCount.Val.Value;
            return new Tuple<int, int>(columns, rows);
        }
    }
}