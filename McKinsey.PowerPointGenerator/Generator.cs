using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.ExcelDataImporter;
using McKinsey.PowerPointGenerator.Processing;
using NLog;

namespace McKinsey.PowerPointGenerator
{
    public class Generator : IDisposable
    {
        public int ErrorsCount { get; set; }
        public Document Document { get; set; }
        public bool CancellationRequested { get; set; }
        public event EventHandler<ProgressEventArgs> ProgressUpdate;
        public List<string> Timing { get; set; }
        private DateTime lastPoint;

        public Generator()
        {
            //Instance = this;
            Timing = new List<string>();
            lastPoint = DateTime.Now;
            ErrorsCount = 0;
        }

        public void Run(Document document, Stream excelStream)
        {            
            ParameterHelpers.NotNull<Document>(document, "document");
            Document = document;

            if (!Document.IsLoaded)
            {
                throw new InvalidOperationException("You need to call Load() method of the document before running the generator.");
            }

            Logger logger = LogManager.GetLogger("Generator");
            try
            {
                logger.Debug("------------------------------------------------------------------------------------------------------------------------------------");
                logger.Debug("Step 1: Loading data");
                AddTiming(2, "Starting Excel loading");
                IList<DataElement> data = ImportData(excelStream, logger);

                logger.Debug("------------------------------------------------------------------------------------------------------------------------------------");
                logger.Debug("Step 2: Loading and parsing slides");
                GetSlides(logger);
                AddTiming(3, "Slides parsed");

                logger.Debug("------------------------------------------------------------------------------------------------------------------------------------");
                logger.Debug("Step 3: Processing data");
                ProcessData(logger, data);
                AddTiming(4, "Data processed");

                logger.Debug("------------------------------------------------------------------------------------------------------------------------------------");
                logger.Debug("Step 4: Generating content");
                AddTiming(5, "Starting generation");
                ProcessSlides(logger, data);

                logger.Debug("------------------------------------------------------------------------------------------------------------------------------------");
                logger.Debug("Done");
                ReportProgress("Finishing...", 1, 1);
                Document.SaveAndClose();
                AddTiming(6, "Output saved");
                ReportProgress("Done!", 0, 0, true);
            }
            catch (Exception ex)
            {
                logger.Debug("Error, processing stopped");
                logger.Debug(ex.ToString());
                throw;
            }
        }

        ~Generator()
        {
            Dispose(false);
        }

        public void AddTiming(int step, string stepName)
        {
            int tens = (int)Math.Floor((double)step / 10d);
            string msg = new string(' ', tens * 2) + step.ToString() + ": " + stepName;
            DateTime now = DateTime.Now;
            Timing.Add(string.Format("{0} {1}", msg, now - lastPoint));
            lastPoint = now;
        }

        private IList<DataElement> ImportData(Stream excelStream, Logger logger)
        {
            var loader = new DataLoader();
            loader.Import(excelStream, (p, t) => { ReportProgress("Loading data...", p, t); });
            AddTiming(20, "Excel file loaded");
            var data = loader.LoadData((p, t) => { ReportProgress("Loading data...", p, t); });
            AddTiming(20, "All named ranges loaded");
            return data;
        }

        private void GetSlides(Logger logger)
        {
            //get all slides
            Document.GetSlides();
            ReportProgress("Loading and parsing slides...", 0, Document.Slides.Count);
            logger.Debug("Loaded {0} slides", Document.Slides.Count);

            //discover shapes and their commands
            for (int slideIndex = 0; slideIndex < Document.Slides.Count; slideIndex++)
            {
                try
                {
                    ReportProgress("Loading and parsing slides...", slideIndex + 1, Document.Slides.Count);
                    Document.Slides[slideIndex].DiscoverShapes();
                    Document.Slides[slideIndex].DiscoverCommands();
                    logger.Debug("Slide {0}: found {1} objects to replace", Document.Slides[slideIndex].Number, Document.Slides[slideIndex].Shapes.Count);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Error while loading slide {0}. See inner exception for details.", Document.Slides[slideIndex].Number), ex);
                }
            }
        }

        private void ProcessData(Logger logger, IList<DataElement> data)
        {
            //apply data placeholders inside data elements itself
            ReportProgress("Processing data...", 0, data.Count);
            for (int dataIndex = 0; dataIndex < data.Count; dataIndex++)
            {
                try
                {
                    ReportProgress("Processing data...", dataIndex + 1, data.Count);
                    DataElementProcessor.ProcessDataElement(data[dataIndex], data);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Error while processing data element {0}. See inner exception for details.", data[dataIndex].Name), ex);
                }
            }
        }

        private void ProcessSlides(Logger logger, IList<DataElement> data)
        {
            ReportProgress("Generating content...", 0, data.Count);
            SlideProcessor.ErrorsCount = 0;
            for (int slideIndex = 0; slideIndex < Document.Slides.Count; slideIndex++)
            {
                ReportProgress("Generating content...", slideIndex + 1, Document.Slides.Count);
                Document.Slides[slideIndex].FindDataElements(data);
                SlideProcessor.ProcessSlide(Document.Slides[slideIndex]);
                AddTiming(51, "Slide " + slideIndex + " processed");
            }
            ErrorsCount = SlideProcessor.ErrorsCount;
        }

        private void ReportProgress(string step, int done, int total, bool completed = false)
        {
            if (ProgressUpdate != null)
            {
                ProgressEventArgs args = new ProgressEventArgs { IsFinished = completed, Total = total, StepName = step };
                args.Progress = (double)done / (double)total;
                ProgressUpdate(this, args);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
                if (Document != null)
                {
                    Document.Dispose();
                    Document = null;
                }
            }
        }
    }

    public class ProgressEventArgs : EventArgs
    {
        public double Progress { get; set; }
        public int Total { get; set; }
        public bool IsFinished { get; set; }
        public bool Cancel { get; set; }
        public string StepName { get; set; }
    }
}
