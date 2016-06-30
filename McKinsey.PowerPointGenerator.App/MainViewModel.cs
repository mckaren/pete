using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using DevExpress.Mvvm;
using DevExpress.Mvvm.DataAnnotations;
using DevExpress.Mvvm.UI;
using MahApps.Metro.Controls.Dialogs;
using McKinsey.PowerPointGenerator.ExcelDataImporter;
using Newtonsoft.Json;
using McKinsey.PowerPointGenerator;
using NLog.Config;
using NLog.Targets;
using NLog;

namespace McKinsey.PowerPointGenerator.App
{
    public class MainViewModel
    {
        public int ErrorsCount { get; set; }
        private FileStream excelStream;
        private FileStream templateStream;
        public virtual string DataFilePath { get; set; }
        public virtual string TemplateFilePath { get; set; }
        public string OutputPath { get; set; }
        public event EventHandler Completed;
        public event EventHandler Failed;

        [ServiceProperty(SearchMode = ServiceSearchMode.LocalOnly, Key = "RootView")]
        public virtual ICurrentWindowService RootWindow
        {
            get { return null; }
        }

        public MainViewModel()
        {
            DataFilePath = @"C:\Users\Karen Jones\Desktop\WIP\04 PeTE\QA\TBM Index - Analysis model v0.03.xlsm";
            TemplateFilePath = @"C:\Users\Karen Jones\Desktop\WIP\04 PeTE\QA\LegendTest.pptm";
            OutputPath = @"C:\Users\Karen Jones\Desktop\WIP\04 PeTE\QA\test.pptm";
        }

        private bool CanExecuteGenerateCommand()
        {
            return !(string.IsNullOrEmpty(DataFilePath) && string.IsNullOrEmpty(TemplateFilePath));
        }

        private ProgressDialogController controller;
        private Generator generator;

        [Command(Name = "BrowseExcelCommand")]
        public void BrowseExcel()
        {
            MainWindow mainWindow = ((CurrentWindowService)RootWindow).Window as MainWindow;
            DataFilePath = mainWindow.OpenXlsOpenDialog();
        }

        [Command(Name = "BrowsePowerPointCommand")]
        public void BrowsePowerPoint()
        {
            MainWindow mainWindow = ((CurrentWindowService)RootWindow).Window as MainWindow;
            TemplateFilePath = mainWindow.OpenPptOpenDialog();
        }

        [Command(Name = "GenerateCommand")]
        public async void Generate()
        {
            MainWindow mainWindow = ((CurrentWindowService)RootWindow).Window as MainWindow;
            //OutputPath = mainWindow.OpenSaveDialog();
            if (string.IsNullOrEmpty(OutputPath))
            {
                return;
            }
            generator = new Generator();
            generator.AddTiming(0, "Initiated");
            generator.ProgressUpdate += GeneratorProgressUpdate;
            CreateLogger(OutputPath);
            Logger logger = LogManager.GetLogger("Generator");
            logger.Debug("Starting generation");
            logger.Debug("Data file: {0}", DataFilePath);
            logger.Debug("Template file: {0}", TemplateFilePath);
            logger.Debug("Output file: {0}", OutputPath);
            controller = await mainWindow.ShowProgressAsync("Please wait...", "Getting ready...");
            controller.SetCancelable(true);
            var task = Task.Run(() =>
            {
                Document doc = new Document();
                File.Copy(TemplateFilePath, OutputPath, true);
                templateStream = File.Open(OutputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                excelStream = File.Open(DataFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                doc.Load(templateStream);
                generator.AddTiming(1, "Template loaded");
                generator.Run(doc, excelStream);
            })
            .ContinueWith(async (t) =>
            {
                if (t.IsFaulted && t.Exception != null)
                {
                    generator.Dispose();
                    excelStream.Dispose();
                    templateStream.Dispose();
                    await controller.CloseAsync();
                    if (Failed != null)
                    {
                        Failed(this, EventArgs.Empty);
                    }
                }
            });
        }

        private void CreateLogger(string OutputPath)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget();
            config.AddTarget("file", fileTarget);
            fileTarget.FileName = OutputPath + ".txt";
            fileTarget.Layout = "${message}";
            fileTarget.AutoFlush = true;
            fileTarget.DeleteOldFileOnStartup = true;
            var rule = new LoggingRule("*", LogLevel.Debug, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
            LogManager.ThrowExceptions = true;
        }

        private async void GeneratorProgressUpdate(object sender, ProgressEventArgs e)
        {
            if (e.IsFinished)
            {
                DumpTimigs(generator.Timing);
                ErrorsCount = generator.ErrorsCount;
                generator.Dispose();
                excelStream.Dispose();
                templateStream.Dispose();
                await controller.CloseAsync();
                if (Completed != null)
                {
                    Completed(this, EventArgs.Empty);
                }
            }
            else
            {
                if (controller.IsCanceled)
                {
                    generator.CancellationRequested = true;
                }
                controller.SetProgress(e.Progress);
                controller.SetMessage(e.StepName);
            }
        }

        private void DumpTimigs(List<string> timmings)
        {
            Logger logger = LogManager.GetLogger("Generator");
            logger.Debug("------------------------------------------------------------------------------------------------------------------------------------");
            logger.Debug("Timings:");
            foreach (string msg in timmings)
            {
                logger.Debug(msg);
            }
        }

        public void ExecuteCancelCommand()
        {
        }
    }
}
