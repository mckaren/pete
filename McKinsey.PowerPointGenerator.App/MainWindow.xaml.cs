using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using McKinsey.PowerPointGenerator.App.Properties;
using McKinsey.PowerPointGenerator.ExcelDataImporter;
using Newtonsoft.Json;

namespace McKinsey.PowerPointGenerator.App
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        string defaultSaveExtension;

        public MainWindow()
        {
            InitializeComponent();
            Accent accent = ThemeManager.Accents.First(a => a.Name.Equals("orange", StringComparison.OrdinalIgnoreCase));
            AppTheme theme = ThemeManager.AppThemes.First(a => a.Name.Equals("BaseLight", StringComparison.OrdinalIgnoreCase));
            ThemeManager.ChangeAppStyle(App.Current, accent, theme);
            Loaded += WindowLoaded;
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            MainViewModel vieModel = DataContext as MainViewModel;
            vieModel.Completed -= GenerationCompleted;
            vieModel.Completed += GenerationCompleted;
            vieModel.Failed -= GenerationFailed;
            vieModel.Failed += GenerationFailed;
        }

        void GenerationFailed(object sender, EventArgs e)
        {
            Dispatcher.Invoke((Action)(() =>
            {
                MainViewModel vieModel = DataContext as MainViewModel;
                this.ShowMessageAsync("Failed", "Detailed error information has been saved to log file\r\n" + vieModel.OutputPath + ".txt", MessageDialogStyle.Affirmative, new MetroDialogSettings { ColorScheme = MetroDialogColorScheme.Accented });
            }));
        }

        void GenerationCompleted(object sender, EventArgs e)
        {
            Dispatcher.Invoke((Action)(() =>
                {
                    MainViewModel vieModel = DataContext as MainViewModel;
                    string message = "Output file has been saved to " + vieModel.OutputPath + "\r\nLog information has been saved to " + vieModel.OutputPath + ".txt";
                    string title = "Completed";
                    MetroDialogSettings newMetroDialogSettings = new MetroDialogSettings { ColorScheme = MetroDialogColorScheme.Theme };
                    if (vieModel.ErrorsCount > 0)
                    {
                        title = "Completed with errors";
                        message = "There were " + vieModel.ErrorsCount + " errors during generation.\r\nPlease check the log for details.\r\n\r\n" + message;
                        newMetroDialogSettings = new MetroDialogSettings { ColorScheme = MetroDialogColorScheme.Accented };
                    }
                    this.ShowMessageAsync(title, message, MessageDialogStyle.Affirmative, newMetroDialogSettings);
                }));
        }

        public string OpenSaveDialog()
        {
            System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog();
            dialog.Title = "Save Generated Document As...";
            dialog.DefaultExt = defaultSaveExtension;
            dialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx|PowerPoint 97-2003 Presentation (*.ppt)|*ppt|PowerPoint Macro-Enabled presentation (*.pptm)|*.pptm";
            switch (defaultSaveExtension)
            {
                case ".pptx":
                    dialog.FilterIndex = 1;
                    break;
                case ".ppt":
                    dialog.FilterIndex = 2;
                    break;
                case ".pptm":
                    dialog.FilterIndex = 3;
                    break;
                default:
                    dialog.FilterIndex = 1;
                    break;
            }
            if (string.IsNullOrEmpty(Settings.Default.LastSaveFolder))
            {
                Settings.Default.LastSaveFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            dialog.InitialDirectory = Settings.Default.LastSaveFolder;
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                Settings.Default.LastSaveFolder = System.IO.Path.GetDirectoryName(dialog.FileName);
                Settings.Default.Save();
                return dialog.FileName;
            }
            return null;
        }

        public string OpenPptOpenDialog()
        {
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
            dialog.Title = "Load Template File...";
            dialog.DefaultExt = ".pptx";
            dialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx|PowerPoint 97-2003 Presentation (*.ppt)|*ppt|PowerPoint Macro-Enabled presentation (*.pptm)|*.pptm";
            if (string.IsNullOrEmpty(Settings.Default.LastTemplateFolder))
            {
                Settings.Default.LastTemplateFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            dialog.InitialDirectory = Settings.Default.LastTemplateFolder;
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                Settings.Default.LastTemplateFolder = System.IO.Path.GetDirectoryName(dialog.FileName);
                Settings.Default.Save();
                defaultSaveExtension = System.IO.Path.GetExtension(dialog.FileName);
                return dialog.FileName;
            }
            return null;
        }

        public string OpenXlsOpenDialog()
        {
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
            dialog.Title = "Load Data File...";
            dialog.DefaultExt = ".xlsx";
            dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            if (string.IsNullOrEmpty(Settings.Default.LastExcelFolder))
            {
                Settings.Default.LastExcelFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            dialog.InitialDirectory = Settings.Default.LastExcelFolder;
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                Settings.Default.LastExcelFolder = System.IO.Path.GetDirectoryName(dialog.FileName);
                Settings.Default.Save();
                return dialog.FileName;
            }
            return null;
        }
    }
}
