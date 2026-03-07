using ExcelConsolidator.Models;
using ExcelConsolidator.Services;
using INIParser;
using Microsoft.Win32;
using System.ComponentModel;
using System.Windows;

namespace ExcelConsolidator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler? PropertyChanged;

        private string? _folderPath;
        private string? _templateFilePath;
        private string? _outputFilePath;
        public bool SubmitIsEnabled { get; set; } = false;

        public string FolderPathDisplay
        {
            get
            {                
                return _folderPath ?? "No folder path selected.";
            }
        }

        public string TemplateFilePath
        {
            get {
                return _templateFilePath ?? "No template selected.";
            }
        }

        public string OutputFilePath {
            get {
                return _outputFilePath ?? "No output file selected.";
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CompleteExtraction();
        }

        private static void CompleteExtraction()
        {
            var iniFile = new IniFile();
            iniFile.LoadFile(@"C:\Users\jvand\source\repos\ExcelConsolidator\config.ini");

            // Retrieve the values using the Section and Key names
            string folderPath = iniFile["Paths", "SourceFolder"];
            string templateFilePath = iniFile["Paths", "TemplateFile"];
            string outputFilePath = iniFile["Paths", "OutputFile"];

            //string folderPath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\Directory Of Files";
            //string templateFilePath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\SampleTemplate.xlsx";
            //string outputFilePath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\OutputFile.xlsx";

            var extractionTempalte = new ExtractionTemplate();

            ExportTemplate template = extractionTempalte.GetTemplateItems(templateFilePath);
            ExcelExtraction extraction = new ExcelExtraction(template);
            ExportRowsCollection rowsCollection = extraction.ExtractDataFromDirectory(folderPath);

            ExcelExport excelExport = new ExcelExport(outputFilePath, rowsCollection, template);
        }

        private void SelectSourceFolder_Click(object sender, RoutedEventArgs e)
        {
            // 1. Initialize the modern folder picker
            OpenFolderDialog dialog = new OpenFolderDialog
            {
                Title = "Select the Excel Folder",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            // 2. Show the dialog and check if the user clicked "OK"
            if (dialog.ShowDialog() == true)
            {
                // 3. Update your private variable
                _folderPath = dialog.FolderName;

                // 4. IMPORTANT: Tell the UI that 'FolderPathDisplay' has changed
                // This is what fixes the "blank display" issue when the path updates
                OnPropertyChanged(nameof(FolderPathDisplay));

                // Optional: Update button states
                CheckIfSubmitShouldBeEnabled();
            }

        }

        private void SelectTemplateFile_Click(object sender, RoutedEventArgs e) {

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Workbooks (*.xlsx)|*.xlsx|All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                _templateFilePath = openFileDialog.FileName;
            }
            OnPropertyChanged(nameof(TemplateFilePath));
            CheckIfSubmitShouldBeEnabled();
        }

        private void SelectOutputFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Excel Workbooks (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Title = "Select where to save the consolidated file";

            if (saveFileDialog.ShowDialog() == true)
            {
                _outputFilePath = saveFileDialog.FileName;

                OnPropertyChanged(nameof(OutputFilePath));
            }
            CheckIfSubmitShouldBeEnabled();
        }

        private void CheckIfSubmitShouldBeEnabled() {
            SubmitIsEnabled = (_folderPath != null && _templateFilePath != null && _outputFilePath != null);
            //MessageBox.Show(SubmitIsEnabled.ToString());
            OnPropertyChanged(nameof(SubmitIsEnabled));
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
