using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;
using ClassToExcel;
using ClassToExcel.Sample.Data;
using Microsoft.Win32;
using Run = System.Windows.Documents.Run;

namespace WpfExample
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DoReadWorkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = FilePath.Text;
                if (File.Exists(filePath))
                {
                    LogMessage($"Loading data from {filePath}");

                    // Passing in LogServiceMessage in optional.  I'm using it here to help debug in issues while reading a file.
                    var readerService = new ClassToExcelReaderService<Person>(LogServiceMessage);
                    var worksheetName = "People";
                    List<Person> people = readerService.ReadWorksheet(filePath, worksheetName, true);
                    if (people.Count > 0)
                    {
                        foreach (var person in people)
                        {
                            LogMessage(person.ToString());
                        }
                    }
                    else LogMessage($"No data found on the {worksheetName} work sheet!");
                }
                else LogMessage("The file does not exists: " + filePath);
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }
        

        private void DoCreateTestFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new SaveFileDialog();
                dialog.FileName = FilePath.Text;
                dialog.InitialDirectory = Path.GetDirectoryName(FilePath.Text);
                if (dialog.ShowDialog(this) == true)
                {
                    List<Person> people = PersonCreator.Work(35);
                    List<Chicken> chickens = ChickenCreator.Work(30);

                    // Passing in LogServiceMessage in optional.  I'm using it here to help debug any issues while writing a file.
                    using (var writerService = new ClassToExcelWriterService(LogServiceMessage))
                    {
                        writerService.AddWorksheet(people, "People", true);
                        writerService.AddWorksheet(chickens, "Chickens", true);
                        writerService.SaveToFile(dialog.FileName);
                    }

                    LogMessage("File saved");
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>Reads everything as a string with no coversion at all.  Keep in mind that Excel does NOT require 
        /// that BLANK cells be written so rows will not always have every column in their data!</summary>
        private void DoRawReadWorkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("I'm using RAW read service (ClassToExcelRawReaderService) to get data and then dumping it directly out in the log.");
                LogMessage("Note that blank rows are skipped by OpenXml.");

                string filePath = FilePath.Text;
                if (File.Exists(filePath))
                {
                    LogMessage($"Loading data from {filePath}");

                    // Passing in LogServiceMessage in optional.  I'm using it here to help debug in issues while reading a file.
                    var rawReaderService = new ClassToExcelRawReaderService(LogServiceMessage);

                    // Get tab names
                    List<string> worksheetNames = string.IsNullOrWhiteSpace(WorksheetNames.Text)
                        ? new List<string>()
                        : WorksheetNames.Text.Split(',').ToList();

                    if (worksheetNames.Count > 0)
                    {
                        foreach (string worksheetName in worksheetNames)
                        {
                            List<ClassToExcelRawRow> rawRows = rawReaderService.ReadWorksheet(filePath, worksheetName.Trim());
                            if (rawRows.Count > 0)
                            {
                                LogMessage($"Work sheet name: {worksheetName}");

                                foreach (var row in rawRows)
                                {
                                    StringBuilder sb = new StringBuilder();
                                    sb.AppendFormat("Row: {0} --> ", row.RowNumber);
                                    foreach (var column in row.Columns)
                                    {
                                        sb.AppendFormat("[Column: {0} Data: {1}]", column.ColumnLetter, column.Data);
                                    }
                                    LogMessage(sb.ToString());
                                }
                            }
                            else LogMessage($"No data found on work sheet named: {worksheetName}");

                            LogMessage("----------------------------------");
                        }
                    }
                    else LogMessage("No work sheet (tab) names specified.");
                }
                else LogMessage($"The file does not exists: {filePath}");
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        private void DoMapRowsToPropertiesWorkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("I'm using RAW read service (ClassToExcelRawReaderService) to get data and then " +
                    "I'm using a converter (ClassToExcelRowConverter) to map rows to a class' properties.");
                LogMessage("Note that blank rows are skipped by OpenXml.");

                // Load the embedded Beverages.xlsx file.
                String someDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string filePath = Path.Combine(someDirectory, "Beverages.xlsx");
                if (File.Exists(filePath))
                {
                    LogMessage($"Loading data from {filePath}");
                    // Get tab names
                    List<string> worksheetNames = string.IsNullOrWhiteSpace(WorksheetNames.Text)
                        ? new List<string>()
                        : WorksheetNames.Text.Split(',').ToList();

                    if (worksheetNames.Count > 0)
                    {
                        var worksheetName = worksheetNames.First().Trim();

                        // Passing in LogServiceMessage in optional.  I'm using it here to help debug in issues while reading a file.
                        var rawReaderService = new ClassToExcelRawReaderService(LogServiceMessage);
                        List<ClassToExcelRawRow> rawRows = rawReaderService.ReadWorksheet(filePath, worksheetName);
                        if (rawRows.Count > 0)
                        {
                            // Rows 2-6 (Column C) have the Beverage class data
                            var converter1 = new ClassToExcelRowConverter<Beverage>(LogServiceMessage);
                            Beverage beverage = converter1.Convert(rawRows);
                            LogMessage(SerializationHelper.SerializeToXml(beverage));

                            // Rows 8-9 (Columns B & C) have the BeverageDates class data
                            var converter2 = new ClassToExcelRowConverter<BeverageDates>(LogServiceMessage);
                            BeverageDates dates = converter2.Convert(rawRows);
                            LogMessage(SerializationHelper.SerializeToXml(dates));

                        }
                        else LogMessage("No data found.");


                    }
                    else LogMessage("No work sheet (tab) names specified");
                }
                else LogMessage($"The file does not exists: {filePath}");
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }



        private void FindFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog(this) == true)
            {
                FilePath.Text = dialog.FileName;
            }
        }

        #region Logging
        private void LogServiceMessage(ClassToExcelMessage message)
        {
            LogMessage(message.ToString());
        }
        private delegate void NoArgsDelegate();

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ClearLog();
        }

        private void ClearLog()
        {
            if (Dispatcher.CheckAccess())
            {
                RtbLog.Document.Blocks.Clear();
            }
            else this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new NoArgsDelegate(ClearLog));
        }

        /// <summary>Threadsafe logging method.</summary>
        private void LogMessage(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                var p = new Paragraph(new Run(message));
                p.Foreground = Brushes.Black;
                RtbLog.Document.Blocks.Add(p);
            }
            else this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action<string>(LogMessage), message);
        }

        private void LogError(Exception ex)
        {
            if (Dispatcher.CheckAccess())
            {
                // We are back on the UI thread here so calling LogMessage will not cause a BeginInvoke for all these LogMessage calls:
                LogMessage(ex.Message);
                LogMessage(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    LogMessage(ex.InnerException.Message);
                    LogMessage(ex.InnerException.StackTrace);
                }
            }
            else this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action<Exception>(LogError), ex);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveLog();
        }

        private void SaveLog()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog();
            if (dialog.ShowDialog() != true)
                return;

            using (var fs = new FileStream(dialog.FileName, FileMode.Create))
            {
                var myTextRange = new TextRange(RtbLog.Document.ContentStart, RtbLog.Document.ContentEnd);
                myTextRange.Save(fs, DataFormats.Text);
            }
        }
        #endregion
    }
}
