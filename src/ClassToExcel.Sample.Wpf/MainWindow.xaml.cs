using System;
using System.Collections.Generic;
using System.IO;
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
                    // Passing in LogServiceMessage in optional.  I'm using it here to help debug in issues while reading a file.
                    var readerService = new ClassToExcelReaderService<Person>(LogServiceMessage);
                    List<Person> people = readerService.ReadWorksheet(filePath, "People", true);
                    if (people.Count > 0)
                    {
                        foreach (var person in people)
                        {
                            LogMessage(person.ToString());
                        }
                    }
                    else LogMessage("No data found.");
                }
                else LogMessage("The file does not exists: " + filePath);
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }
        

        private void DoWriteWorkButton_Click(object sender, RoutedEventArgs e)
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
                string filePath = FilePath.Text;
                if (File.Exists(filePath))
                {
                    // Passing in LogServiceMessage in optional.  I'm using it here to help debug in issues while reading a file.
                    var rawReaderService = new ClassToExcelRawReaderService(LogServiceMessage);
                    List<ClassToExcelRawRow> rawRows = rawReaderService.ReadWorksheet(filePath, "People");
                    if (rawRows.Count > 0)
                    {
                        
                        foreach (var row in rawRows)
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendFormat("Row: {0} --> ", row.RowIndex);
                            foreach (var column in row.Columns)
                            {
                                sb.AppendFormat("[Column: {0} Data: {1}]", column.ColumnLetter, column.Data);
                            }
                            LogMessage(sb.ToString());
                        }
                    }
                    else LogMessage("No data found.");
                }
                else LogMessage("The file does not exists: " + filePath);
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
