using Microsoft.Win32;
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
using TrackingDocGenerator.Models;
using TrackingDocGenerator.Services;

namespace TrackingDocGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _excelPath;
        private readonly ExcelReader _excelReader = new ExcelReader();
        private readonly WordGenerator _wordGenerator = new WordGenerator();
        private readonly PrinterService _printerService = new PrinterService();
        private string _lastGeneratedPath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSelectExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Title = "Select Excel File"
            };
            if (dialog.ShowDialog() == true)
            {
                _excelPath = dialog.FileName;
                txtStatus.Text = $"Selected Excel file: {_excelPath}";
            }
        }
        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_excelPath) )
            {
                MessageBox.Show("Please select both Excel file and Word template.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                List<TrackingInfo> data = _excelReader.ReadTrackingData(_excelPath);
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    Title = "Save Generated Document"
                };
                if(saveDialog.ShowDialog() == true)
                {
                    _wordGenerator.GenerateLabel(saveDialog.FileName, data);
                    txtStatus.Text += $"\nDocument generated successfully: {saveDialog.FileName}";
                    _lastGeneratedPath = saveDialog.FileName;
                    MessageBox.Show("Document generated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_lastGeneratedPath) || !File.Exists(_lastGeneratedPath))
            {
                MessageBox.Show("File still not generate or not fined.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _printerService.Print(_lastGeneratedPath);
            MessageBox.Show("Document was send for print.", "Print", MessageBoxButton.OK, MessageBoxImage.Information);
        }

    }
}
