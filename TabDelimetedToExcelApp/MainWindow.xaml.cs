using Microsoft.Win32;
using System.IO;
using System.Windows;
using TabDelimitedToExcelLibrary;
using System;

namespace TabDelimetedToExcelApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void open_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();

            dialog.DefaultExt = ".txt";
            dialog.Filter = "TXT Files (*.txt)|*.txt|All Files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method 
            bool? result = dialog.ShowDialog();

            // Get the selected file name and display in a textbox 
            if (result == true)
            {
                fileTextBox.Text = dialog.FileName;
            }

        }

        private void transform_Click(object sender, RoutedEventArgs e)
        {
            var file = new ExcelFile();
            try
            {
                var transformer = new TabDelimitedToExcelTransformer(fileTextBox.Text);
                file = transformer.Transform();
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.AddExtension = true;
                saveFileDialog.DefaultExt = ".xls";
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(fileTextBox.Text) + "_toExcel";

                saveFileDialog.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*";
                if (saveFileDialog.ShowDialog() == true)
                {
                    file.SaveAs(saveFileDialog.FileName);
                    label_Copy.Visibility = Visibility.Visible;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                file.Close();
            }
        }

        private void fileTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (File.Exists(fileTextBox.Text))
            {
                transform.IsEnabled = true;
            }
            else
            {
                transform.IsEnabled = false;
            }
        }
    }
}
