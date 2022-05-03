using System;
using System.Windows;
using ChangeNotificationGenerator.Core;
using Microsoft.Win32;

namespace ChangeNotificationGenerator;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window {
    private readonly ChangeNotificationGeneratorController _changeNotificationController;
    private CheckboxesConfig _checkboxesConfig;


    public MainWindow() {
        InitializeComponent();
        _changeNotificationController = new ChangeNotificationGeneratorController();
    }

    private void BtnOpenFile_Click(object sender, RoutedEventArgs e) {
        try {
            var openFileDialog = new OpenFileDialog {
                Filter = "Microsoft Excel file (.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true) {
                _changeNotificationController.ProcessExcelDocument(openFileDialog.FileName);
                MessageBox.Show("File processing finished!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        } catch (Exception ex) {
            MessageBox.Show("ERROR: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void cbFeature_CheckedChanged(object sender, RoutedEventArgs e) {
        _checkboxesConfig = new CheckboxesConfig {
            RowNumberCheckboxBool = (bool)RowNumberCheckBox.IsChecked,
            SapMaterialCheckboxBool = (bool)SapMaterialCheckBox.IsChecked,
            DocumentsCheckboxBool = (bool)DocumentsCheckBox.IsChecked
        };
    }

    private void btnGenerateChangeNotification_Click(object sender, RoutedEventArgs e) {
        try {
            SaveFileDialog saveFileDialog = new SaveFileDialog {
                Filter = "Microsoft Word documents (.docx)|*.docx"
            };

            if (saveFileDialog.ShowDialog() == true) {
                try {
                    _changeNotificationController.GenerateChangeNotificationDocument(saveFileDialog.FileName, _checkboxesConfig);
                    MessageBox.Show("File created successfully! \n You can find created document in following location:\n" + saveFileDialog.FileName, "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                } catch (Exception ex) {
                    if (ex.GetType() == typeof(NullReferenceException) || ex.GetType() == typeof(ArgumentNullException)) {
                        MessageBox.Show("Excel file was not loaded or does not include any data", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    } else {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        } catch (Exception ex) {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}