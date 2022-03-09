using System.Diagnostics;
using System.Windows;
using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator; 

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window {
    public MainWindow() {
        InitializeComponent();
        InteractiveConsole();
    }

    private void InteractiveConsole() {
        Debug.WriteLine("Debug start");
        string filePath =
            @"C:\VisualStudioProjects\COCReator\EnerconCOCreator\EnerconCOCreator\DOCXOutputFiles\CO3718.xlsx";

        ExcelDocumentParser excelDocumentParser = new ExcelDocumentParser(filePath);

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false)) {
            ChangeNotificationDataService sortedData = new ChangeReportDataService(spreadsheetDocument.WorkbookPart, true);
        }

        
        Debug.Print("Debug end");
    }
}