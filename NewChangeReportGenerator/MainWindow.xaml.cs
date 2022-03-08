using System.Diagnostics;
using System.Windows;
using NewChangeReportGenerator.Core;
using NewChangeReportGenerator.OpenXMLProcessor.ExcelProcessor;

namespace NewChangeReportGenerator; 

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window {
    public MainWindow() {
        //InitializeComponent();
        ParseExcelCell();
    }

    private void InteractiveConsole() {
        Debug.WriteLine("Debug start");
        
        string[] lines = System.IO.File.ReadAllLines(@"C:\VisualStudioProjects\NewChangeReportGenerator\NewChangeReportGenerator\NewChangeReportGenerator\OpenXMLProcessor\definedByTest.txt");

        MainSortingAlgorithm itemSorting = new MainSortingAlgorithm(lines, lines, lines);

        //itemSorting.PrintDefinedByArray(lines);
    }

    private void ParseExcelCell() {
        ExcelCellParser excelCellParser = new ExcelCellParser(@"C:\VisualStudioProjects\COCReator\EnerconCOCreator\EnerconCOCreator\DOCXOutputFiles\tc_1646039187079.xlsm");
    }
}