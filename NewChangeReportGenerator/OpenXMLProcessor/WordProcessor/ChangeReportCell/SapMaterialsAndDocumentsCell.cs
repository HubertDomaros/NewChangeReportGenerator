using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.Core;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

internal class SapMaterialsAndDocumentsCell : IChangeReportCell {
    private readonly bool _sapMaterialCheckbox, _documentsCheckbox;
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly string[] _rowNumberArray;
    private readonly string[] _sapObjectArray;
    private readonly Dictionary<string, string>[] _definedByDictionariesArray;

    /// <summary>
    /// Inserts table cell with SAP Material item revision and corresponding documents.
    /// </summary>
    /// <param name="rowNumber">Number of inserted row</param>
    /// <returns>DocumentFormat.OpenXML.Wordprocessing.TableCell</returns>
    public TableCell InsertCell(int rowNumber) {
        var cell = new TableCell();

        //Appending paragraph with SAP Material
        cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(_mainDocumentPart, _sapMaterialCheckbox, _rowNumberArray[rowNumber], _sapObjectArray[rowNumber]));

        //Appending paragraph with document/documents to sap materials
        foreach (var documentKeyValuePair in _definedByDictionariesArray[rowNumber]) {
            cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(_mainDocumentPart, _documentsCheckbox, documentKeyValuePair.Key, documentKeyValuePair.Value));
        }

        //Cell style formatting
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        
        return cell;
    }

    public SapMaterialsAndDocumentsCell(MainDocumentPart mainDocumentPart, SortedData sortedData, CheckboxesConfig checkboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _rowNumberArray = sortedData.RowNumberArray;
        _sapObjectArray = sortedData.SapObjectArray;
        _definedByDictionariesArray = sortedData.DefinedByDictionariesArray;
        _sapMaterialCheckbox = checkboxesConfig.SapMaterialCheckboxBool;
        _documentsCheckbox = checkboxesConfig.DocumentsCheckboxBool;
    }
}