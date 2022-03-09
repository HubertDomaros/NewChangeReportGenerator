﻿using System.Collections.Generic;
using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

internal class SapMaterialsAndDocumentsCell : IChangeReportCell {
    private readonly bool _sapMaterialCheckbox, _documentsCheckbox;
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly List<string> _rowNumberList;
    private readonly List<string> _sapObjectList;
    private readonly Dictionary<string, string>[] _definedByDictionariesArray;

    /// <summary>
    /// Inserts table cell with SAP Material item revision and corresponding documents.
    /// </summary>
    /// <param name="rowNumber">Number of inserted row</param>
    /// <returns>DocumentFormat.OpenXML.Wordprocessing.TableCell</returns>
    public TableCell InsertCell(int rowNumber) {
        var cell = new TableCell();

        //Appending paragraph with SAP Material
        cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(_mainDocumentPart, _sapMaterialCheckbox, _rowNumberList[rowNumber], _sapObjectList[rowNumber]));

        //Appending paragraph with document/documents to sap materials
        foreach (var documentKeyValuePair in _definedByDictionariesArray[rowNumber]) {
            cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(_mainDocumentPart, _documentsCheckbox, documentKeyValuePair.Key, documentKeyValuePair.Value));
        }

        //Cell style formatting
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        
        return cell;
    }

    public SapMaterialsAndDocumentsCell(MainDocumentPart mainDocumentPart, ChangeNotificationDataService sortedData, CheckboxesConfig checkboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _rowNumberList = sortedData.RowNumberList;
        _sapObjectList = sortedData.SapObjectList;
        _definedByDictionariesArray = sortedData.DefinedByDictionariesArray;
        _sapMaterialCheckbox = checkboxesConfig.SapMaterialCheckboxBool;
        _documentsCheckbox = checkboxesConfig.DocumentsCheckboxBool;
    }
}