using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

internal class SapMaterialsAndDocumentsCell : BaseChangeReportCell {
    private readonly bool _sapMaterialCheckbox, _documentsCheckbox;

    /// <summary>
    /// Inserts table cell with SAP Material item revision and corresponding documents.
    /// </summary>
    /// <param name="sapPair">Key-value pair for SAP material; Key should be SAP Material's name/additional name, and value- URL to this SAP Material in Teamcenter</param>
    /// <param name="documentsDictionary">Dictionary for documents and 3D parts; Key should be document's revision name/additional name, value- url to this document's revision in Teamcenter</param>
    /// <returns>DocumentFormat.OpenXML.Wordprocessing.TableCell</returns>
    public TableCell InsertCell(KeyValuePair<string, string> sapPair, Dictionary<string, string> documentsDictionary) {
        var cell = new TableCell();
        //Appending paragraph with SAP Material
        cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(DocumentPart, _sapMaterialCheckbox, sapPair.Key, sapPair.Value));

        //Appending paragraph with document/documents to sap materials
        foreach (var documentKeyValuePair in documentsDictionary) {
            cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(DocumentPart, _documentsCheckbox, documentKeyValuePair.Key, documentKeyValuePair.Value));
        }
        //Cell style formatting
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        return cell;
    }

    public SapMaterialsAndDocumentsCell(MainDocumentPart documentPart, bool sapMaterialCheckbox, bool documentsCheckbox) {
        DocumentPart = documentPart;
        _sapMaterialCheckbox = sapMaterialCheckbox;
        _documentsCheckbox = documentsCheckbox;
    }
}