using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCells;

internal class ContentsCell : ChangeReportCell {
    private readonly bool _sapMaterialCheckbox, _documentsCheckbox;
    private static MainDocumentPart _documentPart;

    /// <summary>
    /// Inserts table cell with SAP Material item revision and corresponding documents.
    /// </summary>
    /// <param name="sapPair">Key-value pair for SAP material; Key should be SAP Material's name/additional name, and value- URL to this SAP Material in Teamcenter</param>
    /// <param name="documentsDictionary">Dictionary for documents and 3D parts; Key should be document's revision name/additional name, value- url to this document's revision in Teamcenter</param>
    /// <returns>DocumentFormat.OpenXML.Wordprocessing.TableCell</returns>
    public TableCell InsertCell(KeyValuePair<string, string> sapPair, Dictionary<string, string> documentsDictionary) {
        var cell = new TableCell();

        cell.Append(HyperlinkUtils.IsHyperlinkParagraph(_documentPart,_sapMaterialCheckbox, sapPair.Key, sapPair.Value));

        foreach (var documentKeyValuePair in documentsDictionary) {
            cell.Append(HyperlinkUtils.IsHyperlinkParagraph(_documentPart, _documentsCheckbox, documentKeyValuePair.Key, documentKeyValuePair.Value));
        }
        return cell;
    }

    public ContentsCell(MainDocumentPart documentPart, bool sapMaterialCheckbox, bool documentsCheckbox) {
        _sapMaterialCheckbox = sapMaterialCheckbox;
        _documentsCheckbox = documentsCheckbox;
    }
}