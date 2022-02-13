using System;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils; 

internal class HyperlinkUtils {

    private static MainDocumentPart _documentPart = null!;

    public static Paragraph InjectParagraphWithOptionalHyperlink(MainDocumentPart documentPart, bool isHyperlink, string itemRevisionName, string itemRevisionUrl) {
        _documentPart = documentPart;

        if (isHyperlink) {
            return InjectHyperlinkIntoTable(itemRevisionUrl, itemRevisionName);
        }
        return new Paragraph(new Run(new Text(itemRevisionName)));
    }

    private static Paragraph InjectHyperlinkIntoTable(string url, string urlLabel) {
        //add the url
        Uri uri = new Uri(url);

        HyperlinkRelationship rel = _documentPart.AddHyperlinkRelationship(uri, true);
        string relationshipId = rel.Id;

        //Set hyperlink style
        RunProperties runProperties = new RunProperties(new RunStyle() { Val = "Hyperlink" });
        Run run = new Run(runProperties, new Text(urlLabel));
        Color color = new Color();
        color.Val = "#0000ff";
        runProperties.Append(new Underline() { Val = UnderlineValues.Single }, color);

        //Add paragraph with hyperlink
        Paragraph newParagraph = new Paragraph(new Hyperlink(new ProofError() { Type = ProofingErrorValues.GrammarStart }, run) { History = OnOffValue.FromBoolean(true), Id = relationshipId });
        newParagraph.PrependChild(new ParagraphProperties());
        Debug.Assert(_documentPart.Document.Body != null, "documentPart.Document.Body != null");

        return newParagraph;
    }
}