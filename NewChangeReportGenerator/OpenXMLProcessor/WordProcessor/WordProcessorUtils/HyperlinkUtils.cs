using System;
using System.Diagnostics;
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
        //create URI from url string
        Uri uri = new Uri(url);

        HyperlinkRelationship rel = _documentPart.AddHyperlinkRelationship(uri, true);
        string relationshipId = rel.Id;

        //Create hyperlink run
        Run run = new Run(SetHyperlinkStyle("#0000ff"), new Text(urlLabel));

        //Add paragraph with hyperlink
        Paragraph hyperlinkParagraph = new Paragraph(new Hyperlink(new ProofError() { Type = ProofingErrorValues.GrammarStart }, run) { History = OnOffValue.FromBoolean(true), Id = relationshipId });
        hyperlinkParagraph.PrependChild(new ParagraphProperties());
        Debug.Assert(_documentPart.Document.Body != null, "documentPart.Document.Body != null");

        return hyperlinkParagraph;
    }

    private static RunProperties SetHyperlinkStyle(string color) {
        RunProperties runProperties = new RunProperties(new RunStyle() { Val = "Hyperlink" });
        Color col = new Color();
        col.Val = color;
        runProperties.Append(new Underline() { Val = UnderlineValues.Single }, col);
        return runProperties;
    }
}