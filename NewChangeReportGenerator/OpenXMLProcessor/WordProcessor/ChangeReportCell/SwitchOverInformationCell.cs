using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;

internal class SwitchOverInformationCell : BaseChangeReportCell {

    //Hardcoded due to time schedule
    private readonly List<string> _productionChangeDropdownList = new() {
        "...",
        "No production change",
        "Immediate production change",
        "Incorporating production change",
        "Release-controlled production change",
        "Other"
    };

    private readonly List<string> _remainingStockDropdownList = new() {
        "...",
        "Not expected",
        "Use up",
        "Rework",
        "Scrap",
        "Other"
    };

    public TableCell InsertCell() {
        var cell = new TableCell();
        var cellParagraph = new Paragraph();

        cellParagraph.Append(ProductionChangeTextRun());

        cell.Append(cellParagraph);
        return cell;
    }

    private Run ProductionChangeTextRun() {
        var productionChangeTextRun = new Run();

        var productionChangeRunProperties = productionChangeTextRun.AppendChild(new RunProperties());
        var bold = new Bold {
            Val = OnOffValue.FromBoolean(true)
        };
        productionChangeRunProperties.AppendChild(bold);

        productionChangeTextRun.Append(new Text("Production change"), new Break());

        return productionChangeTextRun;
    }

    private SdtRun ProductionChangeDropdown() {
        return new WordContentUtils().CreateDropdown(_productionChangeDropdownList, "Other"); ;
    }

    

    private SdtRun ProductionChangeDropdownList() {
        var productionChangeSdtRun = new SdtRun();
        var productionChangeSdtRunProperties = new SdtProperties();

        //Creating new dropdown
        var productionChangeDropDownList = new SdtContentDropDownList();
        var listItem1 = new ListItem() { DisplayText = "...", Value = "..." };
        var listItem2 = new ListItem() { DisplayText = "No production change", Value = "No production change" };
        var listItem3 = new ListItem() { DisplayText = "Immediate production change", Value = "Immediate production change" };
        var listItem4 = new ListItem() { DisplayText = "Incorporating production change", Value = "Incorporating production change" };
        var listItem5 = new ListItem() { DisplayText = "Release-controlled production change", Value = "Release-controlled production change" };
        var listItem6 = new ListItem() { DisplayText = "Other", Value = "Other" };

        productionChangeDropDownList.Append(listItem1);
        productionChangeDropDownList.Append(listItem2);
        productionChangeDropDownList.Append(listItem3);
        productionChangeDropDownList.Append(listItem4);
        productionChangeDropDownList.Append(listItem5);
        productionChangeDropDownList.Append(listItem6);

        productionChangeSdtRunProperties.Append(productionChangeDropDownList);

        //Default displayed text in dropdown
        var defaultTextSdtContentRun = new SdtContentRun();
        var defaultTextRun = new Run();
        var defaultText = new Text { Text = "Other" };

        defaultTextRun.Append(defaultText);
        defaultTextSdtContentRun.Append(defaultTextRun);
        productionChangeSdtRun.Append(defaultTextSdtContentRun);

        productionChangeSdtRun.Append(productionChangeSdtRunProperties);

        return new SdtRun();
    }

    public SwitchOverInformationCell(MainDocumentPart documentPart) {
        DocumentPart = documentPart;
    }
}