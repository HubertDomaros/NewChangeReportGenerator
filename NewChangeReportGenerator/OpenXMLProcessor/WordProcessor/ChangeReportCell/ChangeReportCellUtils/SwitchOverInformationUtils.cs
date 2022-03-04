using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell.ChangeReportCellUtils; 

public class SwitchOverInformationUtils {

    public Run DropdownTitleText(string title) {
        var textRun = new Run();

        //Setting up run formatting
        RunProperties runProperties = textRun.AppendChild(new RunProperties());
        Bold bold = new Bold();
        bold.Val = OnOffValue.FromBoolean(true);
        runProperties.AppendChild(bold);

        //Appending text to formatted run
        textRun.Append(new Break(), new Text(title), new Break());
        return textRun;
    }

    public SdtRun CreateDropdown(List<string> dropdownValuesList, string defaultDropdownText) {
        var dropdownSdtRun = new SdtRun();

        var dropdownSdtRunProperties = new SdtProperties();
        var dropdown = new SdtContentDropDownList();

        foreach (var listItemText in dropdownValuesList) {
            var listItem = new ListItem() { DisplayText = listItemText, Value = listItemText };
            dropdown.Append(listItem);
        }

        dropdownSdtRunProperties.Append(dropdown);

        //Overwriting first dropdown value with new value
        var defaultText = new SdtContentRun(new Run(new Text(defaultDropdownText)));
        dropdownSdtRun.Append(defaultText);

        return dropdownSdtRun;
    }
}