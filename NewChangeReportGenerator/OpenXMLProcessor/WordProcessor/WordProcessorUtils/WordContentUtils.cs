using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils; 

internal class WordContentUtils {
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