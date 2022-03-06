﻿using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell.ChangeReportCellUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;

internal class SwitchOverInformationCell : IChangeReportCell {

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

    private readonly List<string> _serviceUpgradeDropdownList = new() {
        "No rework/retrofitting",
        "Immediate rework/retrofitting",
        "Packaged rework/retrofitting",
        "Optional rework/retrofitting",
        "Other"
    };

    public TableCell InsertCell(int rowNumber) {
        var cell = new TableCell();
        var cellParagraph = new Paragraph();
        var switchOverInformationUtils = new SwitchOverInformationUtils();
        
        //Production change dropdown
        cellParagraph.Append(switchOverInformationUtils.DropdownTitleText("Production change"));
        cellParagraph.Append(switchOverInformationUtils.CreateDropdown(_productionChangeDropdownList, "Other"));
        //Remaining stock dropdown
        cellParagraph.Append(switchOverInformationUtils.DropdownTitleText("Remaining stock"));
        cellParagraph.Append(switchOverInformationUtils.CreateDropdown(_remainingStockDropdownList, "Other"));
        //Service upgrade dropdown
        cellParagraph.Append(switchOverInformationUtils.DropdownTitleText("Service upgrade"));
        cellParagraph.Append(switchOverInformationUtils.CreateDropdown(_serviceUpgradeDropdownList, "Other"));
        
        cell.Append(cellParagraph);
        return cell;
    }
}