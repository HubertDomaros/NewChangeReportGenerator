﻿using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace NewChangeReportGenerator.Core;

internal class ItemSortingAlgorithm {
    private readonly string[] _rowNumberArray; //First row generated by Teamcenter, which contains ordinal number with hyperlink to SAP Material or Document
    private readonly string[] _sapObjectArray;
    private readonly string[] _definedByArray;

    public Dictionary<string, string>[] DefinedByItemsWithUrls() {
        List<string> definedByItemNameList = SplitAndConvertStringArrayToList(_sapObjectArray);
        Dictionary<string, string>[] definedByWithUrlsDictionaries = new Dictionary<string, string>[definedByItemNameList.Count];

        for (int i = 0; i < _definedByArray.Length; i++) {
            foreach (string definedByItemName in definedByItemNameList) {
                for (int j = 0; j < _sapObjectArray.Length; j++) {
                    if (_sapObjectArray[j].Contains(definedByItemName)) {
                        definedByWithUrlsDictionaries[j].Add(definedByItemName, _rowNumberArray[j]);
                    }
                }
            }
        }
        return definedByWithUrlsDictionaries;
    }

    private List<string> SplitAndConvertStringArrayToList(string[] inputArray) {
        List<string> outputList = new List<string>();

        foreach (var inputString in inputArray) {
            string[] splittedStrings = inputString.Split(", D", StringSplitOptions.TrimEntries);

            foreach (string splittedString in splittedStrings) {
                outputList.Add(splittedString);
            }
        }

        return outputList;
    }

    public ItemSortingAlgorithm(string[] rowNumberArray, string[] sapObjectArray, string[] definedByArray) {
        _rowNumberArray = rowNumberArray;
        _sapObjectArray = sapObjectArray;
        _definedByArray = definedByArray;
    }
}