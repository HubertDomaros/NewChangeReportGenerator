using System.Collections.Generic;
using System.Diagnostics;

namespace ChangeNotificationGenerator.DebugUtilities; 

public class DebugUtils {
    public static void PrintDebuggedList(List<string> debuggedList) {
        foreach (var item in debuggedList) {
            Debug.Print(item);
        }
        Debug.Print("--------------------------");
    }

    public static void PrintDebuggedList(List<string> debuggedList, string titleText) {
        Debug.Print(titleText);
        foreach (var item in debuggedList) {
            Debug.Print(item);
        }
        Debug.Print("--------------------------");
    }
    
    public static void PrintDebuggedDictionariesArray(Dictionary<string,string>[] dictionariesArray, string titleText) {
        foreach (var item in dictionariesArray) {
            foreach (KeyValuePair<string, string> kvp in item) {
                Debug.Print("Document name:", kvp.Key, "URL:", kvp.Value);
            }
}