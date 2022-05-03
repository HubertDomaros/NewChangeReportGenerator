using System.Collections.Generic;
using System.Windows;

namespace ChangeNotificationGenerator.Warnings;
public class ElementsNotFoundWarning {
    private readonly List<string> _notFoundElementsList;

    private void ShowMessageBox() {
        var message = "Excel file opened with warnings! \n" +
                         "Following items were not found in Object list: \n";

        foreach (var str in _notFoundElementsList) {
            message = message + str + "\n";
        }

        MessageBox.Show(message, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    public ElementsNotFoundWarning(List<string> notFoundElementsList) {
        _notFoundElementsList = notFoundElementsList;
        ShowMessageBox();
    }
}