namespace ChangeNotificationGenerator.Core;

public class CheckboxesConfig {
    public bool RowNumberCheckboxBool { get; set; }
    public bool SapMaterialCheckboxBool { get; set; }
    public bool DocumentsCheckboxBool { get; set; }

    public CheckboxesConfig(bool rowNumberCheckboxBool, bool sapMaterialCheckboxBool, bool documentsCheckboxBool) {
        RowNumberCheckboxBool = rowNumberCheckboxBool;
        SapMaterialCheckboxBool = sapMaterialCheckboxBool;
        DocumentsCheckboxBool = documentsCheckboxBool;
    }

    public CheckboxesConfig() { }
}