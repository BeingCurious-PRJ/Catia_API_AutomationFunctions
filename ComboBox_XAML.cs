using System.Collections.Generic;

public class ComboBoxViewHandler
{
    List<string> comboBoxItems;
    public ComboBoxViewHandler(List<string> items)
    {
        comboBoxItems = items;
    }
    public List<string> ComboBoxItems {get => comboBoxItems; set => comboBoxItems = value;}
}