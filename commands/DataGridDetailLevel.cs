using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public static class DataGridDetailLevelManager
{
    private static readonly string ModeFilePath = PathHelper.GetRuntimeFilePath("DataGridDetailLevel");

    public enum DetailLevel { Base, Extended }

    public static DetailLevel CurrentLevel
    {
        get
        {
            if (File.Exists(ModeFilePath))
            {
                string val = File.ReadAllText(ModeFilePath).Trim();
                if (System.Enum.TryParse<DetailLevel>(val, out var level))
                    return level;
            }
            return DetailLevel.Base;
        }
        set { File.WriteAllText(ModeFilePath, value.ToString()); }
    }

    public static readonly string[] BaseColumns = new[]
    {
        "Name", "Type Name", "Family", "Category", "Group", "LinkName", "Id"
    };
}

[Transaction(TransactionMode.Manual)]
[CommandMeta("")]
public class SwitchDataGridDetailLevel : IExternalCommand
{
    private class LevelWrapper
    {
        public string Level { get; set; }
        public DataGridDetailLevelManager.DetailLevel EnumValue { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var currentLevel = DataGridDetailLevelManager.CurrentLevel;

        var levels = new List<LevelWrapper>
        {
            new LevelWrapper { Level = "Base",     EnumValue = DataGridDetailLevelManager.DetailLevel.Base },
            new LevelWrapper { Level = "Extended",  EnumValue = DataGridDetailLevelManager.DetailLevel.Extended },
        };

        var propertyNames = new List<string> { "Level" };
        int selectedIndex = levels.FindIndex(l => l.EnumValue == currentLevel);
        var initialSelection = new List<int> { selectedIndex };

        var itemDicts = CustomGUIs.ConvertToDataGridFormat(levels, propertyNames);
        var selectedDicts = CustomGUIs.DataGrid(itemDicts, propertyNames, false, initialSelection);
        var selectedLevels = CustomGUIs.ExtractOriginalObjects<LevelWrapper>(selectedDicts);

        if (selectedLevels == null || selectedLevels.Count == 0)
            return Result.Cancelled;

        DataGridDetailLevelManager.CurrentLevel = selectedLevels.First().EnumValue;
        return Result.Succeeded;
    }
}
