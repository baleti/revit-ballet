using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

[TransactionAttribute(TransactionMode.Manual)]
public class SelectSchedules : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the currently active view
        View activeView = doc.ActiveView;

        // Get all schedules in the project
        FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSchedule));

        // Prepare data for the data grid and map schedule titles to schedule objects
        List<Dictionary<string, object>> scheduleData = new List<Dictionary<string, object>>();
        Dictionary<string, ViewSchedule> titleToScheduleMap = new Dictionary<string, ViewSchedule>();

        foreach (ViewSchedule schedule in scheduleCollector.Cast<ViewSchedule>())
        {
            // Skip schedule templates
            if (schedule.IsTemplate)
                continue;

            // Get schedule type for additional info
            string scheduleType = GetScheduleType(schedule);

            // Get owner view if exists
            string ownerViewName = GetOwnerViewName(schedule);

            // Get detailed schedule properties
            var (schedColumns, filters, sorting, grouping) = GetScheduleDetails(schedule);

            // Assuming titles are unique; otherwise, you might need to use a different key
            titleToScheduleMap[schedule.Title] = schedule;
            Dictionary<string, object> scheduleInfo = new Dictionary<string, object>
            {
                { "Title", schedule.Title },
                { "Type", scheduleType },
                { "Owner View", ownerViewName },
                { "Columns", schedColumns },
                { "Filters", filters },
                { "Sorting", sorting },
                { "Grouping", grouping }
            };
            scheduleData.Add(scheduleInfo);
        }

        // Sort the scheduleData by Title
        scheduleData = scheduleData.OrderBy(s => s["Title"].ToString()).ToList();

        // Find the index of the active view after sorting (if it's a schedule)
        int sortedActiveScheduleIndex = -1;
        if (activeView != null && activeView is ViewSchedule)
        {
            for (int i = 0; i < scheduleData.Count; i++)
            {
                if (scheduleData[i]["Title"].ToString() == activeView.Title)
                {
                    sortedActiveScheduleIndex = i;
                    break;
                }
            }
        }

        // Define the column headers
        List<string> columns = new List<string> { "Title", "Type", "Owner View", "Columns", "Filters", "Sorting", "Grouping" };

        // Prepare initial selection indices (if active view is a schedule and was found)
        List<int> initialSelection = null;
        if (sortedActiveScheduleIndex >= 0)
        {
            initialSelection = new List<int> { sortedActiveScheduleIndex };
        }

        // Show the selection dialog (using your custom GUI)
        List<Dictionary<string, object>> selectedSchedules = CustomGUIs.DataGrid(
            scheduleData,
            columns,
            false,  // Don't span all screens
            initialSelection  // Pass the initial selection
        );

        // If the user made a selection, add those elements to the current selection
        if (selectedSchedules != null && selectedSchedules.Any())
        {
            // Get the current selection
            ICollection<ElementId> currentSelectionIds = uidoc.GetSelectionIds();
            
            // Get the ElementIds of the schedules selected in the dialog
            List<ElementId> newScheduleIds = selectedSchedules
                .Select(s => titleToScheduleMap[s["Title"].ToString()].Id)
                .ToList();
                
            // Add the new schedules to the current selection
            foreach (ElementId id in newScheduleIds)
            {
                if (!currentSelectionIds.Contains(id))
                {
                    currentSelectionIds.Add(id);
                }
            }
            
            // Update the selection with the combined set of elements
            uidoc.SetSelectionIds(currentSelectionIds);
            
            return Result.Succeeded;
        }

        return Result.Cancelled;
    }

    /// <summary>
    /// Helper method to determine the type of schedule
    /// </summary>
    private string GetScheduleType(ViewSchedule schedule)
    {
        if (schedule.IsTitleblockRevisionSchedule)
            return "Revision";
        else if (schedule.IsInternalKeynoteSchedule)
            return "Keynote";
        else if (schedule.Definition != null)
        {
            // Try to get the category name from the schedule definition
            ScheduleDefinition definition = schedule.Definition;
            if (definition.CategoryId != null && definition.CategoryId != ElementId.InvalidElementId)
            {
                Category category = Category.GetCategory(schedule.Document, definition.CategoryId);
                if (category != null)
                    return category.Name;
            }
        }
        
        return "Schedule";
    }

    /// <summary>
    /// Helper method to get the owner view name if the schedule has one
    /// </summary>
    private string GetOwnerViewName(ViewSchedule schedule)
    {
        try
        {
            // Check if the schedule has an owner view
            ElementId ownerViewId = schedule.OwnerViewId;
            if (ownerViewId != null && ownerViewId != ElementId.InvalidElementId)
            {
                View ownerView = schedule.Document.GetElement(ownerViewId) as View;
                if (ownerView != null)
                {
                    return ownerView.Title;
                }
            }
        }
        catch
        {
            // If OwnerViewId property doesn't exist or throws an exception
        }
        
        return "None";
    }

    /// <summary>
    /// Helper method to get detailed schedule properties
    /// </summary>
    private (string columns, string filters, string sorting, string grouping) GetScheduleDetails(ViewSchedule schedule)
    {
        List<string> columnsList = new List<string>();
        List<string> filtersList = new List<string>();
        List<string> sortingList = new List<string>();
        List<string> groupingList = new List<string>();

        try
        {
            ScheduleDefinition definition = schedule.Definition;
            if (definition != null)
            {
                // Get column fields
                int fieldCount = definition.GetFieldCount();
                for (int i = 0; i < fieldCount; i++)
                {
                    ScheduleField field = definition.GetField(i);
                    if (!field.IsHidden)
                    {
                        string fieldName = field.GetName();
                        if (field.IsCalculatedField)
                            fieldName += " (calc)";
                        columnsList.Add(fieldName);
                    }
                }

                // Get filters
                int filterCount = definition.GetFilterCount();
                for (int i = 0; i < filterCount; i++)
                {
                    ScheduleFilter filter = definition.GetFilter(i);
                    ScheduleField filterField = definition.GetField(filter.FieldId);
                    string filterDesc = filterField.GetName();
                    
                    // Add filter type
                    switch (filter.FilterType)
                    {
                        case ScheduleFilterType.Equal:
                            filterDesc += " =";
                            break;
                        case ScheduleFilterType.NotEqual:
                            filterDesc += " ≠";
                            break;
                        case ScheduleFilterType.GreaterThan:
                            filterDesc += " >";
                            break;
                        case ScheduleFilterType.GreaterThanOrEqual:
                            filterDesc += " ≥";
                            break;
                        case ScheduleFilterType.LessThan:
                            filterDesc += " <";
                            break;
                        case ScheduleFilterType.LessThanOrEqual:
                            filterDesc += " ≤";
                            break;
                        case ScheduleFilterType.Contains:
                            filterDesc += " contains";
                            break;
                        case ScheduleFilterType.NotContains:
                            filterDesc += " !contains";
                            break;
                        case ScheduleFilterType.BeginsWith:
                            filterDesc += " begins";
                            break;
                        case ScheduleFilterType.NotBeginsWith:
                            filterDesc += " !begins";
                            break;
                        case ScheduleFilterType.EndsWith:
                            filterDesc += " ends";
                            break;
                        case ScheduleFilterType.NotEndsWith:
                            filterDesc += " !ends";
                            break;
                        case ScheduleFilterType.HasValue:
                            filterDesc += " has value";
                            break;
                        case ScheduleFilterType.HasNoValue:
                            filterDesc += " no value";
                            break;
                    }
                    
                    filtersList.Add(filterDesc);
                }

                // Get sorting and grouping
                int sortGroupCount = definition.GetSortGroupFieldCount();
                
                for (int i = 0; i < sortGroupCount; i++)
                {
                    ScheduleSortGroupField sortGroupField = definition.GetSortGroupField(i);
                    ScheduleField field = definition.GetField(sortGroupField.FieldId);
                    string fieldName = field.GetName();
                    
                    if (sortGroupField.ShowHeader || sortGroupField.ShowFooter || sortGroupField.ShowBlankLine)
                    {
                        // This is a grouping field
                        string groupDesc = fieldName;
                        if (sortGroupField.SortOrder == ScheduleSortOrder.Descending)
                            groupDesc += " ↓";
                        else
                            groupDesc += " ↑";
                        if (sortGroupField.ShowFooter)
                            groupDesc += " (footer)";
                        groupingList.Add(groupDesc);
                    }
                    else
                    {
                        // This is just a sorting field
                        string sortDesc = fieldName;
                        if (sortGroupField.SortOrder == ScheduleSortOrder.Descending)
                            sortDesc += " ↓";
                        else
                            sortDesc += " ↑";
                        sortingList.Add(sortDesc);
                    }
                }
            }
        }
        catch
        {
            // In case of any error accessing schedule definition
        }

        // Return comma-separated strings or "-" if empty
        string columnsStr = columnsList.Count > 0 ? string.Join(", ", columnsList) : "-";
        string filtersStr = filtersList.Count > 0 ? string.Join(", ", filtersList) : "-";
        string sortingStr = sortingList.Count > 0 ? string.Join(", ", sortingList) : "-";
        string groupingStr = groupingList.Count > 0 ? string.Join(", ", groupingList) : "-";

        return (columnsStr, filtersStr, sortingStr, groupingStr);
    }
}
