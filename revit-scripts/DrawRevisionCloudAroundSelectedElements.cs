using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class DrawRevisionCloudsAroundSelectedElements : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the active view and selected elements
        View activeView = doc.ActiveView;
        ICollection<ElementId> selectedElementIds = uidoc.Selection.GetElementIds();

        if (!selectedElementIds.Any())
        {
            message = "Please select one or more elements in the active view.";
            return Result.Failed;
        }

        // Get the sheet that contains the active view
        ViewSheet currentSheet = null;
        Viewport activeViewport = null;

        FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));

        foreach (ViewSheet sheet in sheetCollector)
        {
            ICollection<ElementId> viewports = sheet.GetAllViewports();
            foreach (ElementId viewportId in viewports)
            {
                Viewport viewport = doc.GetElement(viewportId) as Viewport;
                if (viewport != null && viewport.ViewId == activeView.Id)
                {
                    currentSheet = sheet;
                    activeViewport = viewport;
                    break;
                }
            }
            if (currentSheet != null)
                break;
        }

        if (currentSheet == null)
        {
            message = "Could not find the sheet that contains the active view.";
            return Result.Failed;
        }

        // Ensure the view has an active crop box
        if (!activeView.CropBoxActive || activeView.CropBox == null)
        {
            message = "The active view must have an active crop box.";
            return Result.Failed;
        }

        // Offset distance to apply around each element (in model units)
        double offset = 0.2;  // Adjust as needed (in feet)

        // Get the transformation from view to sheet coordinates
        Transform transform;
        try
        {
            transform = GetViewToSheetTransform(activeViewport, doc);
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }

        // List to store all curves from all loops
        List<Curve> allCurves = new List<Curve>();

        // Iterate through each selected element and create curves around it
        foreach (ElementId id in selectedElementIds)
        {
            Element element = doc.GetElement(id);
            BoundingBoxXYZ elementBB = element.get_BoundingBox(activeView);

            if (elementBB != null)
            {
                // Expand the bounding box by the offset in model coordinates
                XYZ minPoint = new XYZ(elementBB.Min.X - offset, elementBB.Min.Y - offset, elementBB.Min.Z);
                XYZ maxPoint = new XYZ(elementBB.Max.X + offset, elementBB.Max.Y + offset, elementBB.Max.Z);

                // Transform expanded bounding box to sheet space
                XYZ minInSheet = transform.OfPoint(minPoint);
                XYZ maxInSheet = transform.OfPoint(maxPoint);

                // Create curves for the bounding box in sheet coordinates
                // Ensure each loop is properly closed by connecting the last point to the first
                allCurves.Add(Line.CreateBound(new XYZ(minInSheet.X, minInSheet.Y, 0), new XYZ(minInSheet.X, maxInSheet.Y, 0)));
                allCurves.Add(Line.CreateBound(new XYZ(minInSheet.X, maxInSheet.Y, 0), new XYZ(maxInSheet.X, maxInSheet.Y, 0)));
                allCurves.Add(Line.CreateBound(new XYZ(maxInSheet.X, maxInSheet.Y, 0), new XYZ(maxInSheet.X, minInSheet.Y, 0)));
                allCurves.Add(Line.CreateBound(new XYZ(maxInSheet.X, minInSheet.Y, 0), new XYZ(minInSheet.X, minInSheet.Y, 0)));
            }
        }

        if (!allCurves.Any())
        {
            message = "Could not determine the bounding boxes of the selected elements.";
            return Result.Failed;
        }

        // Find the latest revision available on the sheet
        Revision latestRevision = GetLatestRevision(doc);

        // If no revisions are found, create a new one
        if (latestRevision == null)
        {
            using (Transaction t = new Transaction(doc, "Create Revision"))
            {
                t.Start();
                latestRevision = Revision.Create(doc);
                t.Commit();
            }
        }

        // Start a transaction to create the revision cloud
        using (Transaction t = new Transaction(doc, "Create Revision Cloud"))
        {
            t.Start();

            // Create a new revision cloud using all the curves
            RevisionCloud cloud = RevisionCloud.Create(doc, currentSheet, latestRevision.Id, allCurves);

            t.Commit();
        }

        return Result.Succeeded;
    }

    private Transform GetViewToSheetTransform(Viewport viewport, Document doc)
    {
        // Get the view associated with the viewport
        View view = doc.GetElement(viewport.ViewId) as View;

        // Get the view's crop box
        BoundingBoxXYZ cropBox = view.CropBox;
        if (cropBox == null)
            throw new InvalidOperationException("View does not have a valid crop box.");

        // Get the crop box transform (accounts for view rotation)
        Transform cropTransform = cropBox.Transform;

        // Center of the crop box in model coordinates
        XYZ cropMin = cropBox.Min;
        XYZ cropMax = cropBox.Max;
        XYZ viewCenterInModel = (cropMin + cropMax) / 2.0;
        viewCenterInModel = cropTransform.OfPoint(viewCenterInModel);

        // Get the viewport's center on the sheet
        XYZ viewportCenterOnSheet = viewport.GetBoxCenter();

        // Get the rotation angle of the viewport
        double rotationAngle = 0.0;
        switch (viewport.Rotation)
        {
            case ViewportRotation.None:
                rotationAngle = 0.0;
                break;
            case ViewportRotation.Clockwise:
                rotationAngle = -0.5 * Math.PI;
                break;
            case ViewportRotation.Counterclockwise:
                rotationAngle = 0.5 * Math.PI;
                break;
            default:
                throw new InvalidOperationException("Viewport rotation is not supported.");
        }

        // Create rotation transform (around the origin)
        Transform rotationTransform = Transform.CreateRotation(XYZ.BasisZ, rotationAngle);

        // Create scaling transform (around the origin)
        double scale = 1.0 / view.Scale; // Convert view scale to scaling factor

        Transform scalingTransform = Transform.Identity;
        scalingTransform.BasisX = scalingTransform.BasisX.Multiply(scale);
        scalingTransform.BasisY = scalingTransform.BasisY.Multiply(scale);

        // Combine scaling and rotation
        Transform srTransform = rotationTransform.Multiply(scalingTransform);

        // Final transform includes translation
        Transform finalTransform = Transform.Identity;
        finalTransform.BasisX = srTransform.BasisX;
        finalTransform.BasisY = srTransform.BasisY;
        finalTransform.BasisZ = srTransform.BasisZ;

        // Set origin so that the transformed view center maps to the viewport center
        XYZ transformedViewCenter = srTransform.OfPoint(viewCenterInModel);
        finalTransform.Origin = viewportCenterOnSheet - transformedViewCenter;

        return finalTransform;
    }

    private Revision GetLatestRevision(Document doc)
    {
        // Collect all revisions in the document
        FilteredElementCollector revCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(Revision));

        // Filter revisions that are visible (not hidden)
        List<Revision> visibleRevisions = revCollector
            .Cast<Revision>()
            .Where(r => r.Visibility != RevisionVisibility.Hidden)
            .ToList();

        // If there are no visible revisions, return null
        if (!visibleRevisions.Any())
            return null;

        // Find the revision with the highest sequence number
        Revision latestRevision = visibleRevisions
            .OrderByDescending(r => r.SequenceNumber)
            .FirstOrDefault();

        return latestRevision;
    }
}
