#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using RevitConduitConnect.Models;
using System;
using System.Collections.Generic;

#endregion

namespace RevitConduitConnect
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Transaction trans = new Transaction(doc);
            FilteredElementCollector collectorType = new FilteredElementCollector(doc).OfClass(typeof(ConduitType));
            FilteredElementCollector collectorLevel = new FilteredElementCollector(doc).OfClass(typeof(Level));
            ConduitType type = collectorType.FirstElement() as ConduitType;
            Level level = collectorLevel.FirstElement() as Level;

            List<Element> conduits = new List<Element>();

            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            foreach (var id in selectedIds)
            {
                conduits.Add(doc.GetElement(id));
            }

            var conduit1 = conduits[0] as Conduit;
            var conduit2 = conduits[1] as Conduit;

            XYZ StartAB = (conduit1.Location as LocationCurve).Curve.GetEndPoint(0);
            XYZ EndAB = (conduit1.Location as LocationCurve).Curve.GetEndPoint(1);

            XYZ StartCD = (conduit2.Location as LocationCurve).Curve.GetEndPoint(0);
            XYZ EndCD = (conduit2.Location as LocationCurve).Curve.GetEndPoint(1);

            LinesBetweenConduit listXYZ = SmallestDistanceConduit(StartAB, EndAB, StartCD, EndCD);
            Line lineComduit2 = Line.CreateBound(StartCD, EndCD);

            var locationCurve = conduit2.Location as LocationCurve;
            var line1End = (conduit2.Location as LocationCurve).Curve.GetEndParameter(1);

            trans.Start("changeTheConduit");

            lineComduit2.MakeBound(MissingDistance(listXYZ,StartAB,EndAB), line1End);
            locationCurve.Curve = lineComduit2;

            trans.Commit();

            StartCD = (conduit2.Location as LocationCurve).Curve.GetEndPoint(0);

            Line lineNewStartCD_OldEndCD = Line.CreateBound(StartCD, EndCD);

            listXYZ = SmallestDistanceConduit(StartAB, EndAB, StartCD, EndCD);

            XYZ StartBC = listXYZ.StartLine;
            XYZ EndBC = listXYZ.EndLine;

            trans = new Transaction(doc);
            trans.Start("createConduit");

            locationCurve.Curve = lineNewStartCD_OldEndCD;
            var connectingConduit = Conduit.Create(doc, type.Id, StartBC, EndBC, level.Id);
            connectingConduit.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(conduit1.Diameter);

            trans.Commit();

            trans = new Transaction(doc);
            trans.Start("createConduitConnect");

            Connect(StartBC, conduit1, connectingConduit, doc);
            Connect(EndBC, conduit2, connectingConduit, doc);

            trans.Commit();

            return Result.Succeeded;
        }

        private double MissingDistance(LinesBetweenConduit listXYZ, XYZ StartAB, XYZ EndAB)
        {
            Line lineConduit1 = Line.CreateBound(StartAB, EndAB);
            lineConduit1.MakeUnbound();
            var leg1 = lineConduit1.Distance(listXYZ.EndLine);
            var hypotenuse = listXYZ.StartLine.DistanceTo(listXYZ.EndLine);
            var leg2 = hypotenuse * hypotenuse - leg1 * leg1;

            leg2 = Math.Sqrt(leg2);

           return -(leg2 - leg1);
        }
        private LinesBetweenConduit SmallestDistanceConduit(XYZ startAB, XYZ endAB, XYZ startCD, XYZ endCD)
        {
            List<LinesBetweenConduit> lines = new List<LinesBetweenConduit>()
            {
               new LinesBetweenConduit{Size = startAB.DistanceTo(startCD), StartLine=startAB,EndLine=startCD},
               new LinesBetweenConduit{Size = startAB.DistanceTo(endCD), StartLine=startAB,EndLine=endCD},
               new LinesBetweenConduit{Size = endAB.DistanceTo(startCD), StartLine=endAB,EndLine=startCD},
               new LinesBetweenConduit{Size = endAB.DistanceTo(endCD), StartLine=endAB,EndLine=endCD}
            };
            lines.Sort();

            return lines[0];
        }

        private static void Connect(XYZ location, Element a, Element b, Document doc)
        {
            ConnectorManager cm = GetConnectorManager(a);

            if (null == cm)
            {
                throw new ArgumentException(
                  "Element a has no connectors.");
            }

            Connector ca = GetConnectorClosestTo(cm.Connectors, location);

            cm = GetConnectorManager(b);

            if (null == cm)
            {
                throw new ArgumentException(
                  "Element b has no connectors.");
            }

            Connector cb = GetConnectorClosestTo(cm.Connectors, location);
            doc.Create.NewElbowFitting(ca, cb);
        }

        private static ConnectorManager GetConnectorManager(Element e)
        {
            MEPCurve mc = e as MEPCurve;
            FamilyInstance fi = e as FamilyInstance;

            if (null == mc && null == fi)
            {
                throw new ArgumentException(
                  "Element is neither an MEP curve nor a fitting.");
            }

            return null == mc
              ? fi.MEPModel.ConnectorManager
              : mc.ConnectorManager;
        }

        private static Connector GetConnectorClosestTo(ConnectorSet connectors, XYZ location)
        {
            Connector targetConnector = null;
            double minDist = double.MaxValue;

            foreach (Connector c in connectors)
            {
                double d = c.Origin.DistanceTo(location);

                if (d < minDist)
                {
                    targetConnector = c;
                    minDist = d;
                }
            }
            return targetConnector;
        }
    }
}
