#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

#endregion

namespace RevitConduitConnect
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        private Line line;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Document doc = commandData.Application.ActiveUIDocument.Document;
            Transaction trans = new Transaction(doc);

            FilteredElementCollector collectorType = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.Electrical.ConduitType));
            FilteredElementCollector collectorLevel = new FilteredElementCollector(doc).OfClass(typeof(Level));

            Autodesk.Revit.DB.Electrical.ConduitType type = collectorType.FirstElement() as Autodesk.Revit.DB.Electrical.ConduitType;
            Level level = collectorLevel.FirstElement() as Level;

            ICollection<Element> allConduit = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements();

            List<Element> conduits = new List<Element>();

            foreach (var conduit in allConduit)
            {
                if (conduit.Id.IntegerValue == 1348848 || conduit.Id.IntegerValue == 1348872)
                {
                    conduits.Add(conduit);
                }
            }

            var element1 = conduits[0] as Conduit;
            var element2 = conduits[1] as Conduit;

            XYZ StartAB = (element1.Location as LocationCurve).Curve.GetEndPoint(0);
            XYZ EndAB = (element1.Location as LocationCurve).Curve.GetEndPoint(1);

            XYZ StartCD = (element2.Location as LocationCurve).Curve.GetEndPoint(0);
            XYZ EndCD = (element2.Location as LocationCurve).Curve.GetEndPoint(1);

            List<XYZ> result =  SmallestDistanceConduit(StartAB, EndAB, StartCD, EndCD);

            XYZ StartBC = result[0];
            XYZ EndBC = result[1];

            var locationCurve = element2.Location as LocationCurve;

            if (StartCD.X > EndCD.X && StartCD.X > EndBC.X)
                line = Line.CreateBound(StartCD, EndBC);

            else if (StartCD.X < EndCD.X && EndCD.X > EndBC.X)
                line = Line.CreateBound(EndCD, EndBC);

            else if (StartCD.X > EndCD.X && StartCD.X < EndBC.X)
                line = Line.CreateBound(EndCD, EndBC);

            else if (StartCD.X < EndCD.X && StartCD.X < EndBC.X)
                line = Line.CreateBound(StartCD, EndBC);
            
            trans.Start("createConduit");

            var connectingConduit = Conduit.Create(doc, type.Id, StartBC, EndBC, level.Id);
            locationCurve.Curve = line;
            trans.Commit();
           
            trans = new Transaction(doc);

            trans.Start("createConduitConnect");

            if (StartBC == StartAB)
            {
                Connect(StartAB, element1, connectingConduit, doc);
            }
            if (StartBC == EndAB)
            {
                Connect(EndAB, element1, connectingConduit, doc);
            }
            
            Connect(EndBC, connectingConduit, element2, doc);

            trans.Commit();

            return Result.Succeeded;
        }

        private List<XYZ> SmallestDistanceConduit(XYZ startAB, XYZ endAB, XYZ startCD, XYZ endCD)
        {
            var a = (Line.CreateBound(startAB, startCD)).ApproximateLength;
            var b = (Line.CreateBound(startAB, endCD)).ApproximateLength;
            var c = (Line.CreateBound(endAB, startCD)).ApproximateLength;
            var d = (Line.CreateBound(endAB, endCD)).ApproximateLength;

            List<(double, string)> list = new List<(double, string)>() { (a, "a"), (b, "b"), (c, "c"), (d, "d") };

            list.Sort();

            var result = list[0];

            if (result.Item2 == "a")
            {
                return CoordinatesSearch(startAB, startCD);
            }
            else if (result.Item2 == "b")
            {
                return CoordinatesSearch(startAB, endCD);
            }
            else if (result.Item2 == "c")
            {
                return CoordinatesSearch(endAB, startCD);
            }
            else if (result.Item2 == "d")
            {
                return CoordinatesSearch(endAB, endCD);
            }
            else
                return null;
        }

        private List<XYZ> CoordinatesSearch(XYZ StartBC, XYZ EndBC)
        {
            List<XYZ> XYZ = new List<XYZ>();

            if(StartBC.X>EndBC.X && StartBC.Y>EndBC.Y)
            {
                XYZ.Add(StartBC);

                var f = EndBC.Y - (EndBC.X + StartBC.Y);

                XYZ.Add(new XYZ( (EndBC.X + StartBC.X) + f,( EndBC.X + StartBC.Y)+ f, StartBC.Z));
            }
            if (StartBC.X > EndBC.X && StartBC.Y < EndBC.Y)
            {
                XYZ.Add(StartBC);

                var f =  (StartBC.Y - EndBC.X)- EndBC.Y;

                XYZ.Add(new XYZ((StartBC.X + EndBC.X)+f, (StartBC.Y - EndBC.X)-f, StartBC.Z));
            }
            if (StartBC.X < EndBC.X && StartBC.Y > EndBC.Y)
            {
                XYZ.Add(StartBC);

                var f =EndBC.Y - (StartBC.Y - EndBC.X);

                XYZ.Add(new XYZ((StartBC.X + EndBC.X)-f, (StartBC.Y - EndBC.X)+f, StartBC.Z));
            }
            if (StartBC.X < EndBC.X && StartBC.Y < EndBC.Y)
            {
                XYZ.Add(StartBC);

                var f = (StartBC.Y + EndBC.X) - EndBC.Y;

                XYZ.Add(new XYZ((StartBC.X + EndBC.X)-f, (StartBC.Y + EndBC.X)-f, StartBC.Z));
            }
            return XYZ;

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
