using Autodesk.Revit.DB;
using System;

namespace RevitConduitConnect.Models
{
    public class LinesBetweenConduit : IComparable<LinesBetweenConduit>
    {
        public XYZ StartLine { get; set; }

        public XYZ EndLine { get; set; }

        public double Size { get; set; }

        public int CompareTo(LinesBetweenConduit linesBetweenConduit)
        {
            return this.Size.CompareTo(linesBetweenConduit.Size);
        }
    }
}
