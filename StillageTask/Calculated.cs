namespace Ortogo.SolidWorks.StillageTask
{
    public struct Calculated
    {
        public Frame frame { get; set; }
        public TraversaCut.Traversa traversa { get; set; }
        public SupportElement.Type supportType { get; set; }
        public Frame.Connection hConn { get; set; }
        public Frame.Connection dConn { get; set; }

        public double countTraversa { get; set; }
        public double sumLenTraversa { get; set; }
        public int countConn { get; set; }
        public double sumLenConn { get; set; }

        public bool success;
    }
}
