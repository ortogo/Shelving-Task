using System;

namespace Ortogo.SolidWorks.ShelvingTask
{
    public class Frame
    {
        public class Connection
        {
            public string Type { get; set; }
            public double L1 { get; set; }
            public double LK { get; set; }
            public double BK { get; set; }
            public double mHK = 30; // mm
            public double mSK = 1.5;
            public int Count { get; set; }

            public override string ToString()
            {
                return $"{Type}.{LK}.{mHK}.{BK}.{mSK * 10}";
            }
        }

        public Connection GetConnection(string Type, SupportElement.Type support)
        {
            var connection = new Connection { Type = Type };
            if (Type.Equals("СГ"))
            {
                connection.L1 = GlobalScope.G - GlobalScope.K * 2;
            } else
            {
                connection.L1 = Math.Sqrt(Math.Pow(GlobalScope.G - support.K * 2,2) + Math.Pow(000, 2));
            }
            connection.LK = connection.L1 + 32;
            connection.BK = support.F / 2;
            connection.Count = GetCountOfConnections(Type);
            return connection;
        }

        public int GetCountOfConnections(string Type)
        {
            if (Type.Equals("СГ"))
            {
                return GlobalScope.SH > 1.2 ? 5 : 2;
            }
            return (int)(GlobalScope.K * GlobalScope.L / 6);
        }
    }
}
