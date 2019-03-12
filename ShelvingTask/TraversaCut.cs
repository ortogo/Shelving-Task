using System;
using System.Collections.Generic;

namespace Ortogo.SolidWorks.ShelvingTask
{
    public class TraversaCut
    {
        public class Traversa
        {
            public string Mark { get; set; }
            public double Length { get; set; } // set on select
            public double HSize { get; set; }
            public double BSize { get; set; }
            public double SSize { get; set; }
            public double WY { get; set; }
            public double JY { get; set; }
            public double QYC { get; set; } // c - calculated
            public double MYC { get; set; }
            public double Fc { get; set; }
            public double WYC { get; set; }
            public double JYC { get; set; }
            public double Sigma { get; set; }

            public override string ToString()
            {
                return $"{Mark}.{Length*1000}.{HSize}.{BSize}.{SSize}";
            }
        }

        public List<Traversa> TypeSizes { get; set; }
        
        public TraversaCut ()
        {
            TypeSizes = new List<Traversa>
            {
                new Traversa
                {
                    Mark = "ТК",
                    Length = 100,
                    HSize = 40,
                    BSize = 40,
                    SSize = 15,
                    WY = 2.75,
                    JY =5.49,
                },
                new Traversa
                {
                    Mark = "ТК",
                    Length = 40,
                    HSize = 40,
                    BSize = 40,
                    SSize = 20,
                    WY = 3.54,
                    JY = 7.07,
                },
                new Traversa
                {
                    Mark = "ТП",
                    Length = 1800,
                    HSize = 100,
                    BSize = 40,
                    SSize = 15,
                    WY = 10.10,
                    JY = 50.49,
                },
            };
        }

        public Traversa Select()
        {
            var qy = (GlobalScope.QI * GlobalScope.Ge) / (GlobalScope.L * 2); //N/m
            var My = (qy * Math.Pow(GlobalScope.L, 2)) / 8;

            

            foreach (var trav in TypeSizes)
            {
                var sigma = My / trav.WY;
                var f = (5 / 384) * ((qy * Math.Pow(GlobalScope.L, 4)) / (GlobalScope.ModuleJunga * trav.JY));
                if (sigma <= (GlobalScope.R * GlobalScope.KoefZap)
                    && f <= (GlobalScope.L/200))
                {
                    trav.Length = GlobalScope.L;
                    trav.QYC = qy;
                    trav.MYC = My;
                    trav.Sigma = sigma;
                    trav.Fc = f;
                    return trav;
                }
            }
            return null;
        }
    }
}
