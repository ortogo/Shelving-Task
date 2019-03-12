using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ortogo.SolidWorks.StillageTask
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
            };
        }

        public Traversa Select()
        {
            var qy = (GlobalScope.QI * GlobalScope.Ge) / (GlobalScope.L * 2);
            var My = (qy * Math.Pow(GlobalScope.L, 2)) / 8;
            var WYc = My / (GlobalScope.R*10E6 * GlobalScope.KoefZap);

            var JYc = (200 * 5 * qy * Math.Pow(GlobalScope.L, 4)) / (384 * GlobalScope.ModuleJunga);

            foreach (var trav in TypeSizes)
            {
                if ((trav.WY * 10E-6) >= WYc && (trav.JY * 10E-6) >= JYc)
                {
                    trav.Length = GlobalScope.L;
                    trav.QYC = qy;
                    trav.MYC = My;
                    trav.WYC = WYc;
                    trav.JYC = JYc;
                    trav.Sigma = My / (trav.WY * 10E-6);
                    trav.Fc = (5 * qy * Math.Pow(GlobalScope.L, 4)) / (384* GlobalScope.ModuleJunga * (trav.JY * 10E-6));
                    return trav;
                }
            }
            return null;
        }
    }
}
