using System;
using System.Collections.Generic;

namespace Ortogo.SolidWorks.StillageTask
{
    class SupportElement
    {
        public class Type
        {
            public string Mark { get; set; }
            public double A { get; set; }
            public double S { get; set; }
            public double Aeff { get; set; }
            public double IY { get; set; }
            public double IZ { get; set; }
            public double B { get; set; }
            public double K { get; set; }
            public double F { get; set; }

            public override string ToString()
            {
                return $"{Mark}.{A}.{S * 10}";
            }
        }

        public List<Type> SupportTypes { get; set; }

        public SupportElement()
        {
            SupportTypes = new List<Type>
            {
                new Type
                {
                    Mark = "СИ",
                    A = 70,
                    S = 1.5,
                    Aeff = 3.25,
                    IY = 2.13,
                    IZ = 2.41,
                    B = 70,
                    K = 55,
                    F = 30,
                },
                new Type
                {
                    Mark = "СИ",
                    A = 90,
                    S = 1.5,
                    Aeff = 3.85,
                    IY = 2.36,
                    IZ = 3.15,
                    B = 80,
                    K = 65,
                    F = 50,
                },
                new Type
                {
                    Mark = "СИ",
                    A = 90,
                    S = 2,
                    Aeff = 5.09,
                    IY = 2.34,
                    IZ = 3.13,
                    B = 80,
                    K = 65,
                    F = 50,
                },
                new Type
                {
                    Mark = "СИ",
                    A = 110,
                    S = 2,
                    Aeff = 6.10,
                    IY = 2.57,
                    IZ = 3.67,
                    B = 90,
                    K = 75,
                    F = 50,
                },
            };
        }

        public Type Select()
        {
            var Nsd = (GlobalScope.QI * GlobalScope.Ge) * GlobalScope.NL / 2;
            var lambda = Math.PI * Math.Sqrt(GlobalScope.ModuleJunga / (GlobalScope.R * 10E6));

            foreach (var item in SupportTypes)
            {
                var lambda_y = GlobalScope.K / (item.IY * 10E2);
                var lambda_z = (GlobalScope.K <= 1.2 ? 0.6 : 1.2) / (item.IZ * 10E2);
                var lambda_ys = lambda_y * Math.Sqrt(GlobalScope.Beta1) / lambda;
                var lambda_zs = lambda_z * Math.Sqrt(GlobalScope.Beta1) / lambda;
                var phi_y = 0.5 * (1 + 0.34 * (lambda_ys - 0.2) + Math.Pow(lambda_ys, 2));
                var phi_z = 0.5 * (1 + 0.34 * (lambda_zs - 0.2) + Math.Pow(lambda_zs, 2));

                var Xi_y = 1 / (phi_y + Math.Sqrt(Math.Pow(phi_y, 2) - Math.Pow(lambda_ys, 2)));
                //if (Xi_y > 1) throw new Exception("Xi_y must less than 1");
                var Xi_z = 1 / (phi_z + Math.Sqrt(Math.Pow(phi_z, 2) - Math.Pow(lambda_zs, 2)));
                //if (Xi_z > 1) throw new Exception("Xi_z must less than 1");

                var Xi = Math.Min(Xi_y, Xi_z);
                var Nd = Xi * item.Aeff * 10E4 * GlobalScope.R * 10E6 / GlobalScope.KoefZap;
                if (Nd >= Nsd)
                {
                    Console.WriteLine($"found support cut {item.A}");
                    return item;
                }
            }

            return null;
        }
    }
}
