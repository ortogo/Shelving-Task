namespace Ortogo.SolidWorks.ShelvingTask
{
    public class GlobalScope
    {
        public static double R = 235; // MPa
        public static double KoefZap = 1;
        public static double Ge = 9.81; // m/s2
        public static double ModuleJunga = 210 * 10E9; // Pa
        public static double Beta1 = 0.8;

        public static double TUSHM = 0.002; // m 

        public static double QI { get; set; }
        public static double NL { get; set; }
        public static double K { get; set; }
        public static double L { get; set; }
        public static double G { get; set; }
        public static double NS { get; set; }

        public static double SH { get => K * (NL + 1) + 0.2; }
        public static double SW { get => NS * L; }
    }
}
