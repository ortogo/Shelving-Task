using Ortogo.SolidWorks.ShelvingTask;

namespace ShelvingTaskTest
{
    public class TestSets
    {
        public static void InitFirst()
        {
            GlobalScope.QI = 150;
            GlobalScope.NL = 8;
            GlobalScope.K = 0.3;
            GlobalScope.L = 1.050;
            GlobalScope.G = 0.5;
            GlobalScope.NS = 5;
        }

        public static void InitSecond()
        {
            GlobalScope.QI = 2100;
            GlobalScope.NL = 6;
            GlobalScope.K = 1.6;
            GlobalScope.L = 1.8;
            GlobalScope.G = 1.1;
            GlobalScope.NS = 5;
        }

        public static void InitThird()
        {
            GlobalScope.QI = 3000;
            GlobalScope.NL = 7;
            GlobalScope.K = 1.8;
            GlobalScope.L = 2.7;
            GlobalScope.G = 1.1;
            GlobalScope.NS = 5;
        }
    }
}
