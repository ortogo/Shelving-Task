using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ortogo.SolidWorks.StillageTask;
using System;

namespace StillageTaskTest
{
    [TestClass]
    public class StillageTaskTests
    {
        [TestMethod]
        public void TestStillageCalculation()
        {
            TestSets.InitFirst();
            var c = Calculator.Calculate();

            Assert.AreNotEqual(false, c.success, "Calculation failed");

            Assert.AreEqual("п.2900.500.ях.70.15", $"п.{GlobalScope.SH * 1000}.{GlobalScope.G*1000}.{c.supportType}");

            TestSets.InitSecond();
            c = Calculator.Calculate();
            Console.WriteLine(c.traversa);
            Assert.AreNotEqual(false, c.success, "Calculation failed");

            Assert.AreEqual("п.11400.1100.ях.90.15", $"п.{GlobalScope.SH * 1000}.{GlobalScope.G*1000}.{c.supportType}");
        }
    }
}
