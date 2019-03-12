using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ortogo.SolidWorks.ShelvingTask;
using System;

namespace ShelvingTaskTest
{
    [TestClass]
    public class ShelvingTaskTests
    {
        [TestMethod]
        public void TestShelvingCalculation()
        {
            TestSets.InitFirst();
            var c = Calculator.Calculate();

            Assert.AreNotEqual(false, c.success, "Calculation failed");

            Assert.AreEqual("�.2900.500.��.70.15", $"�.{GlobalScope.SH * 1000}.{GlobalScope.G*1000}.{c.supportType}");

            TestSets.InitSecond();
            c = Calculator.Calculate();
            Console.WriteLine(c.traversa);
            Assert.AreNotEqual(false, c.success, "Calculation failed");

            Assert.AreEqual("�.11400.1100.��.90.15", $"�.{GlobalScope.SH * 1000}.{GlobalScope.G*1000}.{c.supportType}");
        }
    }
}
