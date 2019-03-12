using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ortogo.SolidWorks.StillageTask;

namespace StillageTaskTest
{
    [TestClass]
    public class TraversaCutTests
    {
        [TestMethod]
        public void TestSelect()
        {
            TestSets.InitFirst();
            var traversa = new TraversaCut().Select();

            Assert.AreNotEqual(null, traversa, "Traversa is not found");
            Assert.AreEqual("ТК.1050.40.40.15", traversa.ToString(), "Traversa is not correct");

            TestSets.InitSecond();
            traversa = new TraversaCut().Select();

            Assert.AreNotEqual(null, traversa, "Traversa is not found");
            Assert.AreEqual("ТК.1800.40.40.15", traversa.ToString(), "Traversa is not correct");

            TestSets.InitThird();
            traversa = new TraversaCut().Select();

            Assert.AreNotEqual(null, traversa, "Traversa is not found");
            Assert.AreEqual("ТК.2700.40.40.15", traversa.ToString(), "Traversa is not correct");
        }
    }
}
