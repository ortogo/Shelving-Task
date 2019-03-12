using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ortogo.SolidWorks.ShelvingTask;

namespace ShelvingTaskTest
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
            Assert.AreEqual("ТП.1800.100.40.15", traversa.ToString(), "Traversa is not correct");

            // case if traversa not found
            TestSets.InitThird();
            traversa = new TraversaCut().Select();

            Assert.AreEqual(null, traversa, "Traversa is found");
        }
    }
}
