using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using ExStudentAddIn;

namespace ExStudentAddInTest
{
    [TestFixture]
    public class SheetHelperTest
    {
        [Test]
        public void TestGetColumnNumber()
        {
            string colName = "A";
            int colNumber = SheetHelper.GetColumnNumber(colName);
            Assert.AreEqual(colNumber, 1);
        }

        [Test]
        public void TestGetColumnNumberComplex()
        {
            string colName = "AB";
            int colNumber = SheetHelper.GetColumnNumber(colName);
            Assert.AreEqual(colNumber, 28);
        }

        [Test]
        public void TestGetColumnName()
        {
            int colNumber = 2;
            string colName = SheetHelper.GetColumnName(colNumber);
            Assert.AreEqual(colName, "B");
        }

        [Test]
        public void TestGetColumnNameComplex()
        {
            //string colName = "AB";
            int colNumber = 28;
            string colName = SheetHelper.GetColumnName(colNumber);
            Assert.AreEqual(colName, "AB");
        }
    }
}
