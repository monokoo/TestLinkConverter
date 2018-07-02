using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ConvertLibrary;

namespace ConvertLibTest
{
    [TestClass]
    public class EPPlusExcelAnalysisTest
    {
        private EPPlusExcelAnalysis epplusExcelAnalysis;

        [TestInitialize]
        public void SetUp()
        {
            string filePath = @"E:\Github\TestLinkConverter\Resource\TestCase.xlsx";

            this.epplusExcelAnalysis = new EPPlusExcelAnalysis(filePath);
        }

        [TestMethod]
        public void ReadExcelTest()
        {
            this.epplusExcelAnalysis.ReadExcel();
        }
    }
}
