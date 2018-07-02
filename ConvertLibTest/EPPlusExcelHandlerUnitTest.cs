using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Xml;
using TransferLibrary;
using TransferModel;

namespace TransferLibTest
{
    [TestClass]
    public class EPPlusExcelHandlerUnitTest
    {
        private List<TestCase> _testCaseList;


        [TestInitialize]
        public void SetUp()
        {
            string filepath = @"E:\testsuite-deep.xml";

            XmlAnalysis xmlAnalysis = new XmlAnalysis(filepath);
            XmlToModel xtm = new XmlToModel(xmlAnalysis.GetAllTestCaseNodes());
            this._testCaseList = xtm.OutputTestCases();
        }

        [TestMethod]
        public void WriteExcelTest()
        {
            var eh = new EPPlusExcelHandler(this._testCaseList);
            eh.WriteExcel();

            Assert.AreEqual(0,0);
        }
    }
}