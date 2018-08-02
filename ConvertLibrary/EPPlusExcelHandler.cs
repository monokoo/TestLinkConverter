using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Threading;
using ConvertLibrary;
using log4net;
using TransferModel;
using OfficeOpenXml;

namespace TransferLibrary
{
    public class EPPlusExcelHandler
    {
        private readonly ILog _logger = LogManager.GetLogger(typeof (EPPlusExcelHandler));
        private readonly List<TestCase> _sourceTestCases;

        public EPPlusExcelHandler(List<TestCase> outputCases)
        {
            this._sourceTestCases = outputCases;
        }

        /// <summary>
        /// 写Excel
        /// </summary>
        public void WriteExcel()
        {
            string currentDir = System.Environment.CurrentDirectory;
            string fileName = $"{currentDir}\\TestCaseTemplate.xlsx";

            if (!System.IO.File.Exists(fileName))
            {
                string message = $"{fileName}文件已不存在.";
                this._logger.Error(new Exception(message));
                throw new Exception(message);
            }

            ExcelPackage excelPackage = new ExcelPackage(new System.IO.FileInfo(fileName));
            ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets[1];
            this.WriteInWorkSheet(workSheet);

            string saveDir = fileName.Replace("TestCaseTemplate.xlsx", $"TestCase_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");
            using (System.IO.Stream stream = new System.IO.FileStream(saveDir, System.IO.FileMode.Create))
            {
                excelPackage.SaveAs(stream);
            }
            excelPackage.Dispose();
        }

        private void WriteInWorkSheet(ExcelWorksheet workSheet)
        {
            int iFlag = 2;
            foreach(TestCase node in this._sourceTestCases)
            {
                OutputDisplay.ShowMessage(node.Name, Color.Chartreuse);
                workSheet.Cells[iFlag, 1].Value = node.ExternalId;
                workSheet.Cells[iFlag, 2].Value = node.Name;
                workSheet.Cells[iFlag, 3].Value = node.Importance.ToString();
                workSheet.Cells[iFlag, 4].Value = node.ExecutionType.ToString();
                workSheet.Cells[iFlag, 5].Value = node.Summary;
                workSheet.Cells[iFlag, 6].Value = node.Preconditions;
                int iMerge = 0;
                if(node.TestSteps == null)
                {
                    workSheet.Cells[iFlag, 7].Value = string.Empty;
                    workSheet.Cells[iFlag, 8].Value = string.Empty;
                    iFlag++;
                }
                else
                {
                    foreach (TestStep step in node.TestSteps)
                    {
                        workSheet.Cells[iFlag, 7].Value = CommonHelper.DelTags(step.Actions);
                        workSheet.Cells[iFlag, 8].Value = CommonHelper.DelTags(step.ExpectedResults);
                        iFlag++;
                        iMerge++;
                    }
                    this.MergeCells(workSheet, iMerge, iFlag - iMerge);
                }
                Thread.Sleep(1000);
            }
            workSheet.Cells[iFlag++, 1].Value = "END";
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="workSheet">指定Sheet页</param>
        /// <param name="iMerge"></param>
        /// <param name="iFlag"></param>
        private void MergeCells(ExcelWorksheet workSheet, int iMerge, int iFlag)
        {
            //导出Excel前6列均需要合并单元格
            for(int i=1; i<=6; i++)
            {
                var startCell = workSheet.Cells[iFlag, i];
                var endCell = workSheet.Cells[iFlag + iMerge - 1, i];
                string addressStr = string.Format("{0}:{1}", startCell.Address.ToString(), endCell.Address.ToString());
                using (ExcelRange er = workSheet.Cells[addressStr])
                {
                    er.Merge = true;
                }
                //workSheet.Cells[] 
                //ExcelRange rangeLecture = workSheet.SelectedRange.Range[workSheet.Cells[iFlag, i], workSheet.Cells[iFlag + iMerge - 1, i]];
            }
        }
    }
}