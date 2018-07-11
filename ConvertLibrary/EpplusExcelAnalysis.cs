﻿using System;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using System.Collections.Generic;
using TransferModel;
using log4net;
using TransferLibrary;

namespace ConvertLibrary
{
    public class EPPlusExcelAnalysis
    {
        private readonly ILog _logger = LogManager.GetLogger(typeof(EPPlusExcelAnalysis));
        private ExcelPackage excelPackage;


        public EPPlusExcelAnalysis(string excelFilePath)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                OutputDisplay.ShowMessage("传入文件地址有误！", Color.Red);
                return;
            }

            if (!File.Exists(excelFilePath))
            {
                OutputDisplay.ShowMessage("文件不存在!", Color.Red);
                return;
            }

            try
            {
                FileInfo fiExcel = new System.IO.FileInfo(excelFilePath);
                this.excelPackage = new ExcelPackage(fiExcel);
            }catch (Exception ex)
            {
                OutputDisplay.ShowMessage(ex.Message, Color.Red);
                return;
            }
        }

        public Dictionary<string, List<TestCase>> ReadExcel()
        {
            Dictionary<string, List<TestCase>> dicAllTestCases = new Dictionary<string, List<TestCase>>();
            int iCount = this.excelPackage.Workbook.Worksheets.Count;

            if(iCount == 0)
            {
                OutputDisplay.ShowMessage("表中无Sheet页！", Color.Red);
                return null;
            }

            for(int iFlag = 1; iFlag <= iCount; iFlag++)
            {
                ExcelWorksheet excelWorksheet = this.excelPackage.Workbook.Worksheets[iFlag];
                var TestCase = this.GetExcelSheetData(excelWorksheet);

                if (dicAllTestCases.ContainsKey(excelWorksheet.Name))
                {
                    OutputDisplay.ShowMessage($"同一页签名:{excelWorksheet.Name}已在本Excel中出现过.",Color.GreenYellow);
                }

                if(TestCase.Count == 0)
                {
                    OutputDisplay.ShowMessage($"页签:{excelWorksheet.Name}无任何可转换用例数据.", Color.GreenYellow);
                    continue;
                }

                dicAllTestCases.Add(excelWorksheet.Name, TestCase);
            }
            this.excelPackage.Dispose();
            return dicAllTestCases;
        }

        public List<TestCase> GetExcelSheetData(ExcelWorksheet eWorksheet)
        {
            List<TestCase> tcList = new List<TestCase>();
            int usedRows, usedCols;

            if(eWorksheet.Dimension is null)
            {
                this._logger.Warn(new Exception("No TestCase, this Sheet is new!"));
                return new List<TestCase>();
            }
            else
            {
                usedRows = eWorksheet.Dimension.End.Row;
                usedCols = eWorksheet.Dimension.End.Column;
            }

            if(usedRows == 0 || usedRows == 1)
            {
                this._logger.Warn(new Exception("No TestCase!"));
                return tcList;
            }

            for(int i=1; i < eWorksheet.Dimension.End.Row; i++)
            {
                if(eWorksheet.Cells[i,1].Text != null || eWorksheet.Cells[i,1].Text.ToString() != string.Empty ||
                    !eWorksheet.Cells[i, 1].Text.ToString().Equals("END"))
                {
                    continue;
                }
                usedRows = i;
                break;
            }

            TestCase tc = new TestCase();

            for (int i = 2; i < usedRows; i++)
            {
                var currentCell = eWorksheet.Cells[i, 1];
                
                if (currentCell.Value is null)
                {
                    TestStep ts = new TestStep();
                    ts.StepNumber = tc.TestSteps.Count + 1;
                    ts.ExecutionType = ExecType.手动;
                    ts.Actions = eWorksheet.Cells[i, 7].Text.ToString();
                    ts.ExpectedResults = eWorksheet.Cells[i, 8].Text.ToString();

                    tc.TestSteps.Add(ts);
                    continue;
                }
                else
                {
                    if(tc.ExternalId != null)
                    {
                        tcList.Add(tc);
                    }
                    tc = new TestCase();

                    tc.ExternalId = string.Format($"{currentCell.Text.ToString()}{DateTime.Now.ToString("yyyyMMddhhmmss")}");

                    tc.Name = eWorksheet.Cells[i, 2].Text.ToString();

                    tc.Importance = CommonHelper.StrToImportanceType(eWorksheet.Cells[i, 3].Text.ToString());

                    tc.ExecutionType = CommonHelper.StrToExecType(eWorksheet.Cells[i, 4].Text.ToString());

                    tc.Summary = eWorksheet.Cells[i, 5].Text.ToString();

                    tc.Preconditions = eWorksheet.Cells[i, 6].Text.ToString();

                    TestStep ts_one = new TestStep();
                    ts_one.StepNumber = 1;
                    ts_one.ExecutionType = ExecType.手动;
                    ts_one.Actions = eWorksheet.Cells[i, 7].Text.ToString();
                    ts_one.ExpectedResults = eWorksheet.Cells[i, 8].Text.ToString();

                    tc.TestSteps = new List<TestStep>();
                    tc.TestSteps.Add(ts_one);
                }
            }

            return tcList;
        }


    }
}
