using System.Xml;
using OfficeOpenXml;

namespace Trx2Xlsx
{
    /// <summary>
    /// Converts TRX (Test Results XML) to XLSX (Excel) format.
    /// </summary>
    public static class Trx2XlsxConverter
    {
        /// <summary>
        /// Main entry point for the Trx2XlsxConverter application.
        /// </summary>
        /// <param name="args">Command-line arguments: [inputFileName.trx] [outputFileName.xlsx]</param>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: trx2xlsx <inputFileName.trx> <outputFileName.xlsx>");
                return;
            }

            string inputFilePath = args[0];
            string outputFilePath = args[1];

            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            ConvertTrxToXlsx(inputFilePath, outputFilePath);

            Console.WriteLine($"Conversion completed. Output saved to: {outputFilePath}");
        }

        /// <summary>
        /// Converts TRX file to XLSX format.
        /// </summary>
        /// <param name="inputFilePath">Path to the input TRX file.</param>
        /// <param name="outputFilePath">Path to save the output XLSX file.</param>
        static void ConvertTrxToXlsx(string inputFilePath, string outputFilePath)
        {
            var package = new ExcelPackage(); 
            var worksheet = package.Workbook.Worksheets.Add("TestResults");

            XmlDocument xmlDocument = new();
            xmlDocument.Load(inputFilePath);

            var nsmgr = new XmlNamespaceManager(xmlDocument.NameTable);
            var xmlNameSpaceString = "http://microsoft.com/schemas/VisualStudio/TeamTest/2010";
            nsmgr.AddNamespace("ns", xmlNameSpaceString);

            XmlNodeList? testResults = xmlDocument.SelectNodes("/ns:TestRun/ns:Results/ns:UnitTestResult", nsmgr);

            if (testResults == null || testResults.Count == 0)
            {
                Console.WriteLine("No test results found.");
                return;
            }   

            worksheet.Cells[1, 1].Value = "TestName";
            worksheet.Cells[1, 2].Value = "TestScenario";
            worksheet.Cells[1, 3].Value = "Outcome";
            worksheet.Cells[1, 4].Value = "Duration";
            worksheet.Cells[1, 5].Value = "StartTime";
            worksheet.Cells[1, 6].Value = "EndTime";
            worksheet.Cells[1, 7].Value = "ErrorMessage";

            int rowIndex = 2;

            worksheet.View.FreezePanes(rowIndex, 1);

            foreach (XmlNode test in testResults)
            {
                worksheet.Cells[rowIndex, 1].Value = test.Attributes?["testName"]?.Value;
                worksheet.Cells[rowIndex, 2].Value = test.SelectSingleNode("ns:Output/ns:StdOut", nsmgr)?.InnerText;

                var outcomeCell = worksheet.Cells[rowIndex, 3];
                outcomeCell.Value = test.Attributes?["outcome"]?.Value;
                if (test.Attributes?["outcome"]?.Value == "Passed")
                {
                    outcomeCell.Style.Font.Color.SetColor(System.Drawing.Color.Green);
                }
                else if (test.Attributes?["outcome"]?.Value == "Failed")
                {
                    outcomeCell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                }
                else if (test.Attributes?["outcome"]?.Value == "NotExecuted")
                {
                    outcomeCell.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                }

                worksheet.Cells[rowIndex, 4].Value = test.Attributes?["duration"]?.Value;
                worksheet.Cells[rowIndex, 5].Value = test.Attributes?["startTime"]?.Value;
                worksheet.Cells[rowIndex, 6].Value = test.Attributes?["endTime"]?.Value;

                XmlNode? errorInfoNode = test.SelectSingleNode("ns:Output/ns:ErrorInfo", nsmgr);
                if (errorInfoNode != null)
                {
                    worksheet.Cells[rowIndex, 7].Value = errorInfoNode.SelectSingleNode("ns:Message", nsmgr)?.InnerText;
                }

                rowIndex++;
            }

            var allRows = worksheet.Cells[1, 1, rowIndex - 1, 7];
            allRows.AutoFilter = true;

            using (var range = worksheet.Cells[1, 1, rowIndex, 7])
            {
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            int maxColumnWidth = 35;

            for (int i = 1; i <= 7; i++)
            {
                worksheet.Column(i).AutoFit();
                worksheet.Column(i).Width = Math.Min(worksheet.Column(i).Width, maxColumnWidth);
            }

            package.SaveAs(new FileInfo(outputFilePath));
        }
    }
}
