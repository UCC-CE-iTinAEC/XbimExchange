using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Xbim.Common.Logging;
using Xbim.COBieLiteUK;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Xbim.CobieLiteUK.Validation.Reporting
{
    /// <summary>
    /// Can create an Excel report containing summary and detailed reports on provided and missing information.
    /// Use the Create function to produce the report.
    /// </summary>
    public class ExcelValidationReport
    {
        internal static readonly ILogger Logger = LoggerFactory.GetLogger();

        /// <summary>
        /// Determines the format to be saved.
        /// </summary>
        public enum SpreadSheetFormat
        {
            ///// <summary>
            ///// Excel Binary File Format
            ///// </summary>
            //Xls,
            /// <summary>
            /// Excel xml File Format
            /// </summary>
            Xlsx
        }

        /// <summary>
        /// Creates the report in file format
        /// </summary>
        /// <param name="facility">the result of a DPoW validation to be transformed into report form.</param>
        /// <param name="suggestedFilename">target file for the spreadsheet (warning, the extension is automatically determined depending on the format)</param>
        /// <param name="format">determines the excel format to use</param>
        /// <returns>true if successful, errors are cought and passed to Logger</returns>
        public bool Create(Facility facility, string suggestedFilename, SpreadSheetFormat format)
        {
            var ssFileName = Path.ChangeExtension(suggestedFilename, format.ToString());
            if (File.Exists(ssFileName))
            {
                File.Delete(ssFileName);
            }
            try
            {
                using (var spreadsheetStream = new FileStream(ssFileName, FileMode.Create, FileAccess.Write))
                {
                    //var result = Create(facility, spreadsheetStream, format);
                    var result = CreateSpreadsheet(facility, spreadsheetStream);
                    spreadsheetStream.Close();
                    return result;
                }
            }
            catch (Exception e)
            {
                Logger.ErrorFormat("Failed to save {0}, {1}", ssFileName, e.Message);
                return false;
            }
        }

        /// <summary>
        /// Creates the report.
        /// </summary>
        /// <param name="facility">the result of a DPoW validation to be transformed into report form.</param>
        /// <param name="filename">target file for the spreadsheet</param>
        /// <returns>true if successful, errors are cought and passed to Logger</returns>
        public bool Create(Facility facility, String filename)
        {
            if (filename == null)
                return false;
            SpreadSheetFormat format;
            var ext = Path.GetExtension(filename).ToLowerInvariant();
            if (ext != "xlsx")
                format = SpreadSheetFormat.Xlsx;
            //else if (ext != "xls")
            //    format = SpreadSheetFormat.Xls;
            else
                return false;
            return Create(facility, filename, format);
        }

        /// <summary>
        /// Used to create links to details reports, before the reports are created
        /// </summary>
        private Dictionary<string, string> LinksToDetailsSheets = new Dictionary<string, string>();

       /// <summary>
       /// 
       /// </summary>
       /// <param name="facility"></param>
       /// <param name="destinationStream"></param>
       /// <returns></returns>
        public bool CreateSpreadsheet(Facility facility, Stream destinationStream)
        {
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(destinationStream))
                {
                    //set the worksheet properties and add a default sheet in it
                    SetWorkbookProperties(excelPackage);


                    var facReport = new FacilityReport(facility);
                    var iRunningWorkBook = 1;
                    foreach (var assetType in facReport.AssetRequirementGroups)
                    {
                        // only report items with any assets submitted (a different report should probably be provided otherwise)
                        if (assetType.GetSubmittedAssetsCount() < 1)
                            continue;

                        var firstOrDefault = assetType.RequirementCategories.FirstOrDefault(cat => cat.Classification == @"Uniclass2015");
                        if (firstOrDefault == null)
                            continue;
                        var tName = firstOrDefault.Code;
                        var validName = String.Format("{0} {1}", iRunningWorkBook++, CreateSafeSheetName(tName));
                        LinksToDetailsSheets.Add(tName, validName);
                    }


                    if (!CreateSummarySheet(excelPackage, facility))
                        return false;

                    if (facility.Documents != null)
                    {
                        if (!CreateDocumentDetailsSheet(excelPackage, facility.Documents))
                            return false;
                    }

                    iRunningWorkBook = 1;
                    foreach (var assetType in facReport.AssetRequirementGroups)
                    {
                        // only report items with any assets submitted (a different report should probably be provided otherwise)
                        if (assetType.GetSubmittedAssetsCount() < 1)
                            continue;

                        var firstOrDefault = assetType.RequirementCategories.FirstOrDefault(cat => cat.Classification == @"Uniclass2015");
                        if (firstOrDefault == null)
                            continue;
                        var tName = firstOrDefault.Code;

                        var validName = String.Format("{0} {1}", iRunningWorkBook++, CreateSafeSheetName(tName));

                        if (!CreateDetailSheet(excelPackage, assetType, validName))
                            return false;
                    }

                    // reports on Zones details

                    // ReSharper disable once LoopCanBeConvertedToQuery // might restore once code is stable
                    foreach (var zoneGroup in facReport.ZoneRequirementGroups)
                    {
                        // only report items with any assets submitted (a different report should probably be provided otherwise)
                        if (zoneGroup.GetSubmittedAssetsCount() < 1)
                            continue;
                        var firstOrDefault = zoneGroup.RequirementCategories.FirstOrDefault(cat => cat.Classification == @"Uniclass2015");
                        if (firstOrDefault == null)
                            continue;
                        var tName = firstOrDefault.Code;
                        var validName = String.Format("{0} {1}", iRunningWorkBook++, CreateSafeSheetName(tName));

                        if (!CreateDetailSheet(excelPackage, zoneGroup, validName))
                            return false;
                        //UNCOMMENT
                    }


                    excelPackage.SaveAs(destinationStream);
                }
            }
            catch (Exception e)
            {
                Logger.ErrorFormat("Failed to save {0}, {1}", "spreadsheet", e.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// e.g Uniclass2015
        /// </summary>
        public string PreferredClassification = "Uniclass2015";

        private bool CreateSummarySheet(ExcelPackage excelPackageIn, Facility facilityIn)
        {
            try
            {
                ExcelWorksheet summarySheet = AddWorkSheet(excelPackageIn, "Summary");

                // Set first rowIndex (after image) and colIndex (leave column1 blank)
                int rowIndex = 8;
                int colIndex = 2;

                string workSheetHeader = String.Format("{0} - Verification report - {1}", facilityIn.Project.Name, DateTime.Now.ToShortDateString());
                AddWorkSheetHeader(summarySheet, ref rowIndex, colIndex, workSheetHeader, 22);

                rowIndex += 2;
                if (facilityIn.AssetTypes != null && facilityIn.AssetTypes.Any())
                {
                    DataTable assetTypesReport = new GroupingObjectSummaryReport<CobieObject>(facilityIn.AssetTypes, "Asset types report").GetReport(PreferredClassification);
                    WriteReportToPage(summarySheet, ref rowIndex, ref colIndex, assetTypesReport, "Asset types report");
                }
                if (facilityIn.Zones != null && facilityIn.Zones.Any())
                {
                    DataTable zonesReport = new GroupingObjectSummaryReport<CobieObject>(facilityIn.Zones, "Zones report").GetReport(PreferredClassification);
                    WriteReportToPage(summarySheet, ref rowIndex, ref colIndex, zonesReport, "Zones report");
                }

                if (facilityIn.Documents != null && facilityIn.Documents.Any())
                {
                    DataTable docReport = new DocumentsReport(facilityIn.Documents).GetReport("ResponsibleRole");
                    WriteReportToPage(summarySheet, ref rowIndex, ref colIndex, docReport, "Documents verification report");
                }

                // set column width
                summarySheet.Column(2).Width = 60;

                return true;
            }
            catch (Exception e)
            {
                //log the error
                Logger.Error("Failed to create Summary Sheet", e);
                return false;
            }
        }

        private bool CreateDocumentDetailsSheet(ExcelPackage excelPackageIn, List<Document> list)
        {
            try
            {
                ExcelWorksheet documentsWorkSheet = AddWorkSheet(excelPackageIn, "Documents");

                // Set first rowIndex (after image) and colIndex (leave column1 blank)
                int rowIndex = 8;
                int colIndex = 2;

                string workSheetHeader = "Documents Report";
                AddWorkSheetHeader(documentsWorkSheet, ref rowIndex, colIndex, workSheetHeader, 22);
                rowIndex += 2;

                DataTable drep = new DocumentsDetailedReport(list).AttributesGrid;
                WriteReportToPage(documentsWorkSheet, ref rowIndex, ref colIndex, drep, "Detailed Documents report");

                return true;
            }
            catch (Exception e)
            {
                //log the error
                Logger.Error("Failed to create Summary Sheet", e);
                return false;
            }
        }

        private bool CreateDetailSheet(ExcelPackage excelPackageIn, TwoLevelRequirementPointer<AssetType, Asset> requirementPointer, string sheetName)
        {
            try
            {
                ExcelWorksheet detailsWorkSheet = AddWorkSheet(excelPackageIn, sheetName);

                // Set first rowIndex (after image) and colIndex (leave column1 blank)
                int rowIndex = 8;
                int colIndex = 2;

                string workSheetHeader = "Asset Type and assets report";
                AddWorkSheetHeader(detailsWorkSheet, ref rowIndex, colIndex, workSheetHeader, 22);
                rowIndex += 2;

                var rep = new TwoLevelDetailedGridReport<AssetType, Asset>(requirementPointer);
                rep.PrepareReport();               

                var table = rep.AttributesGrid;
                WriteReportToPage(detailsWorkSheet, ref rowIndex, ref colIndex, table, workSheetHeader);
                return true;
            }
            catch (Exception e)
            {
                //log the error
                Logger.Error("Failed to create detail Sheet", e);
                return false;
            }
        }


        private bool CreateDetailSheet(ExcelPackage excelPackageIn, TwoLevelRequirementPointer<Zone, Space> requirementPointer, string sheetName)
        {
            try
            {
                ExcelWorksheet detailsWorkSheet = AddWorkSheet(excelPackageIn, sheetName);

                // Set first rowIndex (after image) and colIndex (leave column1 blank)
                int rowIndex = 8;
                int colIndex = 2;

                string workSheetHeader = "Zone and spaces report";
                AddWorkSheetHeader(detailsWorkSheet, ref rowIndex, colIndex, workSheetHeader, 22);
                rowIndex += 2;

                var rep = new TwoLevelDetailedGridReport<Zone, Space>(requirementPointer);
                rep.PrepareReport();

                var table = rep.AttributesGrid;
                WriteReportToPage(detailsWorkSheet, ref rowIndex, ref colIndex, table, workSheetHeader);
                return true;
            }
            catch (Exception e)
            {
                //log the error
                Logger.Error("Failed to create detail Sheet", e);
                return false;
            }
        }
        
        private string CreateSafeSheetName(string nameProposal) 
        {
            char replaceChar = ' ';
            if (String.IsNullOrWhiteSpace(nameProposal))
            {
                return "null";
            }
            
            if (nameProposal.Length < 1) 
            {
                return "empty";
            }

            int length = Math.Min(31, nameProposal.Length);

            string shortenname = nameProposal.Substring(0, length);

            var result = shortenname.ToString();

            IEnumerable<char> badChars = new List<char>{
                '\u0000',
                '\u0003',
                ':',
                '/',
                '\\',
                '?',
                '*',
                ']',
                '[',
                '\''};

            foreach (char badChar in badChars)
            {
                result = result.Replace(badChar.ToString(), replaceChar.ToString());
            }
            
        return result;
        }

        private void WriteReportToPage(ExcelWorksheet excelWorkSheet, ref int rowIndex, ref int colIndex, DataTable report, string reportTitle)
        {           
            // Output report data
            if (report.Rows.Count > 0)
            {
                int col = colIndex;

                if (!String.IsNullOrWhiteSpace(reportTitle))
                {
                    AddWorkSheetHeader(excelWorkSheet, ref rowIndex, col, reportTitle, 14);

                    if (reportTitle == "Documents verification report")
                    {
                        rowIndex++;
                        SetHyperlinkToWorksheet(excelWorkSheet.Cells[rowIndex, colIndex], "Documents", "A1", "Go to report");
                    }
                    rowIndex += 2;
                }

                //Creating Headings
                foreach (DataColumn dataCol in report.Columns)
                {
                    if (dataCol == report.Columns[0]) continue;

                    var cell = excelWorkSheet.Cells[rowIndex, col];

                    //Setting Value in cell
                    cell.Value = dataCol.Caption;
                    FormatTableHeaderCell(cell);

                    col++;
                }

                // Reset column index
                col = colIndex;

                int fromRow = rowIndex + 1;

                foreach (DataRow row in report.Rows)
                {
                    //excelRow = summaryPage.Row(startingRow);
                    rowIndex++;
                    col = 1;

                    if (reportTitle == "Asset types report")
                    {
                        if (LinksToDetailsSheets.ContainsKey((string)row[col]))
                        {
                            string sheetName;
                            LinksToDetailsSheets.TryGetValue((string)row[col], out sheetName);
                            SetHyperlinkToWorksheet(excelWorkSheet.Cells[rowIndex, colIndex], sheetName, "A1", (string)row[col]);
                        }
                    }

                    var writer = new ExcelCellVisualValue(excelWorkSheet);
                    foreach (DataColumn tCol in report.Columns)
                    {
                        if (tCol.AutoIncrement)
                            continue;
                        col++;
                        if (row[tCol] == DBNull.Value)
                            continue;
                        ExcelRange cell = excelWorkSheet.Cells[rowIndex, col];

                        // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                        if (row[tCol] is IVisualValue)
                        {
                            writer.SetCell(cell, (IVisualValue)row[tCol]);
                        }
                        else
                        {
                            switch (tCol.DataType.Name)
                            {
                                case "String":
                                    cell.Value = (string)row[tCol];
                                    break;
                                case "Int32":
                                    cell.Value = (Convert.ToInt32(row[tCol]));
                                    break;
                                default:
                                    cell.Value = ((string)row[tCol]);
                                    break;
                            }
                        }
                    }
                }

                int toRow = rowIndex;
                int toCol = report.Columns.Count;

                FormatTableCellBorders(excelWorkSheet, fromRow, colIndex, toRow, toCol);
                rowIndex += 3;

                SetColumnWidths(excelWorkSheet);
            }
        }

        private void FormatTableHeaderCell(ExcelRange cell)
        {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(243, 245, 244));

            //Set borders.
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            cell.Style.Border.Bottom.Color.SetColor(Color.FromArgb(62, 177, 200));
        }

        private void SetHyperlinkToWorksheet(ExcelRange cell, string targetSheetName, string targetCell, string displayText)
        {
            targetSheetName = String.Format("'{0}'", targetSheetName);
            cell.Hyperlink = new ExcelHyperLink(String.Format("{0}!{1}",targetSheetName, targetCell), displayText);
            SetHyperlinkFormat(cell);
        }
        private void SetHyperlinkFormat(ExcelRange cell)
        {
            cell.Style.Font.Color.SetColor(Color.FromArgb(0, 125, 158));
            cell.Style.Font.Size = 11;
            cell.Style.Font.UnderLine = true;
            cell.Style.Font.Name = "Calibri";

        }

        private void FormatTableCellBorders(ExcelWorksheet workSheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            workSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Color.SetColor(Color.FromArgb(242, 242, 242));
        }

        private void SetWorkbookProperties(ExcelPackage excelPackageIn)
        {
            excelPackageIn.Workbook.Properties.Author = "Xbim Cobie Lite UK";
            excelPackageIn.Workbook.Properties.Title = "Xbim Cobie Lite UK Validation";
        }

        private ExcelWorksheet AddWorkSheet(ExcelPackage excelPackageIn, string sheetNameIn)
        {
            excelPackageIn.Workbook.Worksheets.Add(sheetNameIn);
            ExcelWorksheet workSheet = excelPackageIn.Workbook.Worksheets[sheetNameIn];
            workSheet.Name = sheetNameIn;
            workSheet.Cells.Style.Font.Size = 11;
            workSheet.Cells.Style.Font.Name = "Arial";
            //workSheet.Column(1).Width = 3;
            workSheet.View.ShowGridLines = false;

            AddImageToWorksheet(workSheet, 1, 1, Xbim.CobieLiteUK.Validation.Properties.Resources.btk_logo_beta);

            return workSheet;
        }

        private void SetColumnWidths(ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[excelWorkSheet.Dimension.Address].AutoFitColumns();
            excelWorkSheet.Column(1).Width = 3;
        }

        private void AddWorkSheetHeader(ExcelWorksheet workSheetIn, ref int rowIndexIn, int colIndexIn, string headerIn, float fontSizeIn)
        {
            workSheetIn.Cells[rowIndexIn, colIndexIn].Value = headerIn;
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Color.SetColor(Color.FromArgb(89, 43, 95));
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Size = fontSizeIn;
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Name = "Azo Sans";
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Bold = true;
        }

        private void AddImageToWorksheet(ExcelWorksheet workSheetIn, int colIndexIn, int rowIndexIn, Image image)
        {
            if (image != null)
            {
                ExcelPicture picture = workSheetIn.Drawings.AddPicture("Bim Toolkit", image);
                picture.From.Column = colIndexIn;
                picture.From.Row = rowIndexIn;
                picture.SetSize(image.Width, image.Height);
            }
        }
    }
}
