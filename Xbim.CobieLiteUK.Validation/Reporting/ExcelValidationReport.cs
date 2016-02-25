using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;
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
            /// <summary>
            /// Excel Binary File Format
            /// </summary>
            Xls,
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
            var ssFileName = Path.ChangeExtension(suggestedFilename, format == SpreadSheetFormat.Xlsx ? "xlsx" : "xls");
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
            else if (ext != "xls")
                format = SpreadSheetFormat.Xls;
            else
                return false;
            return Create(facility, filename, format);
        }

        //private List<KeyValuePair<string, string>> LinksToDetailsSheets = new List<KeyValuePair<string, string>>();
        private Dictionary<string, string> LinksToDetailsSheets = new Dictionary<string, string>();

        // ##### EPPlus START ########
        /// <summary>
        /// 
        /// </summary>
        /// <param name="facility"></param>
        /// <param name="suggestedFilename"></param>
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

                        //var detailPage = AddWorkSheet(excelPackage, validName);
                        //UNCOMMENT var detailPage = workBook.CreateSheet(validName);
                        //UNCOMMENT
                        if (!CreateDetailSheet(excelPackage, assetType, validName))
                            return false;
                        //UNCOMMENT
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
                        //UNCOMMENT var validName = WorkbookUtil.CreateSafeSheetName(string.Format(@"{0} {1}", iRunningWorkBook++, tName));
                        var validName = iRunningWorkBook++.ToString();

                        //UNCOMMENT var detailPage = workBook.CreateSheet(validName);
                        //UNCOMMENT
                        //if (!CreateDetailSheet(detailPage, zoneGroup))
                        //    return false;
                        //if (!CreateDetailSheet(excelPackage, zoneGroup, validName))
                        //    return false;
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
                WriteReportToPage(detailsWorkSheet, ref rowIndex, ref colIndex, table, "Asset Type and assets report");
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
                            //string sheetName = String.Format("{0}", CreateSafeSheetName((string)row[col]));
                            string sheetName;
                            LinksToDetailsSheets.TryGetValue((string)row[col], out sheetName);
                            //sheetName = "1 Ss_30_25_22_90 A";
                            SetHyperlinkToWorksheet(excelWorkSheet.Cells[rowIndex, colIndex], sheetName, "A1", "Go to report");
                        }
                    }

                    var writer = new ExcelCellVisualValue2(excelWorkSheet);
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

        //private Color GetColourForResult(VisualAttentionStyle attentionStyle)
        //{
        //    switch (attentionStyle)
        //    {
        //        case VisualAttentionStyle.Red:
        //            return Color.FromArgb(255, 183, 185);

        //        case VisualAttentionStyle.Amber:
        //            return Color.Orange;

        //        case VisualAttentionStyle.Green:
        //            return Color.FromArgb(198, 239, 206);
        //        default:
        //            return Color.White;
        //    }
        //}

        private void SetWorkbookProperties(ExcelPackage excelPackageIn)
        {
            //Here setting some document properties
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

            AddImage(workSheet, 1, 1, Xbim.CobieLiteUK.Validation.Properties.Resources.btk_logo_beta);

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

        private void AddImage(ExcelWorksheet workSheetIn, int colIndexIn, int rowIndexIn, Image image)
        {
            //How to Add a Image using EP Plus
            //Bitmap image = new Bitmap(filePath);
            if (image != null)
            {
                ExcelPicture picture = workSheetIn.Drawings.AddPicture("Bim Toolkit", image);
                picture.From.Column = colIndexIn;
                picture.From.Row = rowIndexIn;
                picture.SetSize(image.Width, image.Height);
                //picture.SetSize(100, 100);
            }
        }


        //=====================================================

        //=====================================================



        // ###### EPPlus END #########



        /// <summary>
        /// Creates the report.
        /// </summary>
        /// <param name="reportFacility">the result of a DPoW validation to be transformed into report form.</param>
        /// <param name="destinationStream">target stream for the spreadsheet</param>
        /// <param name="format">determines the excel format to use</param>
        /// <returns>true if successful, errors are cought and passed to Logger</returns>
        public bool Create(Facility reportFacility, Stream destinationStream, SpreadSheetFormat format)
        {
            //UNCOMMENT
            //var workBook = format == SpreadSheetFormat.Xlsx
            //    // ReSharper disable once RedundantCast
            //    ? (IWorkbook)new XSSFWorkbook()
            //    // ReSharper disable once RedundantCast
            //    : (IWorkbook)new HSSFWorkbook();
            //UNCOMMENT
            var workBook = new ExcelPackage();
            var facReport = new FacilityReport(reportFacility);

            ExcelWorksheet summaryPage = AddWorkSheet(workBook, "Summary");

            //UNCOMMENT var summaryPage = workBook.Worksheets.Add("Summary");
            if (!CreateSummarySheet(summaryPage, reportFacility))
                return false;

            // reports on Documents
            //
            //UNCOMMENT
            //if (reportFacility.Documents != null)
            //{
            //    var documentsPage = workBook.CreateSheet("Documents");
            //    if (!CreateDocumentDetailsSheet(documentsPage, reportFacility.Documents))
            //        return false;
            //}
            //UNCOMMENT

            var iRunningWorkBook = 1;
            // reports on AssetTypes details
            //
            // ReSharper disable once LoopCanBeConvertedToQuery // might restore once code is stable
            foreach (var assetType in facReport.AssetRequirementGroups)
            {
                // only report items with any assets submitted (a different report should probably be provided otherwise)

                if (assetType.GetSubmittedAssetsCount() < 1)
                    continue;
                var firstOrDefault = assetType.RequirementCategories.FirstOrDefault(cat => cat.Classification == @"Uniclass2015");
                if (firstOrDefault == null)
                    continue;
                var tName = firstOrDefault.Code;
                //UNCOMMENT var validName = WorkbookUtil.CreateSafeSheetName(string.Format(@"{0} {1}", iRunningWorkBook++, tName));

                //UNCOMMENT var detailPage = workBook.CreateSheet(validName);
                //UNCOMMENT
                //if (!CreateDetailSheet(detailPage, assetType))
                //    return false;
                //UNCOMMENT
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
                //UNCOMMENT var validName = WorkbookUtil.CreateSafeSheetName(string.Format(@"{0} {1}", iRunningWorkBook++, tName));

                //UNCOMMENT var detailPage = workBook.CreateSheet(validName);
                //UNCOMMENT
                //if (!CreateDetailSheet(detailPage, zoneGroup))
                //    return false;
                //UNCOMMENT
            }
            try
            {
                //UNCOMMENT workBook.Write(destinationStream);
                workBook.Save();
            }
            catch (Exception e)
            {
                Logger.ErrorFormat("Failed to stream excel report: {1}", e.Message);
                return false;
            }
            return true;
        }

        //UNCOMMENT
        //private bool CreateDocumentDetailsSheet(ISheet documentPage, List<Document> list)
        //{

        //    try 
        //    {
        //        var excelRow = documentPage.GetRow(0) ?? documentPage.CreateRow(0);
        //        var excelCell = excelRow.GetCell(0) ?? excelRow.CreateCell(0);
        //        //UNCOMMENT SetHeader(excelCell); 
        //        excelCell.SetCellValue("Documents Report");
        //        var iRunningRow = 2; 

        //        var drep = new DocumentsDetailedReport(list);
        //        iRunningRow = WriteReportToPage(documentPage, drep.AttributesGrid , iRunningRow, false);

        //        Debug.WriteLine(iRunningRow);
        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        //log the error
        //        Logger.Error("Failed to create Summary Sheet", e);
        //        return false;
        //    }
        //}
        //UNCOMMENT


        // UNCOMMENT
        //private static bool CreateDetailSheet(ISheet excelPackageIn, TwoLevelRequirementPointer<AssetType, Asset> requirementPointer)
        //{
        //    try
        //    {
        //        var excelRow = excelPackageIn.GetRow(0) ?? excelPackageIn.CreateRow(0);
        //        var excelCell = excelRow.GetCell(0) ?? excelRow.CreateCell(0);
        //        SetHeader(excelCell);
        //        excelCell.SetCellValue("Asset Type and assets report");

        //        var rep = new TwoLevelDetailedGridReport<AssetType, Asset>(requirementPointer);
        //        rep.PrepareReport();

        //        var iRunningRow = 2;
        //        var iRunningColumn = 0;
        //        excelRow = excelPackageIn.GetRow(iRunningRow++) ?? excelPackageIn.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"Name:"); // writes cell and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(requirementPointer.Name); // writes cell and moves index forward

        //        iRunningColumn = 0;
        //        excelRow = excelPackageIn.GetRow(iRunningRow++) ?? excelPackageIn.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"External system:"); // writes cell and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(requirementPointer.ExternalSystem); // writes cell and moves index forward

        //        iRunningColumn = 0;
        //        excelRow = excelPackageIn.GetRow(iRunningRow++) ?? excelPackageIn.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"External id:"); // writes cell and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(requirementPointer.ExternalId); // writes cell and moves index forward

        //        iRunningRow++; // one empty row

        //        iRunningColumn = 0;
        //        excelRow = excelPackageIn.GetRow(iRunningRow++) ?? excelPackageIn.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"Matching categories:"); // writes cell and moves index forward

        //        foreach (var cat in rep.RequirementCategories)
        //        {
        //            iRunningColumn = 0;
        //            excelRow = excelPackageIn.GetRow(iRunningRow++) ?? excelPackageIn.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //            (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(cat.Classification); // writes cell and moves index forward
        //            (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(cat.Code); // writes cell and moves index forward
        //            (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(cat.Description); // writes cell and moves index forward
        //        }

        //        iRunningRow++; // one empty row
        //        iRunningColumn = 0;

        //        var cellStyle = excelPackageIn.Workbook.CreateCellStyle();
        //        cellStyle.BorderBottom = BorderStyle.Thick;
        //        cellStyle.BorderLeft = BorderStyle.Thin;
        //        cellStyle.BorderRight = BorderStyle.Thin;
        //        cellStyle.BorderTop = BorderStyle.Thin;
        //        cellStyle.FillPattern = FillPattern.SolidForeground;
        //        cellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;


        //        byte[] rose = new byte[3] { 255, 183, 185 };
        //        XSSFCellStyle cellStyle1 = (XSSFCellStyle)excelCell.Sheet.Workbook.CreateCellStyle();
        //        cellStyle1.SetFillForegroundColor(new XSSFColor(rose)); cellStyle.BorderBottom = BorderStyle.Thick;
        //        cellStyle1.BorderLeft = BorderStyle.Thin;
        //        cellStyle1.BorderRight = BorderStyle.Thin;
        //        cellStyle1.BorderTop = BorderStyle.Thin;
        //        cellStyle1.FillPattern = FillPattern.SolidForeground;

        //        cellStyle.FillForegroundColor = cellStyle1.FillForegroundColor;

        //        var table = rep.AttributesGrid;

        //        excelRow = excelPackageIn.GetRow(iRunningRow) ?? excelPackageIn.CreateRow(iRunningRow);
        //        foreach (DataColumn tCol in table.Columns)
        //        {
        //            if (tCol.AutoIncrement)
        //                continue;
        //            excelCell = excelRow.GetCell(iRunningColumn) ?? excelRow.CreateCell(iRunningColumn);
        //            iRunningColumn++;
        //            excelCell.SetCellValue(tCol.Caption);
        //            excelCell.CellStyle = cellStyle;
        //        }
        //        iRunningRow++;

        //        var writer = new ExcelCellVisualValue(excelPackageIn.Workbook);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            excelRow = excelPackageIn.GetRow(iRunningRow) ?? excelPackageIn.CreateRow(iRunningRow);
        //            iRunningRow++;

        //            iRunningColumn = -1;
        //            foreach (DataColumn tCol in table.Columns)
        //            {
        //                if (tCol.AutoIncrement)
        //                    continue;
        //                iRunningColumn++;
        //                if (row[tCol] == DBNull.Value)
        //                    continue;
        //                excelCell = excelRow.GetCell(iRunningColumn) ?? excelRow.CreateCell(iRunningColumn);
        //                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
        //                if (row[tCol] is IVisualValue)
        //                {
        //                    writer.SetCell(excelCell, (IVisualValue)row[tCol]);
        //                }
        //                else
        //                {
        //                    switch (tCol.DataType.Name)
        //                    {
        //                        case "String":
        //                            excelCell.SetCellValue((string)row[tCol]);
        //                            break;
        //                        case "Int32":
        //                            excelCell.SetCellValue(Convert.ToInt32(row[tCol]));
        //                            break;
        //                        default:
        //                            excelCell.SetCellValue((string)row[tCol]);
        //                            break;
        //                    }
        //                }
        //            }
        //        }

        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        //log the error
        //        Logger.Error("Failed to create detail Sheet", e);
        //        return false;
        //    }
        //}
        // UNCOMMENT


        /// <summary>
        /// sets the Classification preferred for priority purposes.
        /// </summary>
        //public string PreferredClassification = "Uniclass2015";

        //UNCOMMENT
        //private bool CreateSummarySheet(ISheet summaryPage, Facility facility)
        //{
        //    try
        //    {
        //        var excelRow = summaryPage.GetRow(0) ?? summaryPage.CreateRow(0);  
        //        var excelCell = excelRow.GetCell(0) ?? excelRow.CreateCell(0);
        //        SetHeader(excelCell);
        //        excelCell.SetCellValue(String.Format("{0} - Verification report - {1}", facility.Project.Name, DateTime.Now.ToShortDateString()));
        //        var iRunningRow = 2;

        //        if (facility.AssetTypes != null && facility.AssetTypes.Any())
        //        {
        //            var assetTypesReport = new GroupingObjectSummaryReport<CobieObject>(facility.AssetTypes, "Asset types report");
        //            iRunningRow = WriteReportToPage(summaryPage, assetTypesReport.GetReport(PreferredClassification), iRunningRow);
        //        }

        //        if (facility.Zones != null && facility.Zones.Any())
        //        {
        //            var zonesReport = new GroupingObjectSummaryReport<CobieObject>(facility.Zones, "Zones report");
        //            iRunningRow = WriteReportToPage(summaryPage, zonesReport.GetReport(PreferredClassification),
        //                iRunningRow);
        //        }

        //        if (facility.Documents != null && facility.Documents.Any())
        //        {
        //            var docReport = new DocumentsReport(facility.Documents);
        //            iRunningRow = WriteReportToPage(summaryPage, docReport.GetReport("ResponsibleRole"), iRunningRow);
        //        }

        //        Debug.WriteLine(iRunningRow);
        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        //log the error
        //        Logger.Error("Failed to create Summary Sheet", e);
        //        return false;
        //    }
        //}
        //UNCOMMENT

        private bool CreateSummarySheet(ExcelWorksheet summaryPage, Facility facility)
        {
            try
            {
                var iRunningRow = 2;

                if (facility.AssetTypes != null && facility.AssetTypes.Any())
                {
                    var assetTypesReport = new GroupingObjectSummaryReport<CobieObject>(facility.AssetTypes, "Asset types report");
                    iRunningRow = WriteReportToPage(summaryPage, assetTypesReport.GetReport(PreferredClassification), iRunningRow);
                }

                if (facility.Zones != null && facility.Zones.Any())
                {
                    var zonesReport = new GroupingObjectSummaryReport<CobieObject>(facility.Zones, "Zones report");
                    iRunningRow = WriteReportToPage(summaryPage, zonesReport.GetReport(PreferredClassification),
                        iRunningRow);
                }

                if (facility.Documents != null && facility.Documents.Any())
                {
                    var docReport = new DocumentsReport(facility.Documents);
                    iRunningRow = WriteReportToPage(summaryPage, docReport.GetReport("ResponsibleRole"), iRunningRow);
                }

                Debug.WriteLine(iRunningRow);
                return true;
            }
            catch (Exception e)
            {
                //log the error
                Logger.Error("Failed to create Summary Sheet", e);
                return false;
            }
        }

        //UNCOMMENT
        //private int WriteReportToPage(ISheet summaryPage, DataTable table, int startingRow, Boolean autoSize = true)
        //{
        //    if (table == null)
        //        return startingRow;

        //    var iRunningColumn = 0;



        //    var cellStyle = summaryPage.Workbook.CreateCellStyle();
        //    cellStyle.BorderBottom = BorderStyle.Thick;
        //    cellStyle.BottomBorderColor = IndexedColors.SkyBlue.Index;
        //    cellStyle.BorderLeft = BorderStyle.Thin;
        //    cellStyle.BorderRight = BorderStyle.Thin;
        //    cellStyle.BorderTop = BorderStyle.Thin;

        //    cellStyle.FillPattern = FillPattern.SolidForeground;
        //    cellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;

        //    var index = IndexedColors.ValueOf("Grey25Percent");

        //    var failCellStyle = summaryPage.Workbook.CreateCellStyle();
        //    failCellStyle.FillPattern = FillPattern.SolidForeground;
        //    failCellStyle.FillForegroundColor = IndexedColors.Rose.Index;

        //    IRow excelRow = summaryPage.GetRow(startingRow) ?? summaryPage.CreateRow(startingRow);
        //    ICell excelCell = excelRow.GetCell(iRunningColumn) ?? excelRow.CreateCell(iRunningColumn);

        //    excelCell.SetCellValue(table.TableName);
        //    startingRow++;

        //    excelRow = summaryPage.GetRow(startingRow) ?? summaryPage.CreateRow(startingRow);
        //    foreach (DataColumn tCol in table.Columns)
        //    {
        //        if (tCol.AutoIncrement)
        //            continue;
        //        var runCell = excelRow.GetCell(iRunningColumn) ?? excelRow.CreateCell(iRunningColumn);
        //        iRunningColumn++;
        //        runCell.SetCellValue(tCol.Caption);
        //        runCell.CellStyle = cellStyle;
        //    }

        //    startingRow++;

        //    var writer = new ExcelCellVisualValue(summaryPage.Workbook);
        //    foreach (DataRow row in table.Rows)
        //    {
        //        excelRow = summaryPage.GetRow(startingRow) ?? summaryPage.CreateRow(startingRow);
        //        startingRow++;
        //        iRunningColumn = -1;
        //        foreach (DataColumn tCol in table.Columns)
        //        {
        //            if (tCol.AutoIncrement)
        //                continue;
        //            iRunningColumn++;
        //            if (row[tCol] == DBNull.Value)
        //                continue;
        //            excelCell = excelRow.GetCell(iRunningColumn) ?? excelRow.CreateCell(iRunningColumn);

        //            // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
        //            if (row[tCol] is IVisualValue)
        //            {
        //                writer.SetCell(excelCell, (IVisualValue) row[tCol]);
        //            }
        //            else
        //            {
        //                switch (tCol.DataType.Name)
        //                {
        //                    case "String":
        //                        excelCell.SetCellValue((string) row[tCol]);
        //                        break;
        //                    case "Int32":
        //                        excelCell.SetCellValue(Convert.ToInt32(row[tCol]));
        //                        break;
        //                    default:
        //                        excelCell.SetCellValue((string) row[tCol]);
        //                        break;
        //                }
        //            }
        //        }
        //    }

        //    if (!autoSize) 
        //        return startingRow + 1;
        //    // sets all used numberCols to autosize
        //    for (int irun = 0; irun < iRunningColumn; irun++)
        //    {
        //        summaryPage.AutoSizeColumn(irun);
        //    }
        //    return startingRow + 1;
        //}
        //UNCOMMENT

        private int WriteReportToPage(ExcelWorksheet summaryPage, DataTable table, int startingRow, Boolean autoSize = true)
        {
            if (table == null)
                return startingRow;

            var iRunningColumn = 0;



            //var cellStyle = summaryPage.Workbook.CreateCellStyle();
            //cellStyle.BorderBottom = BorderStyle.Thick;
            //cellStyle.BottomBorderColor = IndexedColors.SkyBlue.Index;
            //cellStyle.BorderLeft = BorderStyle.Thin;
            //cellStyle.BorderRight = BorderStyle.Thin;
            //cellStyle.BorderTop = BorderStyle.Thin;

            //cellStyle.FillPattern = FillPattern.SolidForeground;
            //cellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;

            //var index = IndexedColors.ValueOf("Grey25Percent");

            //var failCellStyle = summaryPage.Workbook.CreateCellStyle();
            //failCellStyle.FillPattern = FillPattern.SolidForeground;
            //failCellStyle.FillForegroundColor = IndexedColors.Rose.Index;

            ExcelRow excelRow = summaryPage.Row(startingRow);
            ExcelRange excelCell = summaryPage.Cells[startingRow, iRunningColumn];

            excelCell.Value = (table.TableName);
            startingRow++;

            excelRow = summaryPage.Row(startingRow);
            foreach (DataColumn tCol in table.Columns)
            {
                if (tCol.AutoIncrement)
                    continue;
                var runCell = summaryPage.Cells[startingRow, iRunningColumn];
                iRunningColumn++;
                runCell.Value = (tCol.Caption);
                //runCell.CellStyle = cellStyle;
            }

            startingRow++;

            //var writer = new ExcelCellVisualValue(summaryPage.Workbook);
            foreach (DataRow row in table.Rows)
            {
                excelRow = summaryPage.Row(startingRow);
                startingRow++;
                iRunningColumn = -1;
                foreach (DataColumn tCol in table.Columns)
                {
                    if (tCol.AutoIncrement)
                        continue;
                    iRunningColumn++;
                    if (row[tCol] == DBNull.Value)
                        continue;
                    excelCell = summaryPage.Cells[startingRow, iRunningColumn];

                    // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                    if (row[tCol] is IVisualValue)
                    {
                        //writer.SetCell(excelCell, (IVisualValue)row[tCol]);
                    }
                    else
                    {
                        switch (tCol.DataType.Name)
                        {
                            case "String":
                                excelCell.Value = ((string)row[tCol]);
                                break;
                            case "Int32":
                                excelCell.Value = (Convert.ToInt32(row[tCol]));
                                break;
                            default:
                                excelCell.Value = ((string)row[tCol]);
                                break;
                        }
                    }
                }
            }

            if (!autoSize)
                return startingRow + 1;
            // sets all used numberCols to autosize
            for (int irun = 0; irun < iRunningColumn; irun++)
            {
                //summaryPage.AutoSizeColumn(irun);
            }
            return startingRow + 1;
        }

        //UNCOMMENT
        //private static void SetHeader(ICell excelCell)
        //{
        //    var font = excelCell.Sheet.Workbook.CreateFont();
        //    font.FontHeightInPoints = 22;
        //    font.FontName = "Azo Sans";
        //    font.Boldweight = (short) FontBoldWeight.Bold;
        //    font.Color = IndexedColors.Orchid.Index;
        //    excelCell.CellStyle = excelCell.Sheet.Workbook.CreateCellStyle();
        //    excelCell.CellStyle.SetFont(font);
        //}
        //UNCOMMENT
    }
}
