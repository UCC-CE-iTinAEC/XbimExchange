using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
//using NPOI.HSSF.UserModel;
//using NPOI.SS.Formula.Functions;
//using NPOI.SS.UserModel;
//using NPOI.SS.Util;
//using NPOI.XSSF.UserModel;
//using NPOI.XSSF;
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
                    var result = Create(facility, spreadsheetStream, format);
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

        // ##### EPPlus START ########
        /// <summary>
        /// 
        /// </summary>
        /// <param name="facility"></param>
        /// <param name="suggestedFilename"></param>
        /// <returns></returns>
        public bool CreateSpreadsheet(Facility facility, string suggestedFilename)
        {
            SpreadSheetFormat format = SpreadSheetFormat.Xlsx;
            var ssFileName = Path.ChangeExtension(suggestedFilename, format.ToString());
            if (File.Exists(ssFileName))
            {
                File.Delete(ssFileName);
            }
            try
            {
                var outputDir = @"C:\github\XbimExchange\TestResults\";
                // Create the file using the FileInfo objectdata
                var fileInfo = new FileInfo(outputDir + suggestedFilename);

                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    //set the workbook properties and add a default sheet in it
                    SetWorkbookProperties(excelPackage);

                    CreateSummarySheet(excelPackage, facility);

                    Byte[] bin = excelPackage.GetAsByteArray();                    
                    File.WriteAllBytes(fileInfo.FullName, bin);
                }
            }
            catch (Exception e)
            {
                Logger.ErrorFormat("Failed to save {0}, {1}", ssFileName, e.Message);
                return false;
            }
            return true;
        }

        private static ExcelWorksheet CreateSummarySheet(ExcelPackage excelPackageIn, Facility facilityIn)
        {
            ExcelWorksheet summarySheet = AddWorkSheet(excelPackageIn, "Summary");
            // set column widths
            summarySheet.Column(2).Width = 60;
            summarySheet.Column(3).Width = 17;
            summarySheet.Column(4).Width = 14;
            summarySheet.Column(5).Width = 14;
            summarySheet.Column(6).Width = 9;

            // Set first rowIndex (after image) and colIndex (leave column1 blank)
            int rowIndex = 8;
            int colIndex = 2;

            string workSheetHeader = String.Format("{0} - Verification report - {1}", facilityIn.Project.Name, DateTime.Now.ToShortDateString());
            AddWorkSheetHeader(summarySheet, ref rowIndex, colIndex, workSheetHeader, 22);

            rowIndex += 2;
            string PreferredClassification = "Uniclass2015";

            if (facilityIn.AssetTypes != null && facilityIn.AssetTypes.Any())
            {
                var assetTypesReport = new GroupingObjectSummaryReport<CobieObject>(facilityIn.AssetTypes, "Asset types report");
                var report = assetTypesReport.GetReport(PreferredClassification);

                int numberRows = report.Rows.Count;
                int numberCols = numberRows > 0 ? report.Rows[0].ItemArray.Count()-1 : 0;

                //DataTable assetTypesReportTable = CreateDataTable(summarySheet, facilityIn, ref rowIndex, colIndex, numberRows, numberCols, true); // Generates DataTable
                
                //CreateData(summarySheet, ref rowIndex, colIndex, assetTypesReportTable);

                foreach (DataColumn dataCol in report.Columns) //Creating Headings
                {
                    if (dataCol == report.Columns[0]) continue;

                    var cell = summarySheet.Cells[rowIndex, colIndex];

                    //Setting the background color of header cells
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(243, 245, 244));

                    //Set borders.
                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    cell.Style.Border.Bottom.Color.SetColor(Color.FromArgb(62, 177, 200));

                    //Setting Value in cell
                    cell.Value = dataCol.Caption;

                    colIndex++;
                }

                // Reset column index
                colIndex = 2;

                //// Output report data to assetTypesReportTable
                //int dataRowIndex = 0;
                if (numberRows > 0)
                {
                    foreach (DataRow row in report.Rows)
                    {
                        rowIndex++;
                        for (int col = 0; col < numberCols; col++)
                        {
                            string name = row.ItemArray[col + 1].GetType().Name;
                            switch (name)
                            {
                                case "String":
                                    summarySheet.Cells[rowIndex, col + colIndex].Value = row.ItemArray[col + 1];
                                    break;
                                case "Int32":
                                    summarySheet.Cells[rowIndex, col + colIndex].Value = Convert.ToInt32(row.ItemArray[col + 1]);
                                    break;
                                case "VisualValue":
                                    IVisualValue value = row.ItemArray[col + 1] as IVisualValue;

                                    if (value.VisualValue is Xbim.COBieLiteUK.IntegerAttributeValue)
                                    {
                                        var intValue = value.VisualValue as Xbim.COBieLiteUK.IntegerAttributeValue;
                                        summarySheet.Cells[rowIndex, col + colIndex].Value = intValue.Value;
                                        summarySheet.Cells[rowIndex, col + colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        summarySheet.Cells[rowIndex, col + colIndex].Style.Fill.BackgroundColor.SetColor(GetColourForResult(value.AttentionStyle));
                                    }

                                    break;
                                default:
                                    summarySheet.Cells[rowIndex, col + colIndex].Value = row.ItemArray[col + 1];
                                    break;
                            }

                        }
                    }
                }

            }
            return summarySheet;
        }

        private static Color GetColourForResult(VisualAttentionStyle attentionStyle)
        {
            switch (attentionStyle)
            {
                case VisualAttentionStyle.Red:
                    return Color.Red;

                case VisualAttentionStyle.Amber:
                    return Color.Yellow;

                case VisualAttentionStyle.Green:
                    return Color.Green;

                default:
                    return Color.White;
            }
        }

        private static ExcelWorksheet AddWorkSheet(ExcelPackage excelPackageIn, string sheetNameIn)
        {
            excelPackageIn.Workbook.Worksheets.Add(sheetNameIn);
            ExcelWorksheet workSheet = excelPackageIn.Workbook.Worksheets[sheetNameIn];
            workSheet.Name = sheetNameIn;
            workSheet.Cells.Style.Font.Size = 11; 
            workSheet.Cells.Style.Font.Name = "Arial";
            workSheet.Column(1).Width = 3;
            workSheet.View.ShowGridLines = false;

            AddImage(workSheet, 1, 1, Xbim.CobieLiteUK.Validation.Properties.Resources.btk_logo_beta);

            return workSheet;
        }

        private static void AddWorkSheetHeader(ExcelWorksheet workSheetIn, ref int rowIndexIn, int colIndexIn, string headerIn, float fontSizeIn)
        {
            workSheetIn.Cells[rowIndexIn, colIndexIn].Value = headerIn;
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Color.SetColor(Color.FromArgb(89, 43, 95));
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Size = fontSizeIn;
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Name = "Azo Sans";
            workSheetIn.Cells[rowIndexIn, colIndexIn].Style.Font.Bold = true;
        }

        private static DataTable CreateDataTable(ExcelWorksheet workSheetIn, Facility facilityIn, ref int rowIndex, int colIndexIn, int numberRows, int numberCols, bool addHeader)
        {
            DataTable dataTable = new DataTable();
            for (int col = 0; col < numberCols; col++)
            {
                dataTable.Columns.Add(col.ToString());
            }

            for (int row = 0; row < numberRows; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                foreach (DataColumn dataCol in dataTable.Columns)
                {
                    dataRow[dataCol.ToString()] = row;
                }

                dataTable.Rows.Add(dataRow);
            }

            if (addHeader)
            {
                CreateDataTableHeader(workSheetIn, ref rowIndex, colIndexIn, dataTable);
            }
            return dataTable;
        }

        private static void CreateDataTableHeader(ExcelWorksheet workSheetIn, ref int rowIndex, int colIndex, DataTable dataTableIn)
        {
            //int colIndex = 1;
            foreach (DataColumn dataCol in dataTableIn.Columns) //Creating Headings
            {
                var cell = workSheetIn.Cells[rowIndex, colIndex];

                //Setting the background color of header cells
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(243, 245, 244));

                //Set borders.
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                cell.Style.Border.Bottom.Color.SetColor(Color.FromArgb(62, 177, 200));

                //Setting Value in cell
                cell.Value = dataCol.Caption;

                colIndex++;
            }
        }

       
        private static void SetWorkbookProperties(ExcelPackage excelPackageIn)
        {
            //Here setting some document properties
            excelPackageIn.Workbook.Properties.Author = "Xbim Cobie Lite UK";
            excelPackageIn.Workbook.Properties.Title = "Xbim Cobie Lite UK Validation";
        }

        private static void AddImage(ExcelWorksheet workSheetIn, int colIndexIn, int rowIndexIn, Image image)
        {
            //How to Add a Image using EP Plus
            //Bitmap image = new Bitmap(filePath);
            ExcelPicture picture = null;
            if (image != null)
            {
                picture = workSheetIn.Drawings.AddPicture("pic" + rowIndexIn.ToString() + colIndexIn.ToString(), image);
                picture.From.Column = colIndexIn;
                picture.From.Row = rowIndexIn;
                picture.From.ColumnOff = Pixel2MTU(2); //Two pixel space for better alignment
                picture.From.RowOff = Pixel2MTU(2);//Two pixel space for better alignment
                //picture.SetSize(100, 100);
            }
        }

        public static int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
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
        //private static bool CreateDetailSheet(ISheet detailSheet, TwoLevelRequirementPointer<AssetType, Asset> requirementPointer)
        //{
        //    try
        //    {
        //        var excelRow = detailSheet.GetRow(0) ?? detailSheet.CreateRow(0);
        //        var excelCell = excelRow.GetCell(0) ?? excelRow.CreateCell(0);
        //        SetHeader(excelCell);
        //        excelCell.SetCellValue("Asset Type and assets report");

        //        var rep = new TwoLevelDetailedGridReport<AssetType, Asset>(requirementPointer);
        //        rep.PrepareReport();

        //        var iRunningRow = 2;
        //        var iRunningColumn = 0;
        //        excelRow = detailSheet.GetRow(iRunningRow++) ?? detailSheet.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"Name:"); // writes cell and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(requirementPointer.Name); // writes cell and moves index forward

        //        iRunningColumn = 0;
        //        excelRow = detailSheet.GetRow(iRunningRow++) ?? detailSheet.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"External system:"); // writes cell and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(requirementPointer.ExternalSystem); // writes cell and moves index forward

        //        iRunningColumn = 0;
        //        excelRow = detailSheet.GetRow(iRunningRow++) ?? detailSheet.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"External id:"); // writes cell and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(requirementPointer.ExternalId); // writes cell and moves index forward

        //        iRunningRow++; // one empty row

        //        iRunningColumn = 0;
        //        excelRow = detailSheet.GetRow(iRunningRow++) ?? detailSheet.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //        (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(@"Matching categories:"); // writes cell and moves index forward

        //        foreach (var cat in rep.RequirementCategories)
        //        {
        //            iRunningColumn = 0;
        //            excelRow = detailSheet.GetRow(iRunningRow++) ?? detailSheet.CreateRow(iRunningRow - 1); // prepares a row and moves index forward
        //            (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(cat.Classification); // writes cell and moves index forward
        //            (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(cat.Code); // writes cell and moves index forward
        //            (excelRow.GetCell(iRunningColumn++) ?? excelRow.CreateCell(iRunningColumn - 1)).SetCellValue(cat.Description); // writes cell and moves index forward
        //        }

        //        iRunningRow++; // one empty row
        //        iRunningColumn = 0;

        //        var cellStyle = detailSheet.Workbook.CreateCellStyle();
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

        //        excelRow = detailSheet.GetRow(iRunningRow) ?? detailSheet.CreateRow(iRunningRow);
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

        //        var writer = new ExcelCellVisualValue(detailSheet.Workbook);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            excelRow = detailSheet.GetRow(iRunningRow) ?? detailSheet.CreateRow(iRunningRow);
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
        public string PreferredClassification = "Uniclass2015";

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



            OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml cellStyle = summaryPage.Workbook.Styles.CreateNamedStyle("cellStyle");
            cellStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            cellStyle.Style.Border.Bottom.Color.SetColor(Color.SkyBlue);
            cellStyle.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cellStyle.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cellStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;

            cellStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellStyle.Style.Fill.PatternColor.SetColor(Color.LightSlateGray);


            var failCellStyle = summaryPage.Workbook.Styles.CreateNamedStyle("failCellStyle");
            failCellStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            failCellStyle.Style.Fill.PatternColor.SetColor(Color.LightPink);

            //ExcelRow excelRow = summaryPage.Row(startingRow);

            ExcelRange excelCell = summaryPage.Cells[startingRow, iRunningColumn];

            excelCell.Value = table.TableName;
            startingRow++;

            //excelRow = summaryPage.Row(startingRow);
            foreach (DataColumn tCol in table.Columns)
            {
                if (tCol.AutoIncrement)
                    continue;
                var runCell = summaryPage.Cells[startingRow, iRunningColumn];
                iRunningColumn++;
                runCell.Value = tCol.Caption;
                runCell.StyleName = "cellStyle";                
            }

            startingRow++;

            //var writer = new ExcelCellVisualValue(summaryPage.Workbook);

            foreach (DataRow row in table.Rows)
            {
                //excelRow = summaryPage.Row(startingRow);
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
                                excelCell.Value = (string)row[tCol];
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
                summaryPage.Column(iRunningColumn).Width = irun;
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
