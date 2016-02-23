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
using System.Globalization;

namespace Xbim.CobieLiteUK.Validation.Reporting
{
    internal class ExcelCellVisualValue2
    {
        public ExcelCellVisualValue2()
        { }

        private readonly ExcelStyle _orange;
        private readonly ExcelStyle _lightGreen;
        private readonly ExcelStyle _red;
        private readonly ExcelStyle _rose;
        private readonly ExcelStyle _neutral;

        public ExcelCellVisualValue2(ExcelWorksheet worksheet)
        {
        }

        /// <summary>
        /// Sets cell value and style based on IVisualValue
        /// </summary>
        /// <param name="excelCell">Cell to apply value and style to</param>
        /// <param name="visualValue"></param>
        internal void SetCell(ExcelRange excelCell, IVisualValue visualValue)
        {
            if (visualValue.AttentionStyle == VisualAttentionStyle.None)
            {
                //excelCell.CellStyle = _neutral;
            }
            switch (visualValue.AttentionStyle)
            {
                case VisualAttentionStyle.Amber:
                    excelCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 204));
                    break;
                case VisualAttentionStyle.Green:
                    excelCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));
                    break; ;
                case VisualAttentionStyle.Red:
                    excelCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 183, 185));
                    //excelCell.CellStyle = _red;
                    break;
            }

            var attribute = visualValue.VisualValue;
            if (attribute is StringAttributeValue)
            {
                excelCell.Value = ((StringAttributeValue)(attribute)).Value;
            }
            else if (attribute is IntegerAttributeValue)
            {
                var v = ((IntegerAttributeValue)(attribute)).Value;
                if (v.HasValue)
                {
                    excelCell.Value = (double)v.Value;
                }
            }
            else if (attribute is DecimalAttributeValue)
            {
                var v = ((DecimalAttributeValue)(attribute)).Value;
                if (v.HasValue)
                {
                    excelCell.Value = (double)v.Value;
                }
            }
            else if (attribute is BooleanAttributeValue)
            {
                var v = ((BooleanAttributeValue)(attribute)).Value;
                if (v.HasValue)
                {
                    excelCell.Value = v.Value;
                }
            }
            else if (attribute is DateTimeAttributeValue)
            {

                // var dataFormatStyle = excelCell.Sheet.Workbook.CreateDataFormat();
                //excelCell.CellStyle.DataFormat = 0x16; //  dataFormatStyle.GetFormat("yyyy/MM/dd HH:mm:ss");
                var v = ((DateTimeAttributeValue)(attribute)).Value;
                if (!v.HasValue)
                    return;
                // dataformats from: https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
                //excelCell.CellStyle.DataFormat = 0x16;
                //excelCell.SetCellValue(v.Value);
                excelCell.Value = v.Value.ToLongDateString();
                excelCell.Value = v.Value.ToString("G", DateTimeFormatInfo.InvariantInfo);
            }
        }
    }
}
