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
            //_orange = GetBaseStyle(worksheet);
            //_orange.Fill.BackgroundColor.SetColor(Color.Orange);

            //_lightGreen = GetBaseStyle(worksheet);
            //_lightGreen.Fill.BackgroundColor.SetColor(Color.LightGreen);

            //_red = GetBaseStyle(worksheet);
            //_red.Fill.BackgroundColor.SetColor(Color.Red);


            //_neutral = GetBaseStyle(worksheet);
        }

        //private ExcelStyle GetBaseStyle(ExcelWorksheet worksheet)
        //{
        //    var style = worksheet.Cells.Style;
        //    //style.Borders. = style.BorderLeft = style.BorderRight = style.BorderTop = BorderStyle.Thin;
        //    style.Border.BorderAround(ExcelBorderStyle.Thin);
        //    return style;
        //}

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
                    excelCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    //excelCell.CellStyle = _orange;
                    break;
                case VisualAttentionStyle.Green:
                    excelCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));
                    //excelCell.CellStyle = _lightGreen;
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
                //excelCell.SetCellType(CellType.String);
                //excelCell.SetCellValue(((StringAttributeValue)(attribute)).Value);

                excelCell.Value = ((StringAttributeValue)(attribute)).Value;

                // todo: can we set here ? cellStyle.Alignment = HorizontalAlignment.Fill;
            }
            else if (attribute is IntegerAttributeValue)
            {
                //excelCell.SetCellType(CellType.Numeric);
                var v = ((IntegerAttributeValue)(attribute)).Value;
                if (v.HasValue)
                {
                    // ReSharper disable once RedundantCast
                    //excelCell.SetCellValue((double)v.Value);
                    excelCell.Value = (double)v.Value;
                }
            }
            else if (attribute is DecimalAttributeValue)
            {
                //excelCell.SetCellType(CellType.Numeric);
                var v = ((DecimalAttributeValue)(attribute)).Value;
                if (v.HasValue)
                {
                    // ReSharper disable once RedundantCast
                    //excelCell.SetCellValue((double)v.Value);
                    excelCell.Value = (double)v.Value;
                }
            }
            else if (attribute is BooleanAttributeValue)
            {
                //excelCell.SetCellType(CellType.Boolean);
                var v = ((BooleanAttributeValue)(attribute)).Value;
                if (v.HasValue)
                {
                    //excelCell.SetCellValue(v.Value);
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
                excelCell.Value = v.Value;
            }
        }
    }
}
