using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Drawing;
using System.Globalization;
using NPOI;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using System.Text.RegularExpressions;

namespace To1C
{
    public class FileXLS
    {
        /// <summary>
        /// HSSFRow Copy Command
        /// 
        /// Description:  Inserts a existing row into a new row, will automatically push down
        ///               any existing rows.  Copy is done cell by cell and supports, and the
        ///               command tries to copy all properties available (style, merged cells, values, etc...)
        /// </summary>
        /// <param name="workbook">Workbook containing the worksheet that will be changed</param>
        /// <param name="worksheet">WorkSheet containing rows to be copied</param>
        /// <param name="sourceRowNum">Source Row Number</param>
        /// <param name="destinationRowNum">Destination Row Number</param>
        private static void CopyRow(IWorkbook workbook, ISheet worksheet, int sourceRowNum, List<List<DataEntry>> entries)
        {
            int shift = entries.Count;
            IRow sourceRow = worksheet.GetRow(sourceRowNum);
            worksheet.ShiftRows(sourceRowNum + 1, worksheet.LastRowNum, shift, true, true);
            for (int i = 1; i <= shift; i++)
            {
                IRow row = worksheet.GetRow(sourceRowNum + i);
                if (row != null)
                    worksheet.RemoveRow(row);
            }
            for (int r = 1; r <= shift; r++)
            {
                IRow newRow = worksheet.GetRow(sourceRowNum + r);
                if (newRow == null)
                    newRow = worksheet.CreateRow(sourceRowNum + r);
                // If the row exist in destination, push down all rows by 1 else create a new row
                //if (newRow != null)
                //{
                //    worksheet.ShiftRows(sourceRowNum + r, worksheet.LastRowNum, 1, true, false);
                //}
                //else
                //{
                //    newRow = worksheet.CreateRow(sourceRowNum + r);
                //}

                // Loop through source columns to add to new row
                for (int i = 0; i < sourceRow.LastCellNum; i++)
                {
                    // Grab a copy of the old/new cell
                    ICell oldCell = sourceRow.GetCell(i);
                    ICell newCell = newRow.CreateCell(i);

                    // If the old cell is null jump to next cell
                    if (oldCell == null)
                    {
                        newCell = null;
                        continue;
                    }

                    // Copy style from old cell and apply to new cell
                    //ICellStyle newCellStyle = workbook.CreateCellStyle();
                    //newCellStyle.CloneStyleFrom(oldCell.CellStyle); ;
                    //newCell.CellStyle = newCellStyle;
                    newCell.CellStyle = oldCell.CellStyle;

                    // If there is a cell comment, copy
                    if (oldCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                    // If there is a cell hyperlink, copy
                    if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                    // Set the cell data type
                    newCell.SetCellType(oldCell.CellType);

                    // Set the cell data value
                    switch (oldCell.CellType)
                    {
                        case CellType.Blank:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            newCell.SetCellValue(oldCell.BooleanCellValue);
                            break;
                        case CellType.Error:
                            newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                            break;
                        case CellType.Formula:
                            newCell.SetCellFormula(oldCell.CellFormula);
                            break;
                        case CellType.Numeric:
                            newCell.SetCellValue(oldCell.NumericCellValue);
                            break;
                        case CellType.String:
                            string cellText = oldCell.StringCellValue;
                            var variants = Regex.Matches(cellText, @"(\A|\s*)\{(\w+)\}(\s*|\z)")
                            .Cast<Match>()
                            .Select(m => m.Value.Trim());
                            foreach (string variant in variants)
                            {
                                DataEntry entry = entries[r - 1].Find(x => "{" + x.Имя + "}" == variant);
                                if (entry != null)
                                {
                                    cellText = cellText.Replace(variant, entry.Значение);
                                }
                            }
                            newCell.SetCellValue(cellText);
                            //newCell.SetCellValue(oldCell.RichStringCellValue);
                            break;
                        case CellType.Unknown:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                    }
                }

                //newRow.Height = -1;
                // If there are are any merged regions in the source row, copy to new row
                for (int i = 0; i < worksheet.NumMergedRegions; i++)
                {
                    CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);
                    if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                    {
                        CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                                                                                    (newRow.RowNum +
                                                                                     (cellRangeAddress.FirstRow -
                                                                                      cellRangeAddress.LastRow)),
                                                                                    cellRangeAddress.FirstColumn,
                                                                                    cellRangeAddress.LastColumn);
                        worksheet.AddMergedRegion(newCellRangeAddress);
                    }
                }
            }

        }
        public static void RemoveRow(ISheet sheet, int rowIndex)
        {
            if (rowIndex >= 0)
                sheet.ShiftRows(rowIndex + 1, sheet.LastRowNum + 1, -1);
        }
        private static void CopyRow(IWorkbook sourceWB, ISheet sourceWS, int sourceRowNum, IWorkbook destWB, ISheet destWS, int destinationRowNum)
        {
            // Get the source / new row
            IRow sourceRow = sourceWS.GetRow(sourceRowNum);
            IRow newRow = destWS.GetRow(destinationRowNum);
            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                destWS.ShiftRows(destinationRowNum, destWS.LastRowNum, 1);
            }
            else
            {
                newRow = destWS.CreateRow(destinationRowNum);
            }

            // Loop through source columns to add to new row
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);
                ICell newCell = newRow.CreateCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }

                // Copy style from old cell and apply to new cell
                ICellStyle newCellStyle = destWB.CreateCellStyle();
                newCellStyle.CloneStyleFrom(oldCell.CellStyle); ;
                newCell.CellStyle = newCellStyle;

                // If there is a cell comment, copy
                if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch (oldCell.CellType)
                {
                    case CellType.Blank:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula:
                        newCell.SetCellFormula(oldCell.CellFormula);
                        break;
                    case CellType.Numeric:
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String:
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }

            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < sourceWS.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = sourceWS.GetMergedRegion(i);
                if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                                                                                (newRow.RowNum +
                                                                                 (cellRangeAddress.FirstRow -
                                                                                  cellRangeAddress.LastRow)),
                                                                                cellRangeAddress.FirstColumn,
                                                                                cellRangeAddress.LastColumn);
                    destWS.AddMergedRegion(newCellRangeAddress);
                }
            }

        }

        private static string GetStringCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return cell.BooleanCellValue == true ? "1" : "0";
                case CellType.Error:
                    return "Error";
                case CellType.Formula:
                    return cell.CellFormula;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.RichStringCellValue.String;
                case CellType.Unknown:
                    return cell.StringCellValue;
                default:
                    return cell.StringCellValue;
            }
        }

        //private static void CopyRange(AreaReference range, ISheet destinationSheet)
        //{
        //    for (var rowNum = range.FirstCell.Row; rowNum <= range.LastCell.Row; rowNum++)
        //    {
        //        IRow sourceRow = sourceSheet.GetRow(rowNum);

        //        if (destinationSheet.GetRow(rowNum) == null)
        //            destinationSheet.CreateRow(rowNum);

        //        if (sourceRow != null)
        //        {
        //            IRow destinationRow = destinationSheet.GetRow(rowNum);

        //            for (var col = range.FirstColumn; col < sourceRow.LastCellNum && col <= range.LastColumn; col++)
        //            {
        //                destinationRow.CreateCell(col);
        //                CopyCell(sourceRow.GetCell(col), destinationRow.GetCell(col));
        //            }
        //        }
        //    }
        //}
        static public class PixelUtil
        {

            public static short EXCEL_COLUMN_WIDTH_FACTOR = 256;
            public static short EXCEL_ROW_HEIGHT_FACTOR = 20;
            public static int UNIT_OFFSET_LENGTH = 7;
            public static int[] UNIT_OFFSET_MAP = new int[] { 0, 36, 73, 109, 146, 182, 219 };

            public static int pixel2WidthUnits(int pxs)
            {
                int widthUnits = (short)(EXCEL_COLUMN_WIDTH_FACTOR * (pxs / UNIT_OFFSET_LENGTH));
                widthUnits += UNIT_OFFSET_MAP[(pxs % UNIT_OFFSET_LENGTH)];
                return widthUnits;
            }

            public static int widthUnits2Pixel(short widthUnits)
            {
                int pixels = (widthUnits / EXCEL_COLUMN_WIDTH_FACTOR) * UNIT_OFFSET_LENGTH;
                int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
                pixels += (int)Math.Floor((float)offsetWidthUnits / ((float)EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));
                return pixels;
            }

            public static int heightUnits2Pixel(short heightUnits)
            {
                int pixels = (heightUnits / EXCEL_ROW_HEIGHT_FACTOR);
                int offsetWidthUnits = heightUnits % EXCEL_ROW_HEIGHT_FACTOR;
                pixels += (int)Math.Floor((float)offsetWidthUnits / ((float)EXCEL_ROW_HEIGHT_FACTOR / UNIT_OFFSET_LENGTH));
                return pixels;
            }
        }
        private static void CopyCell(ICell source, ISheet destinationSheet, int offset)
        {
            if (source != null)
            {
                var h = PixelUtil.heightUnits2Pixel((short)source.Row.Height);
                var w = PixelUtil.widthUnits2Pixel((short)source.Row.Sheet.GetColumnWidth(source.ColumnIndex));

                if (destinationSheet.GetRow(source.RowIndex + offset) == null)
                {
                    destinationSheet.CreateRow(source.RowIndex + offset);
                }
                IRow destinationRow = destinationSheet.GetRow(source.RowIndex + offset);
                destinationRow.Height = (short)h;
                ICell destination = destinationRow.GetCell(source.ColumnIndex);
                if (destination == null)
                {
                    destinationRow.CreateCell(source.ColumnIndex);
                    var width = source.Row.Sheet.GetColumnWidth(source.ColumnIndex);
                    destinationSheet.SetColumnWidth(source.ColumnIndex, w);
                    destination = destinationRow.GetCell(source.ColumnIndex);
                }

                //you can comment these out if you don't want to copy the style ...
                // Copy style from old cell and apply to new cell
                ICellStyle newCellStyle = destinationSheet.Workbook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(source.CellStyle);
                destination.CellStyle = newCellStyle;

                destination.CellComment = source.CellComment;
                //destination.CellStyle = source.CellStyle;
                //destination.Hyperlink = source.Hyperlink;

                switch (source.CellType)
                {
                    case CellType.Formula:
                        destination.CellFormula = source.CellFormula; break;
                    case CellType.Numeric:
                        destination.SetCellValue(source.NumericCellValue); break;
                    case CellType.String:
                        destination.SetCellValue(source.StringCellValue); break;
                }
            }
        }

        private static void SetHeader(IWorkbook workbook, ISheet ws, int row, int column, string value, short FontSize = 10, bool Bold = true, bool Underline = true, bool Italic = true)
        {
            IFont fontHeader = workbook.CreateFont();
            fontHeader.FontName = HSSFFont.FONT_ARIAL;
            fontHeader.FontHeightInPoints = FontSize;
            fontHeader.Color = HSSFColor.Red.Index;
            fontHeader.IsItalic = Italic;
            if (Bold)
                fontHeader.Boldweight = (short)FontBoldWeight.Bold;
            if (Underline)
                fontHeader.Underline = FontUnderlineType.Single;
            ICellStyle styleHeader = workbook.CreateCellStyle();
            styleHeader.SetFont(fontHeader);
            IRow rowHeader = ws.CreateRow(row);
            ICell cell = rowHeader.CreateCell(column);
            cell.CellStyle = styleHeader;
            cell.SetCellValue(value);
        }

        private static void SetHead(IWorkbook workbook, ISheet ws, int rowStart, int rowEnd, int columnStart, int columnEnd, int width, string value,
                                    HorizontalAlignment HorAlignment, short FontSize, bool Bold, short Rotation = 0, short Color = HSSFColor.LemonChiffon.Index, ICellStyle parentStyle = null)
        {
            IFont font = workbook.CreateFont();
            font.FontName = HSSFFont.FONT_ARIAL;
            font.FontHeightInPoints = FontSize;
            if (Bold)
                font.Boldweight = (short)FontBoldWeight.Bold;

            ICellStyle styleGroup = workbook.CreateCellStyle();
            if (parentStyle != null)
                styleGroup.CloneStyleFrom(parentStyle);
            else
            {
                styleGroup.FillPattern = FillPattern.SolidForeground;
                styleGroup.FillForegroundColor = Color;
            }
            styleGroup.WrapText = true;
            styleGroup.VerticalAlignment = VerticalAlignment.Center;
            styleGroup.Alignment = HorAlignment;
            styleGroup.BorderRight = BorderStyle.Medium;
            styleGroup.BorderLeft = BorderStyle.Medium;
            styleGroup.BorderTop = BorderStyle.Medium;
            styleGroup.BorderBottom = BorderStyle.Medium;
            styleGroup.SetFont(font);
            styleGroup.Rotation = Rotation;
            for (int i = rowStart; i <= rowEnd; i++)
            {
                IRow rowTable = ws.GetRow(i);
                if (rowTable == null)
                    rowTable = ws.CreateRow(i);
                for (int y = columnStart; y <= columnEnd; y++)
                {
                    ICell cellGroup = rowTable.CreateCell(y);
                    cellGroup.CellStyle = styleGroup;
                    cellGroup.SetCellValue(value);
                }
            }
            if ((rowEnd > rowStart) | (columnEnd > columnStart))
            {
                CellRangeAddress cra = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
                ws.AddMergedRegion(cra);
            }
            if ((columnStart == columnEnd) && (width > 0))
                ws.SetColumnWidth(columnStart, width);
        }

        private static void SetValue(ISheet ws, ICellStyle cellStyle, int rowStart, int rowEnd, int columnStart, int columnEnd, string value)
        {
            for (int i = rowStart; i <= rowEnd; i++)
            {
                IRow rowTable = ws.GetRow(i);
                if (rowTable == null)
                    rowTable = ws.CreateRow(i);
                for (int y = columnStart; y <= columnEnd; y++)
                {
                    ICell cellGroup = rowTable.CreateCell(y);
                    if (cellStyle != null)
                        cellGroup.CellStyle = cellStyle;
                    cellGroup.SetCellValue(value);
                }
            }
            if ((rowEnd > rowStart) | (columnEnd > columnStart))
            {
                CellRangeAddress cra = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
                ws.AddMergedRegion(cra);
            }
        }
        private static void SetFormula(ISheet ws, ICellStyle cellStyle, int rowStart, int rowEnd, int columnStart, int columnEnd, string value)
        {
            for (int i = rowStart; i <= rowEnd; i++)
            {
                IRow rowTable = ws.GetRow(i);
                if (rowTable == null)
                    rowTable = ws.CreateRow(i);
                for (int y = columnStart; y <= columnEnd; y++)
                {
                    ICell cellGroup = rowTable.CreateCell(y);
                    if (cellStyle != null)
                        cellGroup.CellStyle = cellStyle;
                    cellGroup.SetCellType(CellType.Formula);
                    cellGroup.SetCellFormula(value);
                }
            }
            if ((rowEnd > rowStart) | (columnEnd > columnStart))
            {
                CellRangeAddress cra = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
                ws.AddMergedRegion(cra);
            }
        }
        private static void SetValue(ISheet ws, ICellStyle cellStyle, int rowStart, int rowEnd, int columnStart, int columnEnd, double value)
        {
            for (int i = rowStart; i <= rowEnd; i++)
            {
                IRow rowTable = ws.GetRow(i);
                if (rowTable == null)
                    rowTable = ws.CreateRow(i);
                for (int y = columnStart; y <= columnEnd; y++)
                {
                    ICell cellGroup = rowTable.CreateCell(y);
                    if (cellStyle != null)
                        cellGroup.CellStyle = cellStyle;
                    cellGroup.SetCellValue(value);
                }
            }
            if ((rowEnd > rowStart) | (columnEnd > columnStart))
            {
                CellRangeAddress cra = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
                ws.AddMergedRegion(cra);
            }
        }
        public static byte[] ReportZakazReglamentXLS(string extension, string SheetName, DateTime StartDate, DateTime EndDate, bool fullInterface,
                                bool S_CostZak, bool S_PrihRas, bool S_Filial, bool S_Upak, bool S_MinOtgr, bool S_Sertif, bool IsMaster, List<SkladProperties> SklProperties, List<ZakazReglament> rez, Results Result,
                                DateTime startDate, DateTime endDate, bool flBrend)
        {

            IWorkbook workbook;
            IWorkbook template;
            using (var tmpFile = new FileStream(HttpContext.Current.Server.MapPath("~/bin/") + "template." + extension, FileMode.Open, FileAccess.Read))
            {
                if (extension == "xls")
                {
                    workbook = new HSSFWorkbook();
                    template = new HSSFWorkbook(tmpFile);
                }
                else
                {
                    workbook = new XSSFWorkbook();
                    template = new XSSFWorkbook(tmpFile);
                }
            }
            ISheet ws = workbook.CreateSheet(SheetName);
            ISheet tempSheet = template.GetSheetAt(0);

            ICellStyle tempStyle = tempSheet.GetRow(1).GetCell(7).CellStyle;

            int columnSrMesProd = 0;
            int columnOstatok = 0;
            int columnZakaz = 0;

            IFont font8 = workbook.CreateFont();
            font8.FontName = HSSFFont.FONT_ARIAL;
            font8.FontHeightInPoints = 8;
            IFont font8Bold = workbook.CreateFont();
            font8Bold.FontName = HSSFFont.FONT_ARIAL;
            font8Bold.FontHeightInPoints = 8;
            font8Bold.Boldweight = (short)FontBoldWeight.Bold;

            ICellStyle styleRow = workbook.CreateCellStyle();
            styleRow.WrapText = true;
            styleRow.VerticalAlignment = VerticalAlignment.Center;
            styleRow.BorderRight = BorderStyle.Thin;
            styleRow.BorderLeft = BorderStyle.Thin;
            styleRow.BorderTop = BorderStyle.Thin;
            styleRow.BorderBottom = BorderStyle.Thin;
            styleRow.FillPattern = FillPattern.SolidForeground;
            styleRow.Alignment = HorizontalAlignment.Left;
            styleRow.FillForegroundColor = HSSFColor.White.Index;
            styleRow.SetFont(font8);

            ICellStyle styleRowNum = workbook.CreateCellStyle();
            styleRowNum.WrapText = true;
            styleRowNum.VerticalAlignment = VerticalAlignment.Center;
            styleRowNum.BorderRight = BorderStyle.Thin;
            styleRowNum.BorderLeft = BorderStyle.Thin;
            styleRowNum.BorderTop = BorderStyle.Thin;
            styleRowNum.BorderBottom = BorderStyle.Thin;
            styleRowNum.FillPattern = FillPattern.SolidForeground;
            styleRowNum.Alignment = HorizontalAlignment.Right;
            styleRowNum.FillForegroundColor = HSSFColor.White.Index;
            styleRowNum.SetFont(font8);
            styleRowNum.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.00");

            ICellStyle styleRowNumB = workbook.CreateCellStyle();
            styleRowNumB.WrapText = true;
            styleRowNumB.VerticalAlignment = VerticalAlignment.Center;
            styleRowNumB.BorderRight = BorderStyle.Thin;
            styleRowNumB.BorderLeft = BorderStyle.Thin;
            styleRowNumB.BorderTop = BorderStyle.Thin;
            styleRowNumB.BorderBottom = BorderStyle.Thin;
            styleRowNumB.FillPattern = FillPattern.SolidForeground;
            styleRowNumB.Alignment = HorizontalAlignment.Right;
            styleRowNumB.FillForegroundColor = HSSFColor.White.Index;
            styleRowNumB.SetFont(font8Bold);
            styleRowNumB.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.00");

            ICellStyle styleRowN = workbook.CreateCellStyle();
            styleRowN.WrapText = true;
            styleRowN.VerticalAlignment = VerticalAlignment.Center;
            styleRowN.BorderRight = BorderStyle.Thin;
            styleRowN.BorderLeft = BorderStyle.Thin;
            styleRowN.BorderTop = BorderStyle.Thin;
            styleRowN.BorderBottom = BorderStyle.Thin;
            styleRowN.FillPattern = FillPattern.SolidForeground;
            styleRowN.Alignment = HorizontalAlignment.Right;
            styleRowN.FillForegroundColor = HSSFColor.White.Index;
            styleRowN.SetFont(font8);

            ICellStyle styleRowN_BrightGreen = workbook.CreateCellStyle();
            styleRowN_BrightGreen.CloneStyleFrom(tempStyle);
            styleRowN_BrightGreen.WrapText = true;
            styleRowN_BrightGreen.VerticalAlignment = VerticalAlignment.Center;
            styleRowN_BrightGreen.BorderRight = BorderStyle.Thin;
            styleRowN_BrightGreen.BorderLeft = BorderStyle.Thin;
            styleRowN_BrightGreen.BorderTop = BorderStyle.Thin;
            styleRowN_BrightGreen.BorderBottom = BorderStyle.Thin;
            styleRowN_BrightGreen.Alignment = HorizontalAlignment.Right;
            styleRowN_BrightGreen.SetFont(font8);

            SetHeader(workbook, ws, 0, 0, String.Format("Заказ по регламенту на {0}", DateTime.Now.ToString("dd MMMM yyyy")));
            SetHeader(workbook, ws, 1, 0, String.Format("Остатки по складам на {0}, продажи за период с {1} по {2}", EndDate.AddDays(1).ToString("dd MMMM yyyy"), StartDate.ToString("dd MMMM yyyy"), EndDate.ToString("dd MMMM yyyy")), 8, false, false, false);
            int column = 0;
            SetHead(workbook, ws, 2, 3, 0, 2, 0, "ТМЦ", HorizontalAlignment.Center, 12, true);
            SetHead(workbook, ws, 4, 4, 0, 0, 2500, "Адрес", HorizontalAlignment.Center, 8, true);
            SetHead(workbook, ws, 4, 4, 1, 1, 2750, "Артикул", HorizontalAlignment.Center, 8, true);
            SetHead(workbook, ws, 4, 4, 2, 2, 11700, "Наименование", HorizontalAlignment.Center, 8, true);
            ws.GetRow(2).HeightInPoints = 15;
            ws.GetRow(3).HeightInPoints = 19.5F;
            ws.GetRow(2).HeightInPoints = 11.25F;
            column = 2;
            if (fullInterface)
            {
                SetHead(workbook, ws, 2, 4, column + 1, column + 1, 1000, "Сайт", HorizontalAlignment.Center, 8, false, 90);
                SetHead(workbook, ws, 2, 4, column + 2, column + 2, 550, "Яндекс", HorizontalAlignment.Center, 8, false, 90);
                SetHead(workbook, ws, 2, 4, column + 3, column + 3, 550, "Розница", HorizontalAlignment.Center, 8, false, 90);
                SetHead(workbook, ws, 2, 4, column + 4, column + 4, 550, "Аналоги", HorizontalAlignment.Center, 8, false, 90);
                SetHead(workbook, ws, 2, 4, column + 5, column + 5, 550, "Не закуп", HorizontalAlignment.Center, 8, false, 90);
                SetHead(workbook, ws, 2, 4, column + 6, column + 6, 2200, "Себестоимость", HorizontalAlignment.Center, 8, false, 90);
                column = column + 6;
            }
            if (S_CostZak)
            {
                SetHead(workbook, ws, 2, 3, column + 1, column + 5, 0, "Цена, руб.", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 4, 4, column + 1, column + 1, 2200, "СП", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 4, 4, column + 2, column + 2, 2200, "СП розн.", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 4, 4, column + 3, column + 3, 2200, "Зак", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 4, 4, column + 4, column + 4, 2200, "Опт", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 4, 4, column + 5, column + 5, 2200, "Розн", HorizontalAlignment.Center, 8, false);
                column = column + 5;
            }
            if (S_PrihRas)
            {
                if (fullInterface)
                {
                    SetHead(workbook, ws, 2, 3, column + 1, column + 3, 0, "Рентабельность (к себестоимости)", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 1, column + 1, 2200, "Опт", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 2, column + 2, 2200, "Розн", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 3, column + 3, 2200, "Факт", HorizontalAlignment.Center, 8, false);
                    column = column + 3;
                }
                SetHead(workbook, ws, 2, 4, column + 1, column + 1, 2200, "Ост. общ. на нач. периода", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 2, 4, column + 2, column + 2, 1200, "Прих. за пери- од", HorizontalAlignment.Center, 8, false);
                if (IsMaster)
                {
                    SetHead(workbook, ws, 2, 3, column + 3, column + 5, 0, "Расх. за пери- од", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 3, column + 3, 1100, "Ремонт", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 4, column + 4, 1100, "Продажи", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 5, column + 5, 1100, "ВСЕГО", HorizontalAlignment.Center, 8, false);
                    column = column + 5;
                }
                else
                {
                    SetHead(workbook, ws, 2, 4, column + 3, column + 3, 1200, "Расх. за пери- од", HorizontalAlignment.Center, 8, false);
                    column = column + 3;
                }
                SetHead(workbook, ws, 2, 2, column + 1, column + 2, 0, "Замена", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1000, "+", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 2, column + 2, 1000, "-", HorizontalAlignment.Center, 8, false);
                column = column + 2;
            }
            if (S_Filial)
            {
                SetHead(workbook, ws, 2, 2, column + 1, column + 11 + (IsMaster ? 2 : 0), 0, "Общие данные по всем филиалам", HorizontalAlignment.Center, 8, true);
                SetHead(workbook, ws, 3, 4, column + 1, column + 1, 2250, "Сред. мес. продажи", HorizontalAlignment.Center, 8, false);
                column = column + 1;

                if (IsMaster)
                {
                    SetHead(workbook, ws, 3, 3, column + 1, column + 3, 0, "Прод. за послед. месяц", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 1, column + 1, 1100, "Ремонт", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 2, column + 2, 1100, "Продажа", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 4, 4, column + 3, column + 3, 1100, "ВСЕГО", HorizontalAlignment.Center, 8, false);
                    column = column + 3;
                }
                else
                {
                    columnSrMesProd = column + 1;
                    //SetHead(workbook, ws, 3, 4, column + 1, column + 1, 2250, "Сред. мес. продажи", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 3, 4, column + 2, column + 2, 2250, "Прод. за послед. месяц", HorizontalAlignment.Center, 8, false);
                    column = column + 1;
                }
                SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1150, "Пред заказ", HorizontalAlignment.Center, 8, false, 0, 0, tempStyle);
                SetHead(workbook, ws, 3, 4, column + 2, column + 2, 1150, "Спрос", HorizontalAlignment.Center, 8, false, 0, 0, tempStyle);
                SetHead(workbook, ws, 3, 4, column + 3, column + 3, 1200, "В до- став- ке", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 4, column + 4, 1200, "В наборе", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 5, column + 5, 1200, "В ре- зер- вах", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 6, column + 6, 1200, "Ожид. приход", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 7, column + 7, 1200, "Ост-к ТМЦ", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 8, column + 8, 1200, "Ост-к свобод", HorizontalAlignment.Center, 8, false);
                SetHead(workbook, ws, 3, 4, column + 9, column + 9, 1200, "Ост-к аналогов", HorizontalAlignment.Center, 8, false);
                //SetHead(workbook, ws, 3, 4, column + 10, column + 10, 1200, "План", HorizontalAlignment.Center, 8, false);
                columnOstatok = column + 7;
                column = column + 9;
                if (fullInterface)
                {
                    columnZakaz = 0;
                    if (IsMaster)
                    {
                        SetHead(workbook, ws, 2, 4, column + 1, column + 1, 1200, "ЗА- КАЗ аварийный Г", HorizontalAlignment.Center, 8, true);
                        SetHead(workbook, ws, 2, 4, column + 2, column + 2, 1200, "ЗА- КАЗ аварийный П", HorizontalAlignment.Center, 8, true);
                        SetHead(workbook, ws, 2, 4, column + 3, column + 3, 1200, "ЗА- КАЗ", HorizontalAlignment.Center, 8, true);
                        SetHead(workbook, ws, 2, 4, column + 4, column + 4, 2400, "Сумма ЗАКАЗА, руб.", HorizontalAlignment.Center, 8, true);
                        column = column + 4;
                        columnZakaz = column + 4;
                    }
                }
                if (IsMaster)
                {
                    SetHead(workbook, ws, 2, 2, column + 1, column + 1, 1200, "ЦС", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1200, "План", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    column = column + 1;
                }
                if (!IsMaster)
                {
                    SetHead(workbook, ws, 2, 2, column + 1, column + 1, 1200, "ТЗ", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1200, "План / Факт", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    column = column + 1;
                }
                foreach (SkladProperties sp in SklProperties)
                {
                    SetHead(workbook, ws, 2, 2, column + 1, column + 3, 0, sp.Name, HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1200, "Ост. ТМЦ", HorizontalAlignment.Center, 8, false);
                    SetHead(workbook, ws, 3, 4, column + 2, column + 2, 1200, "Сред. мес. прод.", HorizontalAlignment.Center, 8, false);
                    if (sp.Name == "Экран")
                    {
                        SetHead(workbook, ws, 2, 2, column + 1, column + 2, 0, sp.Name, HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                        SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1200, "Ост. ТМЦ", HorizontalAlignment.Center, 8, false);
                        SetHead(workbook, ws, 3, 4, column + 2, column + 2, 1200, "Сред. мес. прод.", HorizontalAlignment.Center, 8, false);
                        //SetHead(workbook, ws, 3, 4, column + 3, column + 3, 1200, "ЦС План", HorizontalAlignment.Center, 8, false);
                        column = column + 2;
                    }
                    else
                    {
                        SetHead(workbook, ws, 2, 2, column + 1, column + 3, 0, sp.Name, HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                        SetHead(workbook, ws, 3, 4, column + 1, column + 1, 1200, "Ост. ТМЦ", HorizontalAlignment.Center, 8, false);
                        SetHead(workbook, ws, 3, 4, column + 2, column + 2, 1200, "Сред. мес. прод.", HorizontalAlignment.Center, 8, false);
                        SetHead(workbook, ws, 3, 4, column + 3, column + 3, 1200, "План", HorizontalAlignment.Center, 8, false);
                        column = column + 3;
                    }
                }
                if (fullInterface)
                {
                    SetHead(workbook, ws, 2, 3, column + 1, column + 1, 1200, "Акт", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 1, column + 1, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 2, column + 2, 1200, "Комиссия", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 2, column + 2, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 3, column + 3, 1200, "Некондиция", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 3, column + 3, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 4, column + 4, 1200, "Возврат", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 4, column + 4, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 5, column + 5, 1200, "Ремонт", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 5, column + 5, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 6, column + 6, 1200, "Розыск", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 6, column + 6, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 7, column + 7, 1200, "Рабочий", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 7, column + 7, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 2, 3, column + 8, column + 8, 1200, "Собст нужд", HorizontalAlignment.Center, 8, true, 0, HSSFColor.White.Index);
                    SetHead(workbook, ws, 4, 4, column + 8, column + 8, 1200, "Ост-к", HorizontalAlignment.Center, 8, false, 0, HSSFColor.White.Index);
                    column = column + 8;
                }
            }
            if (S_Upak)
            {
                SetHead(workbook, ws, 2, 4, column + 1, column + 1, 700, "Упак-ка пост-щика", HorizontalAlignment.Center, 8, false, 90);
                column = column + 1;
            }
            if (S_MinOtgr)
            {
                SetHead(workbook, ws, 2, 4, column + 1, column + 1, 700, "Мин. отгр.", HorizontalAlignment.Center, 8, false, 90);
                SetHead(workbook, ws, 2, 4, column + 2, column + 2, 700, "Мин. отг. опт", HorizontalAlignment.Center, 8, false, 90);
                column = column + 2;
            }
            if (S_Sertif)
            {
                SetHead(workbook, ws, 2, 4, column + 1, column + 1, 2500, "Сер- тифи- кат", HorizontalAlignment.Center, 8, false, 90);
                column = column + 1;
            }
            int row = 5;
            double SumZakaz = 0;
            double SumReglament = 0;
            foreach (ZakazReglament zr in rez)
            {
                column = 0;
                SetValue(ws, styleRow, row, row, 0, 0, zr.Address);
                SetValue(ws, styleRow, row, row, 1, 1, zr.Articul);
                SetValue(ws, styleRow, row, row, 2, 2, zr.Nomenklatura + (flBrend ? " /" + zr.Brend : ""));
                column = 2;
                if (fullInterface)
                {
                    SetValue(ws, styleRow, row, row, column + 1, column + 1, (zr.Site ? "да" : ""));
                    SetValue(ws, styleRow, row, row, column + 2, column + 2, zr.YandexMarket);
                    SetValue(ws, styleRow, row, row, column + 3, column + 3, zr.TolkoRozn);
                    SetValue(ws, styleRow, row, row, column + 4, column + 4, zr.ContainAnalog);
                    SetValue(ws, styleRow, row, row, column + 5, column + 5, zr.NeZakup);
                    SetValue(ws, styleRowNum, row, row, column + 6, column + 6, zr.Sebestoim);
                    column = column + 6;
                }
                if (S_CostZak)
                {
                    SetValue(ws, styleRowNum, row, row, column + 1, column + 1, zr.CostOptSP);
                    SetValue(ws, styleRowNum, row, row, column + 2, column + 2, zr.CostRoznSP);
                    SetValue(ws, styleRowNum, row, row, column + 3, column + 3, zr.CostZak);
                    SetValue(ws, styleRowNum, row, row, column + 4, column + 4, zr.CostOpt);
                    SetValue(ws, styleRowNum, row, row, column + 5, column + 5, zr.CostRozn);
                    column = column + 5;
                }
                if (S_PrihRas)
                {
                    if (fullInterface)
                    {
                        SetValue(ws, styleRowNum, row, row, column + 1, column + 1, ((zr.Sebestoim == 0) ? 0 : Math.Round(zr.CostOpt / zr.Sebestoim, 2)));
                        SetValue(ws, styleRowNum, row, row, column + 2, column + 2, ((zr.Sebestoim == 0) ? 0 : Math.Round(zr.CostRozn / zr.Sebestoim, 2)));
                        SetValue(ws, styleRowNum, row, row, column + 3, column + 3, ((zr.Spisstoim == 0) ? 0 : Math.Round(zr.ProdStoim / zr.Spisstoim, 2)));
                        column = column + 3;
                    }
                    SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.OstatokStartAll);
                    SetValue(ws, styleRowN, row, row, column + 2, column + 2, zr.PrihodAll);
                    if (IsMaster)
                    {
                        SetValue(ws, styleRowN, row, row, column + 3, column + 3, zr.Remont);
                        SetValue(ws, styleRowN, row, row, column + 4, column + 4, zr.RashodAll);
                        SetValue(ws, styleRowN, row, row, column + 5, column + 5, zr.RashodAll + zr.Remont);
                        column = column + 5;
                    }
                    else
                    {
                        SetValue(ws, styleRowN, row, row, column + 3, column + 3, zr.RashodAll);
                        column = column + 3;
                    }
                    SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.ZamenaPlus);
                    SetValue(ws, styleRowN, row, row, column + 2, column + 2, zr.ZamenaMinus);
                    column = column + 2;
                }
                if (S_Filial)
                {
                    SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.ProdanoAll);
                    column = column + 1;
                    if (IsMaster)
                    {
                        SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.RemontM);
                        SetValue(ws, styleRowN, row, row, column + 2, column + 2, zr.ProdanoAllM - zr.RemontM);
                        SetValue(ws, styleRowN, row, row, column + 3, column + 3, zr.ProdanoAllM);
                        column = column + 3;
                    }
                    else
                    {
                        //SetValue(ws, styleRowN, row, row, column + 1, column + 1, "", zr.ProdanoAll);
                        SetValue(ws, styleRowN, row, row, column + 2, column + 2, zr.ProdanoAllM);
                        column = column + 1;
                    }
                    SetValue(ws, styleRowN_BrightGreen, row, row, column + 1, column + 1, zr.PredZakaz);
                    SetValue(ws, styleRowN_BrightGreen, row, row, column + 2, column + 2, zr.Spros);
                    SetValue(ws, styleRowN, row, row, column + 3, column + 3, (zr.Dostavka > 0 ? zr.Dostavka.ToString() : "") + " / " + (zr.RezRaspred > 0 ? zr.RezRaspred.ToString() : ""));
                    SetValue(ws, styleRowN, row, row, column + 4, column + 4, zr.VNabore);
                    SetValue(ws, styleRowN, row, row, column + 5, column + 5, zr.VZayavke);
                    SetValue(ws, styleRowN, row, row, column + 6, column + 6, zr.OPrihod);
                    SetValue(ws, styleRowN, row, row, column + 7, column + 7, zr.OstatokEndAll);
                    SetValue(ws, styleRowN, row, row, column + 8, column + 8, zr.OstatokEndSvobodAll);
                    SetValue(ws, styleRowN, row, row, column + 9, column + 9, zr.OstatokAnalog);
                    //SetValue(ws, styleRowN, row, row, column + 10, column + 10, "", zr.NalFilAll);
                    SumReglament = SumReglament + (zr.CostZak > 0 ? (zr.CostZak * zr.NalFilAll) : (zr.Sebestoim * zr.NalFilAll));
                    column = column + 9;
                    if (fullInterface)
                    {
                        SumZakaz = SumZakaz + (zr.zakaz * zr.CostZak);
                        if (IsMaster)
                        {
                            SetValue(ws, styleRowN, row, row, column + 1, column + 1, Math.Max((zr.PodZakaz - zr.PodZakazPlatno - zr.OstatokEndSvobodAll), 0));
                            SetValue(ws, styleRowN, row, row, column + 2, column + 2, Math.Max((zr.PodZakazPlatno - zr.OstatokEndSvobodAll), 0));
                            if ((zr.zakaz == 0) && (zr.RashodAll + zr.Remont > 0) && (zr.ReglamentCenter == 0))
                                SetValue(ws, styleRowN, row, row, column + 3, column + 3, "?");
                            else
                                SetValue(ws, styleRowN, row, row, column + 3, column + 3, zr.zakaz);
                            SetValue(ws, styleRowNum, row, row, column + 4, column + 4, zr.zakaz * zr.CostZak);
                            column = column + 4;
                        }
                    }
                    if (IsMaster)
                    {
                        SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.ReglamentCenter);
                        column = column + 1;
                    }
                    if (!IsMaster)
                    {
                        SetValue(ws, styleRowN, row, row, column + 1, column + 1, (zr.ReglamentTZ == 0 ? "-" : zr.ReglamentTZ.ToString()) + " / " + (zr.OstatokTZ > 0 ? zr.OstatokTZ.ToString() : ""));
                        column = column + 1;
                    }
                    foreach (SkladProperties sp in zr.sp)
                    {
                        SetValue(ws, styleRowN, row, row, column + 1, column + 1, sp.Ostatok);
                        SetValue(ws, styleRowN, row, row, column + 2, column + 2, sp.prodagi);
                        column = column + 2;
                        if (sp.Name != "Экран")
                        {
                            SetValue(ws, styleRowN, row, row, column + 1, column + 1, (sp.nalFil > 0 ? sp.nalFil.ToString() : "--"));
                            column = column + 1;
                        }
                    }
                    if (fullInterface)
                    {
                        SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.OstatokAkt);
                        SetValue(ws, styleRowN, row, row, column + 2, column + 2, zr.Komiss);
                        SetValue(ws, styleRowN, row, row, column + 3, column + 3, zr.OstatokBrakMX);
                        SetValue(ws, styleRowN, row, row, column + 4, column + 4, zr.OstatokProkat);
                        SetValue(ws, styleRowN, row, row, column + 5, column + 5, zr.OstatokPP);
                        SetValue(ws, styleRowN, row, row, column + 6, column + 6, zr.OstatokRoz);
                        SetValue(ws, styleRowN, row, row, column + 7, column + 7, zr.OstatokOtstoy);
                        SetValue(ws, styleRowN, row, row, column + 8, column + 8, zr.OstatokSobNugdy);
                        column = column + 8;
                    }
                }
                if (S_Upak)
                {
                    SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.KolUpak);
                    column = column + 1;
                }
                if (S_MinOtgr)
                {
                    SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.KolPeremes);
                    SetValue(ws, styleRowN, row, row, column + 2, column + 2, zr.KolPeremesOpt);
                    column = column + 2;
                }
                if (S_Sertif)
                {
                    SetValue(ws, styleRowN, row, row, column + 1, column + 1, zr.Certif);
                    column = column + 1;
                }

                row++;
            }
            Result.ProdanoAllSum = rez.Sum(o => o.ProdanoAllSum);
            Result.ProdanoAllMSum = rez.Sum(o => o.ProdanoAllMSum);
            Result.OstatokEndAllSum = rez.Sum(o => o.OstatokEndAllSum);
            Result.OstatokEndSvobodAllSum = rez.Sum(o => o.OstatokEndSvobodAllSum);
            Result.SumReglament = Math.Round(SumReglament, 2);
            Result.SumZakaz = Math.Round(SumZakaz, 2);
            SetValue(ws, styleRowNumB, row, row, 0, 2, "ИТОГИ:");
            if (!IsMaster)
            {
                SetValue(ws, styleRowNumB, row, row, columnSrMesProd, columnSrMesProd, Result.ProdanoAllSum);
                SetValue(ws, styleRowNumB, row, row, columnSrMesProd + 1, columnSrMesProd + 1, Result.ProdanoAllMSum);
            }
            SetValue(ws, styleRowNumB, row, row, columnOstatok, columnOstatok, Result.OstatokEndAllSum);
            SetValue(ws, styleRowNumB, row, row, columnOstatok + 1, columnOstatok + 1, Result.OstatokEndSvobodAllSum);
            if (IsMaster)
            {
                SetValue(ws, styleRowNumB, row, row, columnOstatok + 2, columnOstatok + 3, Result.SumReglament);
                SetValue(ws, styleRowNumB, row, row, columnZakaz, columnZakaz, Result.SumZakaz);
            }

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }
        public static byte[] CreateResultYandexFBY(string extension, byte[] data, List<RowYandexFBY> report)
        {
            IWorkbook workbook;
            using (var stream = new MemoryStream(data))
            {
                if (extension == "xls")
                    workbook = new HSSFWorkbook(stream);
                else
                    workbook = new XSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheet("Поставка");
            List<int> deleteIndexes = new List<int>();
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    string sku = GetCellStringValue(row, 0);
                    if (!string.IsNullOrEmpty(sku))
                    {
                        var reportData = report.Where(x => x.Sku == sku).FirstOrDefault();
                        if (reportData != null && reportData.НовыйЗаказ1 > 0)
                        {
                            SetValue(sheet, null, i, i, 3, 3, reportData.НовыйЗаказ1);
                            SetValue(sheet, null, i, i, 4, 4, reportData.Цена);
                            SetValue(sheet, null, i, i, 5, 5, 7);
                            SetValue(sheet, null, i, i, 6, 6, 1);
                        }
                        else
                            deleteIndexes.Add(i);
                    }
                    else
                        deleteIndexes.Add(i);
                }
            }
            foreach (int i in deleteIndexes.OrderByDescending(x => x))
                RemoveRow(sheet, i);

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }
        public static List<RowYandexFBY> LoadReportExcel(string extension, byte[] data, string sheetName)
        {
            List<RowYandexFBY> result = new List<RowYandexFBY>();
            IWorkbook workbook;
            using (var stream = new MemoryStream(data))
            {
                if (extension == "xls")
                    workbook = new HSSFWorkbook(stream);
                else
                    workbook = new XSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
                sheet = workbook.GetSheetAt(0);
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    result.Add(new RowYandexFBY
                    {
                        Id = GetCellStringValue(row, 1),
                        Sku = GetCellStringValue(row, 2),
                        Артикул = GetCellStringValue(row, 3),
                        Наименование = GetCellStringValue(row, 4),
                        СвободныйОстаток = GetCellNumericValue(row, 5),
                        Поставлено = GetCellNumericValue(row, 6),
                        Впути = GetCellNumericValue(row, 7),
                        ОстаткиНаСкладеЯндекс1 = GetCellNumericValue(row, 8),
                        НовыйЗаказ1 = GetCellNumericValue(row, 9)
                    });
                }
            }
            return result;
        }
        public static List<string> GetSKUfromExcel(string extension, byte[] data, int columnIndexZeroBased, int startLine)
        {
            List<string> result = new List<string>();
            IWorkbook workbook;
            using (var stream = new MemoryStream(data))
            {
                if (extension == "xls")
                    workbook = new HSSFWorkbook(stream);
                else
                    workbook = new XSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = sheet.FirstRowNum + startLine; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    var sku = GetCellStringValue(row, columnIndexZeroBased);
                    if (!string.IsNullOrEmpty(sku))
                        result.Add(GetCellStringValue(row, columnIndexZeroBased));
                }
            }
            return result;
        }
        public static Dictionary<string, double> GetSkuValuesfromExcel(string extension, byte[] data, int columnSKUZeroBased, int columnAmountZeroBased, int startLine, int sheetNumber = 0)
        {
            Dictionary<string, double> result = new Dictionary<string, double>();
            IWorkbook workbook;
            using (var stream = new MemoryStream(data))
            {
                if (extension == "xls")
                    workbook = new HSSFWorkbook(stream);
                else
                    workbook = new XSSFWorkbook(stream);
            }
            ISheet sheet = workbook.GetSheetAt(sheetNumber);
            for (int i = sheet.FirstRowNum + startLine; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null) 
                {
                    string key = GetCellStringValue(row, columnSKUZeroBased);
                    if (!string.IsNullOrEmpty(key))
                    {
                        double value = GetCellNumericValue(row, columnAmountZeroBased);
                        if (result.Any(x => x.Key == key))
                            result[key] += value;
                        else
                            result.Add(key, value);
                    }
                }
            }
            return result;
        }
        private static string GetCellStringValue(IRow row, int index)
        {
            ICell cell = row.GetCell(index, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell != null && cell.CellType == CellType.String)
                return cell.StringCellValue;
            return "";
        }
        private static double GetCellNumericValue(IRow row, int index)
        {
            ICell cell = row.GetCell(index, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell != null && cell.CellType == CellType.Numeric)
                try
                {
                    return cell.NumericCellValue;
                }
                catch
                {
                    return 0;
                }
            return 0;
        }
        private static void CreateSheetYandexFBY(IWorkbook workbook, string sheetName, Dictionary<string, ICellStyle> styles,
            List<RowYandexFBY> data, DateTime startDate, DateTime endDate, List<string> skus, List<string> unusedSkus)
        {
            ISheet sheet = workbook.CreateSheet(sheetName);

            sheet.SetColumnWidth(0, 500);
            sheet.SetColumnWidth(1, 2200);
            sheet.SetColumnWidth(2, 2800);
            sheet.SetColumnWidth(3, 4300);
            sheet.SetColumnWidth(4, 8000);
            sheet.SetColumnWidth(5, 2900);//ожидаемый приход
            sheet.SetColumnWidth(6, 2900); 
            sheet.SetColumnWidth(7, 2900);
            sheet.SetColumnWidth(8, 2900);
            sheet.SetColumnWidth(9, 2900);//остаток1 == Самара
            sheet.SetColumnWidth(10, 2900);//заказ Самара
            sheet.SetColumnWidth(11, 2900);//квант
            sheet.SetColumnWidth(12, 2900);//цена закупочная
            sheet.SetColumnWidth(13, 2900);//цена
            sheet.SetColumnWidth(14, 2900);//цена сп
            sheet.SetColumnWidth(15, 2200);//коэф
            sheet.SetColumnWidth(16, 2900);//цена яндекс
            sheet.SetColumnWidth(17, 2200);//коэф яндекса
            sheet.SetColumnWidth(18, 2900);//цена первая
            sheet.SetColumnWidth(19, 2200);//коэф
            sheet.SetColumnWidth(20, 2900);//новая цена
            sheet.SetColumnWidth(21, 2900);//новая сп
            sheet.SetColumnWidth(22, 2900);//новая цена на яндексе
            sheet.SetColumnWidth(23, 2200);//коэф 

            string columnЗаказСамара = CellReference.ConvertNumToColString(10);
            string columnЦенаЗакупочная = CellReference.ConvertNumToColString(12);
            string columnЦенаЯндекса = CellReference.ConvertNumToColString(16);
            string columnНоваяЦенаНаЯндексе = CellReference.ConvertNumToColString(22);

            int rowNum = 0;
            SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 7, sheetName + " " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));
            rowNum++;
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 1, 1, "ID");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 2, 2, "SKU");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 3, 3, "Артикул");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 4, 4, "Наименование");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 5, 5, "Ожидаемый приход");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 6, 6, "Свободный остаток");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 7, 7, "Поставлено с " + startDate.ToString("dd.MM.yyyy") + " по " + endDate.ToString("dd.MM.yyyy"));
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 8, 8, "В пути c " + endDate.AddDays(1).ToString("dd.MM.yyyy"));
            SetValue(sheet, styles["Header"], rowNum, rowNum, 9, 10, "Самара");
            SetValue(sheet, styles["Header"], rowNum + 1, rowNum + 1, 9, 9, "Остаток на складе");
            SetValue(sheet, styles["Header"], rowNum + 1, rowNum + 1, 10, 10, "Новый заказ");
            SetValue(sheet, styles["Header"], rowNum + 1, rowNum + 1, 11, 11, "Квант");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 12, 12, "Цена закупочная");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 13, 13, "Цена");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 14, 14, "Цена СП");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 15, 15, "Коэф цена/з");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 16, 16, "Цена на Яндексе");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 17, 17, "Коэф я/з");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 18, 18, "Первая цена");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 19, 19, "Коэф п/з");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 20, 20, "Новая цена");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 21, 21, "Новое СП");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 22, 22, "Новая цена на яндексе");
            SetValue(sheet, styles["Header"], rowNum, rowNum + 1, 23, 23, "Коэф ня/з");

            rowNum++;
            rowNum++;

            int startRow = rowNum;
            int endRow = startRow;
            foreach (var entry in data)
            {
                SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 1, entry.Id);
                SetValue(sheet, styles["Value"], rowNum, rowNum, 2, 2, entry.Sku);
                SetValue(sheet, styles["Value"], rowNum, rowNum, 3, 3, entry.Артикул);
                SetValue(sheet, styles["Value"], rowNum, rowNum, 4, 4, entry.Наименование);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 5, 5, entry.ОжидаемыйПриход);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 6, 6, entry.СвободныйОстаток);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 7, 7, entry.Поставлено);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 8, 8, entry.Впути);
                SetValue(sheet, entry.Allowed1 ? styles["ValueNumMark"] : styles["ValueNum"], rowNum, rowNum, 9, 9, entry.ОстаткиНаСкладеЯндекс1);
                SetValue(sheet, entry.Allowed1 ? styles["ValueNumMark"] : styles["ValueNum"], rowNum, rowNum, 10, 10, entry.НовыйЗаказ1);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 11, 11, entry.Квант);
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 12, 12, entry.ЦенаЗакуп);
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 13, 13, entry.Цена);
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 14, 14, entry.ЦенаСП);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 15, 15, Math.Min(entry.Цена, entry.ЦенаСП > 0 ? entry.ЦенаСП : entry.Цена) / entry.ЦенаЗакуп);
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 16, 16, entry.ЦенаНаЯндексе);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 17, 17, (entry.ЦенаЗакуп == 0 ? 0 : entry.ЦенаНаЯндексе / entry.ЦенаЗакуп));
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 18, 18, entry.ЦенаПервая);
                SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 19, 19, (entry.ЦенаЗакуп == 0 ? 0 : entry.ЦенаПервая / entry.ЦенаЗакуп));
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 20, 20, 0);
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 21, 21, 0);
                SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 22, 22, 0);
                SetFormula(sheet, styles["ValueNum"], rowNum, rowNum, 23, 23, "IF(" + columnЦенаЗакупочная + (rowNum+1).ToString() + "," + columnНоваяЦенаНаЯндексе + (rowNum + 1).ToString() + "/" + columnЦенаЗакупочная + (rowNum + 1).ToString() + ",0)");
                endRow = rowNum;
                rowNum++;
            }
            var startInd = startRow.ToString();
            var endInd = endRow.ToString();
            rowNum++;
            SetValue(sheet, styles["Header"], rowNum, rowNum, 1, 4, "Итого");
            SetValue(sheet, styles["Header"], rowNum, rowNum, 5, 5, "шт");
            SetValue(sheet, styles["Header"], rowNum, rowNum, 6, 6, "руб. (закуп)");
            SetValue(sheet, styles["Header"], rowNum, rowNum, 7, 7, "руб. (продаж)");
            SetValue(sheet, styles["Header"], rowNum, rowNum, 8, 8, "Заказ, шт");
            SetValue(sheet, styles["Header"], rowNum, rowNum, 9, 9, "Заказ, руб");
            rowNum++;
            SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 4, "Поставлено");
            SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 5, 5, data.Sum(x => x.Поставлено));
            SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 6, 6, data.Sum(x => x.Поставлено * x.ЦенаЗакуп));
            SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 7, 7, data.Sum(x => x.Поставлено * x.ЦенаНаЯндексе));
            SetValue(sheet, styles["Value"], rowNum, rowNum, 8, 8, "");
            SetValue(sheet, styles["Value"], rowNum, rowNum, 9, 9, "");
            rowNum++;
            SetValue(sheet, styles["Value"], rowNum , rowNum, 1, 4, "В пути");
            SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 5, 5, data.Sum(x => x.Впути));
            SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 6, 6, data.Sum(x => x.Впути * x.ЦенаЗакуп));
            SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 7, 7, data.Sum(x => x.Впути * x.ЦенаНаЯндексе));
            SetValue(sheet, styles["Value"], rowNum, rowNum, 8, 8, "");
            SetValue(sheet, styles["Value"], rowNum, rowNum, 9, 9, "");
            rowNum++;
            SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 4, "Склад Самара");
            SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 5, 5, data.Sum(x => x.ОстаткиНаСкладеЯндекс1));
            SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 6, 6, data.Sum(x => x.ОстаткиНаСкладеЯндекс1 * x.ЦенаЗакуп));
            SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 7, 7, data.Sum(x => x.ОстаткиНаСкладеЯндекс1 * x.ЦенаНаЯндексе));
            SetFormula(sheet, styles["ValueNum"], rowNum, rowNum, 8, 8, "SUM(" + columnЗаказСамара + startInd + ":" + columnЗаказСамара + endInd + ")");
            SetFormula(sheet, styles["ValueMoney"], rowNum, rowNum, 9, 9, "SUMPRODUCT(" + columnЗаказСамара + startInd + ":" + columnЗаказСамара + endInd + "," + columnЦенаЯндекса + startInd + ":" + columnЦенаЯндекса + endInd + ")");
            //rowNum++;
            //SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 4, "Склад Софьино");
            //SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 5, 5, data.Sum(x => x.ОстаткиНаСкладеЯндекс2));
            //SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 6, 6, data.Sum(x => x.ОстаткиНаСкладеЯндекс2 * x.ЦенаЗакуп));
            //SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 7, 7, data.Sum(x => x.ОстаткиНаСкладеЯндекс2 * x.ЦенаНаЯндексе));
            //SetFormula(sheet, styles["ValueNum"], rowNum, rowNum, 8, 8, "SUM(" + columnЗаказСофьино + startInd + ":" + columnЗаказСофьино + endInd + ")");
            //SetFormula(sheet, styles["ValueMoney"], rowNum, rowNum, 9, 9, "SUMPRODUCT(" + columnЗаказСофьино + startInd + ":" + columnЗаказСофьино + endInd + "," + columnЦенаЯндекса + startInd + ":" + columnЦенаЯндекса + endInd + ")");
            //rowNum++;
            //SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 4, "Склад Томилино");
            //SetValue(sheet, styles["ValueNum"], rowNum, rowNum, 5, 5, data.Sum(x => x.ОстаткиНаСкладеЯндекс3));
            //SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 6, 6, data.Sum(x => x.ОстаткиНаСкладеЯндекс3 * x.ЦенаЗакуп));
            //SetValue(sheet, styles["ValueMoney"], rowNum, rowNum, 7, 7, data.Sum(x => x.ОстаткиНаСкладеЯндекс3 * x.ЦенаНаЯндексе));
            //SetFormula(sheet, styles["ValueNum"], rowNum, rowNum, 8, 8, "SUM(" + columnЗаказТомилино + startInd + ":" + columnЗаказТомилино + endInd + ")");
            //SetFormula(sheet, styles["ValueMoney"], rowNum, rowNum, 9, 9, "SUMPRODUCT(" + columnЗаказТомилино + startInd + ":" + columnЗаказТомилино + endInd + "," + columnЦенаЯндекса + startInd + ":" + columnЦенаЯндекса + endInd + ")");
            rowNum++;

            if (unusedSkus != null && unusedSkus.Count > 0)
            {
                rowNum++;
                SetValue(sheet, styles["Value"], rowNum, rowNum, 1, 7, "Обнаружены неизвестные SKU " + string.Join(",", unusedSkus.Select(x => "'" + x + "'")));
            }
        }
        public static byte[] CreateExcelОстаткиФирмы(List<NomenklaturaFirma> data, string firmaName, DateTime period, string extension)
        {
            IWorkbook workbook;
            if (extension == "xls")
                workbook = new HSSFWorkbook();
            else
                workbook = new XSSFWorkbook();

            IFont fontArialBold8 = workbook.CreateFont();
            fontArialBold8.FontName = HSSFFont.FONT_ARIAL;
            fontArialBold8.FontHeightInPoints = 8;
            fontArialBold8.Boldweight = (short)FontBoldWeight.Bold;

            IFont fontArialBold10 = workbook.CreateFont();
            fontArialBold8.FontName = HSSFFont.FONT_ARIAL;
            fontArialBold8.FontHeightInPoints = 10;
            fontArialBold8.Boldweight = (short)FontBoldWeight.Bold;

            IFont fontArial8 = workbook.CreateFont();
            fontArial8.FontName = HSSFFont.FONT_ARIAL;
            fontArial8.FontHeightInPoints = 8;
            fontArial8.Boldweight = (short)FontBoldWeight.Normal;

            ICellStyle styleHeader = workbook.CreateCellStyle();
            styleHeader.WrapText = true;
            styleHeader.VerticalAlignment = VerticalAlignment.Center;
            styleHeader.Alignment = HorizontalAlignment.Center;
            styleHeader.BorderRight = BorderStyle.Medium;
            styleHeader.BorderLeft = BorderStyle.Medium;
            styleHeader.BorderTop = BorderStyle.Medium;
            styleHeader.BorderBottom = BorderStyle.Medium;
            styleHeader.FillPattern = FillPattern.SolidForeground;
            styleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            styleHeader.SetFont(fontArialBold8);

            ICellStyle styleHeader2 = workbook.CreateCellStyle();
            styleHeader2.WrapText = true;
            styleHeader2.VerticalAlignment = VerticalAlignment.Center;
            styleHeader2.Alignment = HorizontalAlignment.Center;
            styleHeader2.BorderRight = BorderStyle.None;
            styleHeader2.BorderLeft = BorderStyle.None;
            styleHeader2.BorderTop = BorderStyle.None;
            styleHeader2.BorderBottom = BorderStyle.None;
            styleHeader2.SetFont(fontArialBold10);

            ICellStyle styleHeaderMoney = workbook.CreateCellStyle();
            styleHeaderMoney.WrapText = true;
            styleHeaderMoney.VerticalAlignment = VerticalAlignment.Center;
            styleHeaderMoney.Alignment = HorizontalAlignment.Right;
            styleHeaderMoney.BorderRight = BorderStyle.Medium;
            styleHeaderMoney.BorderLeft = BorderStyle.Medium;
            styleHeaderMoney.BorderTop = BorderStyle.Medium;
            styleHeaderMoney.BorderBottom = BorderStyle.Medium;
            styleHeaderMoney.FillPattern = FillPattern.SolidForeground;
            styleHeaderMoney.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            styleHeaderMoney.SetFont(fontArialBold8);
            styleHeaderMoney.DataFormat = workbook.CreateDataFormat().GetFormat("# ### ##0.00;-# ### ##0.00;;@");

            ICellStyle styleValue = workbook.CreateCellStyle();
            styleValue.WrapText = true;
            styleValue.VerticalAlignment = VerticalAlignment.Center;
            styleValue.Alignment = HorizontalAlignment.Left;
            styleValue.BorderRight = BorderStyle.Thin;
            styleValue.BorderLeft = BorderStyle.Thin;
            styleValue.BorderTop = BorderStyle.Thin;
            styleValue.BorderBottom = BorderStyle.Thin;
            styleValue.SetFont(fontArial8);

            ICellStyle styleValueNum = workbook.CreateCellStyle();
            styleValueNum.WrapText = true;
            styleValueNum.VerticalAlignment = VerticalAlignment.Center;
            styleValueNum.Alignment = HorizontalAlignment.Right;
            styleValueNum.BorderRight = BorderStyle.Thin;
            styleValueNum.BorderLeft = BorderStyle.Thin;
            styleValueNum.BorderTop = BorderStyle.Thin;
            styleValueNum.BorderBottom = BorderStyle.Thin;
            styleValueNum.SetFont(fontArial8);
            styleValueNum.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.000;-# ##0.000;;@");

            ICellStyle styleValueMoney = workbook.CreateCellStyle();
            styleValueMoney.WrapText = true;
            styleValueMoney.VerticalAlignment = VerticalAlignment.Center;
            styleValueMoney.Alignment = HorizontalAlignment.Right;
            styleValueMoney.BorderRight = BorderStyle.Thin;
            styleValueMoney.BorderLeft = BorderStyle.Thin;
            styleValueMoney.BorderTop = BorderStyle.Thin;
            styleValueMoney.BorderBottom = BorderStyle.Thin;
            styleValueMoney.SetFont(fontArial8);
            styleValueMoney.DataFormat = workbook.CreateDataFormat().GetFormat("# ### ##0.00;-# ### ##0.00;;@");

            Dictionary<string, ICellStyle> styles = new Dictionary<string, ICellStyle>();
            styles.Add("Header", styleHeader);
            styles.Add("Header2", styleHeader2);
            styles.Add("HeaderMoney", styleHeaderMoney);
            styles.Add("Value", styleValue);
            styles.Add("ValueNum", styleValueNum);
            styles.Add("ValueMoney", styleValueMoney);

            //CreateSheetYandexFBY(workbook, sheetName, Styles, data, startDate, endDate, sku, skuUnused);
            ISheet sheet = workbook.CreateSheet("Лист 1");

            sheet.SetColumnWidth(0, 1000);
            sheet.SetColumnWidth(1, 2000); //#
            sheet.SetColumnWidth(2, 2400); //Код
            sheet.SetColumnWidth(3, 4600); //Наименование
            sheet.SetColumnWidth(4, 5200); //ПолнНаименование
            sheet.SetColumnWidth(5, 3300); //Артикул
            sheet.SetColumnWidth(6, 2200); //ВидНоменклатуры
            sheet.SetColumnWidth(7, 2200); //СтавкаНДС
            sheet.SetColumnWidth(8, 2200); //ЕдиницаКод 
            sheet.SetColumnWidth(9, 2200); //ЕдиницаНаименование
            sheet.SetColumnWidth(10, 2200); //СтранаКод
            sheet.SetColumnWidth(11, 3300); //СтранаНаименование
            sheet.SetColumnWidth(12, 2900); //Остаток
            sheet.SetColumnWidth(13, 3300); //Себестоимость
            sheet.SetColumnWidth(14, 2900); //ОстатокПринятый
            sheet.SetColumnWidth(15, 3300); //СебестоимостьПринятый

            SetValue(sheet, styles["Header2"], 0, 0, 1, 15, "Остатки " + firmaName + " на " + period.ToString("dd.MM.yyyy"));

            SetValue(sheet, styles["Header"], 2, 3, 1, 1, "№");
            SetValue(sheet, styles["Header"], 2, 3, 2, 2, "Код");
            SetValue(sheet, styles["Header"], 2, 3, 3, 3, "Наименование");
            SetValue(sheet, styles["Header"], 2, 3, 4, 4, "Полное наименование");
            SetValue(sheet, styles["Header"], 2, 3, 5, 5, "Артикул");
            SetValue(sheet, styles["Header"], 2, 3, 6, 6, "Вид номенклатуры");
            SetValue(sheet, styles["Header"], 2, 3, 7, 7, "Ставка НДС");
            SetValue(sheet, styles["Header"], 2, 2, 8, 9, "Единица");
            SetValue(sheet, styles["Header"], 3, 3, 8, 8, "Код");
            SetValue(sheet, styles["Header"], 3, 3, 9, 9, "Наименование");
            SetValue(sheet, styles["Header"], 2, 2, 10, 11, "Страна");
            SetValue(sheet, styles["Header"], 3, 3, 10, 10, "Код");
            SetValue(sheet, styles["Header"], 3, 3, 11, 11, "Наименование");
            SetValue(sheet, styles["Header"], 2, 2, 12, 13, "Собственный");
            SetValue(sheet, styles["Header"], 3, 3, 12, 12, "Остаток");
            SetValue(sheet, styles["Header"], 3, 3, 13, 13, "Себестоимость");
            SetValue(sheet, styles["Header"], 2, 2, 14, 15, "Принятый (СОХ)");
            SetValue(sheet, styles["Header"], 3, 3, 14, 14, "Остаток");
            SetValue(sheet, styles["Header"], 3, 3, 15, 15, "Себестоимость");

            int rowNum = 4;
            int c = 0;
            foreach (var entry in data.OrderBy(x => x.Наименование))
            {
                IRow row = sheet.CreateRow(rowNum);
                rowNum++;
                c++;
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 1, 1, c.ToString());
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 2, 2, entry.Код);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 3, 3, entry.Наименование);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 4, 4, entry.ПолнНаименование);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 5, 5, entry.Артикул);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 6, 6, entry.ВидНоменклатуры);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 7, 7, entry.СтавкаНДС);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 8, 8, entry.ЕдиницаКод);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 9, 9, entry.ЕдиницаНаименование);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 10, 10, entry.СтранаКод);
                SetValue(sheet, styles["Value"], row.RowNum, row.RowNum, 11, 11, entry.СтранаНаименование);
                SetValue(sheet, styles["ValueNum"], row.RowNum, row.RowNum, 12, 12, entry.Остаток);
                SetValue(sheet, styles["ValueMoney"], row.RowNum, row.RowNum, 13, 13, entry.Себестоимость);
                SetValue(sheet, styles["ValueNum"], row.RowNum, row.RowNum, 14, 14, entry.ОстатокПринятый);
                SetValue(sheet, styles["ValueMoney"], row.RowNum, row.RowNum, 15, 15, entry.СебестоимостьПринятый);
            }
            SetValue(sheet, styles["Header"], rowNum, rowNum, 1, 12, "Итого");
            SetValue(sheet, styles["Header"], rowNum, rowNum, 14, 14, "");
            SetValue(sheet, styles["HeaderMoney"], rowNum, rowNum, 13, 13, data.Sum(x => x.Себестоимость));
            SetValue(sheet, styles["HeaderMoney"], rowNum, rowNum, 15, 15, data.Sum(x => x.СебестоимостьПринятый));

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }
        public static byte[] CreateExcelYandexFBY(List<RowYandexFBY> data,
            DateTime startDate, DateTime endDate,
            List<string> sku, string sheetName,
            List<string> skuUnused,
            string extension)
        {
            IWorkbook workbook;
            IFormulaEvaluator formula;
            if (extension == "xls")
            {
                workbook = new HSSFWorkbook();
                formula = new HSSFFormulaEvaluator(workbook);
            }
            else
            {
                workbook = new XSSFWorkbook();
                formula = new XSSFFormulaEvaluator(workbook);
            }

            IFont fontArialBold8 = workbook.CreateFont();
            fontArialBold8.FontName = HSSFFont.FONT_ARIAL;
            fontArialBold8.FontHeightInPoints = 8;
            fontArialBold8.Boldweight = (short)FontBoldWeight.Bold;

            IFont fontArial8 = workbook.CreateFont();
            fontArial8.FontName = HSSFFont.FONT_ARIAL;
            fontArial8.FontHeightInPoints = 8;
            fontArial8.Boldweight = (short)FontBoldWeight.Normal;

            ICellStyle styleHeader = workbook.CreateCellStyle();
            styleHeader.WrapText = true;
            styleHeader.VerticalAlignment = VerticalAlignment.Center;
            styleHeader.Alignment = HorizontalAlignment.Center;
            styleHeader.BorderRight = BorderStyle.Medium;
            styleHeader.BorderLeft = BorderStyle.Medium;
            styleHeader.BorderTop = BorderStyle.Medium;
            styleHeader.BorderBottom = BorderStyle.Medium;
            styleHeader.FillPattern = FillPattern.SolidForeground;
            styleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            styleHeader.SetFont(fontArialBold8);

            ICellStyle styleValue = workbook.CreateCellStyle();
            styleValue.WrapText = true;
            styleValue.VerticalAlignment = VerticalAlignment.Center;
            styleValue.Alignment = HorizontalAlignment.Left;
            styleValue.BorderRight = BorderStyle.Thin;
            styleValue.BorderLeft = BorderStyle.Thin;
            styleValue.BorderTop = BorderStyle.Thin;
            styleValue.BorderBottom = BorderStyle.Thin;
            styleValue.SetFont(fontArial8);

            ICellStyle styleValueNum = workbook.CreateCellStyle();
            styleValueNum.WrapText = true;
            styleValueNum.VerticalAlignment = VerticalAlignment.Center;
            styleValueNum.Alignment = HorizontalAlignment.Right;
            styleValueNum.BorderRight = BorderStyle.Thin;
            styleValueNum.BorderLeft = BorderStyle.Thin;
            styleValueNum.BorderTop = BorderStyle.Thin;
            styleValueNum.BorderBottom = BorderStyle.Thin;
            styleValueNum.SetFont(fontArial8);
            styleValueNum.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.000;-# ##0.000;;@");

            ICellStyle styleValueNumMark = workbook.CreateCellStyle();
            styleValueNumMark.WrapText = true;
            styleValueNumMark.VerticalAlignment = VerticalAlignment.Center;
            styleValueNumMark.Alignment = HorizontalAlignment.Right;
            styleValueNumMark.BorderRight = BorderStyle.Thin;
            styleValueNumMark.BorderLeft = BorderStyle.Thin;
            styleValueNumMark.BorderTop = BorderStyle.Thin;
            styleValueNumMark.BorderBottom = BorderStyle.Thin;
            styleValueNumMark.FillPattern = FillPattern.SolidForeground;//.AltBars.LessDots;
            styleValueNumMark.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            styleValueNumMark.SetFont(fontArial8);
            styleValueNumMark.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.000;-# ##0.000;;@");

            ICellStyle styleValueMoney = workbook.CreateCellStyle();
            styleValueMoney.WrapText = true;
            styleValueMoney.VerticalAlignment = VerticalAlignment.Center;
            styleValueMoney.Alignment = HorizontalAlignment.Right;
            styleValueMoney.BorderRight = BorderStyle.Thin;
            styleValueMoney.BorderLeft = BorderStyle.Thin;
            styleValueMoney.BorderTop = BorderStyle.Thin;
            styleValueMoney.BorderBottom = BorderStyle.Thin;
            styleValueMoney.SetFont(fontArial8);
            styleValueMoney.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.00;-# ##0.00;;@");

            Dictionary<string, ICellStyle> Styles = new Dictionary<string, ICellStyle>();
            Styles.Add("Header", styleHeader);
            Styles.Add("Value", styleValue);
            Styles.Add("ValueNum", styleValueNum);
            Styles.Add("ValueNumMark", styleValueNumMark);
            Styles.Add("ValueMoney", styleValueMoney);

            CreateSheetYandexFBY(workbook, sheetName, Styles, data, startDate, endDate, sku, skuUnused);

            formula.EvaluateAll();

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }
        public static byte[] CreateExcelFromTemplate(string templateName, List<nDataEntry> Data)
        {
            string extension = Path.GetExtension(templateName).ToLower();
            //IWorkbook workbook;
            IWorkbook template;
            using (var tmpFile = new FileStream(HttpContext.Current.Server.MapPath("~/bin/Templates/") + templateName, FileMode.Open, FileAccess.Read))
            {
                if (extension == ".xls")
                {
                    //workbook = new HSSFWorkbook();
                    template = new HSSFWorkbook(tmpFile);
                }
                else
                {
                    //workbook = new XSSFWorkbook();
                    template = new XSSFWorkbook(tmpFile);
                }
            }
            //ISheet ws = workbook.CreateSheet(SheetName);

            //for (int i = 0; i < Data.Length; i++)
            //{
            //    int namedCellIdx = template.GetNameIndex(Data[i].Имя);
            //    IName aNamedCell = template.GetNameAt(namedCellIdx);

            //    AreaReference aref = new AreaReference(aNamedCell.RefersToFormula);
            //    CellReference[] crefs = aref.GetAllReferencedCells();
            //    //ISheet sourceWS = template.GetSheet(aref.FirstCell.SheetName);
            //    //for (int j = aref.FirstCell.Row; j <= aref.LastCell.Row; j++)
            //    //{
            //    //    CopyRow(template, sourceWS, j, workbook, ws, j);
            //    //}
            //    //CellReference[] crefs = aref.GetAllReferencedCells();
            //    for (int j = 0; j < crefs.Length; j++)
            //    {
            //        ISheet s = template.GetSheet(crefs[j].SheetName);
            //        IRow r = s.GetRow(crefs[j].Row);
            //        if (r != null)
            //        {
            //            ICell c = r.GetCell(crefs[j].Col);
            //            CopyCell(c, ws, 0);
            //        }
            //        //c.SetCellValue(Data[i].Значение);
            //    }
            //}
            string sheetName = "";
            List<int> nRowNum = new List<int>();
            List<DataEntry> staticVariants = Data.Find(x => x.мнИмя == "" && x.мнЗначения != null).мнЗначения[0];
            foreach (var entries in Data.Where(x => x.мнИмя != "" && x.мнЗначения != null))
            {
                int namedCellIdx = template.GetNameIndex(entries.мнИмя);
                IName aNamedCell = template.GetNameAt(namedCellIdx);

                AreaReference aref = new AreaReference(aNamedCell.RefersToFormula);
                nRowNum.Add(aref.FirstCell.Row);
                sheetName = aref.FirstCell.SheetName;
            }
            //int namedCellIdx2 = template.GetNameIndex("мнТаблицы");
            //IName aNamedCell2 = template.GetNameAt(namedCellIdx2);

            //AreaReference aref2 = new AreaReference(aNamedCell2.RefersToFormula);
            //CellReference[] crefs2 = aref2.GetAllReferencedCells();
            ISheet s;
            if (string.IsNullOrEmpty(sheetName))
                s = template.GetSheetAt(0);
            else
                s = template.GetSheet(sheetName);
            //ISheet s = template.GetSheet(crefs2[0].SheetName);
            //IRow r = s.GetRow(crefs2[0].Row);
            //ICell c = r.GetCell(crefs2[0].Col);
            //nRowNum = crefs2[0].Row;
            //CopyRow(template, s, r.RowNum, r.RowNum + 1);
            for (int i = s.FirstRowNum; i <= s.LastRowNum; i++)
            {
                if (!nRowNum.Contains(i))
                {
                    IRow row = s.GetRow(i);
                    if ((row != null) && (row.Cells.Count > 0))
                    {
                        for (int j = row.FirstCellNum; j <= row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null && cell.CellType == CellType.String)
                            {
                                string cellText = cell.StringCellValue; //GetStringCellValue(cell);
                                var variants = Regex.Matches(cellText, @"(\A|\s*)\{(\w+)\}(\s*|\z)")
                                .Cast<Match>()
                                .Select(m => m.Value.Trim());
                                foreach (string variant in variants)
                                {
                                    DataEntry entry = staticVariants.Find(x => "{" + x.Имя + "}" == variant);
                                    if (entry != null)
                                    {
                                        cellText = cellText.Replace(variant, entry.Значение);
                                        //cell.SetCellValue(cellText.Replace(variant, entry.Значение));
                                    }
                                }
                                cell.SetCellValue(cellText);
                            }
                        }
                    }
                }
            }

            //nRowNum.Sort();
            //nRowNum.Reverse();
            //foreach (int i in nRowNum)
            //{
            //    //IRow r = s.GetRow(i);
            //    CopyRow(template, s, i, 1);
            //}
            foreach (var entries in Data.Where(x => x.мнИмя != "" && x.мнЗначения != null))
            {
                int namedCellIdx = template.GetNameIndex(entries.мнИмя);
                IName aNamedCell = template.GetNameAt(namedCellIdx);

                AreaReference aref = new AreaReference(aNamedCell.RefersToFormula);
                CopyRow(template, s, aref.FirstCell.Row, entries.мнЗначения);
                IRow row = s.GetRow(aref.FirstCell.Row);
                if (row != null)
                {
                    s.RemoveRow(row);
                    row = s.CreateRow(aref.FirstCell.Row);
                    row.ZeroHeight = true;
                    //s.ShiftRows(aref.FirstCell.Row + 1, s.LastRowNum, -1, false, false);
                }
            }

            using (var stream = new MemoryStream())
            {
                template.Write(stream);
                return stream.ToArray();
            }
        }
    }
}