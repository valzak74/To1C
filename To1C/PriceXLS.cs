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

namespace To1C
{
    public class PriceXLS
    {

        private static void SetHeader(IWorkbook workbook, ISheet ws, int column, IFont font, int width, string value, bool IsSpec = false)
        {
            ws.SetColumnWidth(column, width);
            ICellStyle styleHeader = workbook.CreateCellStyle();
            styleHeader.WrapText = true;
            styleHeader.VerticalAlignment = VerticalAlignment.Center;
            styleHeader.Alignment = HorizontalAlignment.Center;
            styleHeader.BorderRight = BorderStyle.Medium;
            styleHeader.BorderLeft = BorderStyle.Medium;
            styleHeader.BorderTop = BorderStyle.Medium;
            styleHeader.BorderBottom = BorderStyle.Medium;
            if (IsSpec)
            {
                // cell background
                //styleHeader.FillForegroundColor = GetXLColour(Color.FromArgb(255, 255, 192));
                styleHeader.FillPattern = FillPattern.SolidForeground;
                styleHeader.FillForegroundColor = HSSFColor.LemonChiffon.Index;
            }
            styleHeader.SetFont(font);

            IRow rowHeader = ws.GetRow(4);
            ICell cell = rowHeader.CreateCell(column);
            cell.CellStyle = styleHeader;
            cell.SetCellValue(value);
            ws.SetColumnWidth(column, width);
        }

        public static byte[] CreateXLS(List<string> UsedID, List<string> SklNames, int harakter, int opt, int osob, int spec, int rozn, int zakaz, int svoya, Price[] Prices, string template, string Address, int OtsrDney, int ShowOstatok)
        {
            string extension = Path.GetExtension(template);

            IWorkbook workbook;
            using (var xlsxFile = new FileStream(template, FileMode.Open, FileAccess.Read))
            {
                if (extension == ".xls")
                    workbook = new HSSFWorkbook(xlsxFile);
                else
                    workbook = new XSSFWorkbook(xlsxFile);
            }

            ISheet ws = workbook.GetSheetAt(0);

            ws.GetRow(1).GetCell(1).SetCellValue(Address);
            ws.GetRow(2).GetCell(1).SetCellValue(String.Format(ws.GetRow(2).GetCell(1).StringCellValue, DateTime.Now.ToString("dd MMMM yyyy")));
            if (svoya == 1)
            {
                ws.GetRow(3).CreateCell(1).SetCellValue("* - " + (OtsrDney <= 0 ? "предоплата" : "отсрочка " + OtsrDney.ToString() + " дней"));
                //cell.CellStyle = styleHeader;
                //cell.SetCellValue(value);
                //ws.SetColumnWidth(column, width);
                //ws.GetRow(3).GetCell(1).SetCellValue("* - " + (OtsrDney <= 0 ? "предоплата" : "отсрочка " + OtsrDney.ToString() + " дней"));
            }
            // font 
            IFont fontArialBold10 = workbook.CreateFont();
            fontArialBold10.FontName = HSSFFont.FONT_ARIAL;
            fontArialBold10.FontHeightInPoints = 10;
            fontArialBold10.Boldweight = (short)FontBoldWeight.Bold;
            IFont fontArialBold8 = workbook.CreateFont();
            fontArialBold8.FontName = HSSFFont.FONT_ARIAL;
            fontArialBold8.FontHeightInPoints = 8;
            fontArialBold8.Boldweight = (short)FontBoldWeight.Bold;
            IFont fontArial8 = workbook.CreateFont();
            fontArial8.FontName = HSSFFont.FONT_ARIAL;
            fontArial8.FontHeightInPoints = 8;
            fontArial8.Boldweight = (short)FontBoldWeight.Normal;

            int NumHarak = harakter * 4;
            int NumSvoya = svoya * (4 + harakter);
            int NumOpt = opt * (4 + harakter + svoya);
            int NumSpec = spec * (4 + harakter + svoya + opt);
            int NumOsob = osob * (4 + harakter + svoya + opt + spec);
            int NumRozn = rozn * (4 + harakter + svoya + opt + spec + osob);
            int NumSpecR = rozn * spec + NumRozn;
            int NumSklFirst = 4 + harakter + svoya + opt + spec + osob + rozn;
            if (harakter == 1)
                SetHeader(workbook, ws, NumHarak, fontArialBold10, 6000, "Характеристики");
            if (svoya == 1)
                SetHeader(workbook, ws, NumSvoya, fontArialBold10, 3900, "Ваша цена*");
            if (opt == 1)
                SetHeader(workbook, ws, NumOpt, fontArialBold10, 3900, "Оптовая цена");
            if (spec == 1)
                SetHeader(workbook, ws, NumSpec, fontArialBold10, 3900, "Оптовая Спец цена", true);
            if (osob == 1)
                SetHeader(workbook, ws, NumOsob, fontArialBold10, 3900, "Особая цена");
            if (rozn == 1)
            {
                SetHeader(workbook, ws, NumRozn, fontArialBold10, 3900, "Розничная цена");
                if (spec == 1)
                {
                    NumSklFirst += 1;
                    SetHeader(workbook, ws, NumSpecR, fontArialBold10, 3900, "Розничная Спец цена", true);
                }
            }
            int SklNum = 0;
            foreach (string SklName in SklNames)
            {
                SetHeader(workbook, ws, SklNum + NumSklFirst, fontArialBold10, 3800, (SklNames.Count == 1 ? "Наличие" : SklName));
                SklNum++;
            }
            //SetHeader(workbook, ws, SklNum + NumSklFirst, fontArialBold10, 3800, "Наличие на других складах");
            if (zakaz == 1)
                SetHeader(workbook, ws, SklNum + NumSklFirst, fontArialBold10, 3800, "Ожидаемый приход");

            ICellStyle styleGroup = workbook.CreateCellStyle();
            styleGroup.WrapText = true;
            styleGroup.VerticalAlignment = VerticalAlignment.Center;
            styleGroup.Alignment = HorizontalAlignment.Left;
            styleGroup.BorderRight = BorderStyle.Medium;
            styleGroup.BorderLeft = BorderStyle.Medium;
            styleGroup.BorderTop = BorderStyle.Medium;
            styleGroup.BorderBottom = BorderStyle.Medium;
            styleGroup.FillPattern = FillPattern.SolidForeground;
            styleGroup.FillForegroundColor = HSSFColor.Plum.Index;
            styleGroup.SetFont(fontArialBold8);

            ICellStyle styleValue = workbook.CreateCellStyle();
            styleValue.WrapText = true;
            styleValue.VerticalAlignment = VerticalAlignment.Center;
            styleValue.Alignment = HorizontalAlignment.Left;
            styleValue.BorderRight = BorderStyle.Thin;
            styleValue.BorderLeft = BorderStyle.Thin;
            styleValue.BorderTop = BorderStyle.Thin;
            styleValue.BorderBottom = BorderStyle.Thin;
            styleValue.SetFont(fontArial8);

            ICellStyle styleValueUsed = workbook.CreateCellStyle();
            styleValueUsed.WrapText = true;
            styleValueUsed.VerticalAlignment = VerticalAlignment.Center;
            styleValueUsed.Alignment = HorizontalAlignment.Left;
            styleValueUsed.BorderRight = BorderStyle.Thin;
            styleValueUsed.BorderLeft = BorderStyle.Thin;
            styleValueUsed.BorderTop = BorderStyle.Thin;
            styleValueUsed.BorderBottom = BorderStyle.Thin;
            styleValueUsed.FillPattern = FillPattern.SolidForeground;
            styleValueUsed.FillForegroundColor = HSSFColor.LightCornflowerBlue.Index;
            styleValueUsed.SetFont(fontArial8);

            ICellStyle styleValueRight = workbook.CreateCellStyle();
            styleValueRight.WrapText = true;
            styleValueRight.VerticalAlignment = VerticalAlignment.Center;
            styleValueRight.Alignment = HorizontalAlignment.Right;
            styleValueRight.BorderRight = BorderStyle.Thin;
            styleValueRight.BorderLeft = BorderStyle.Thin;
            styleValueRight.BorderTop = BorderStyle.Thin;
            styleValueRight.BorderBottom = BorderStyle.Thin;
            styleValueRight.SetFont(fontArial8);

            ICellStyle styleValueNum = workbook.CreateCellStyle();
            styleValueNum.WrapText = true;
            styleValueNum.VerticalAlignment = VerticalAlignment.Center;
            styleValueNum.Alignment = HorizontalAlignment.Right;
            styleValueNum.BorderRight = BorderStyle.Thin;
            styleValueNum.BorderLeft = BorderStyle.Thin;
            styleValueNum.BorderTop = BorderStyle.Thin;
            styleValueNum.BorderBottom = BorderStyle.Thin;
            styleValueNum.SetFont(fontArial8);
            styleValueNum.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.00");

            ICellStyle styleValueSp = workbook.CreateCellStyle();
            styleValueSp.WrapText = true;
            styleValueSp.VerticalAlignment = VerticalAlignment.Center;
            styleValueSp.Alignment = HorizontalAlignment.Right;
            styleValueSp.BorderRight = BorderStyle.Thin;
            styleValueSp.BorderLeft = BorderStyle.Thin;
            styleValueSp.BorderTop = BorderStyle.Thin;
            styleValueSp.BorderBottom = BorderStyle.Thin;
            styleValueSp.FillPattern = FillPattern.SolidForeground;
            styleValueSp.FillForegroundColor = HSSFColor.LemonChiffon.Index;
            styleValueSp.SetFont(fontArial8);
            styleValueSp.DataFormat = workbook.CreateDataFormat().GetFormat("# ##0.00");

            int row = 5;
            try
            {
                foreach (Price pr in Prices)
                {
                    IRow rowTable = ws.CreateRow(row);
                    rowTable.HeightInPoints = 11;
                    if (pr.IsFolder == 1)
                    {
                        for (int i = 1; i <= SklNum + NumSklFirst + zakaz - 1; i++)
                        {
                            ICell cellGroup = rowTable.CreateCell(i);
                            cellGroup.CellStyle = styleGroup;
                            cellGroup.SetCellValue(String.Empty.PadLeft(pr.level * 2) + pr.Nomenklatura);
                        }
                        CellRangeAddress cra = new CellRangeAddress(row, row, 1, SklNum + NumSklFirst + zakaz - 1);
                        ws.AddMergedRegion(cra);
                    }
                    else
                    {
                        rowTable.CreateCell(1).SetCellType(CellType.String);
                        rowTable.CreateCell(2).SetCellType(CellType.String);
                        rowTable.CreateCell(3).SetCellType(CellType.String);
                        rowTable.GetCell(1).SetCellValue(pr.Articul);
                        rowTable.GetCell(2).SetCellValue(pr.Nomenklatura);
                        rowTable.GetCell(3).SetCellValue(pr.Brend);
                        if (NumHarak > 0)
                        {
                            rowTable.CreateCell(NumHarak).SetCellType(CellType.String);
                            rowTable.GetCell(NumHarak).SetCellValue(pr.Harakter);
                        }
                        if (NumSvoya > 0)
                        {
                            rowTable.CreateCell(NumSvoya).SetCellType(CellType.Numeric);
                            rowTable.GetCell(NumSvoya).SetCellValue(pr.CostClient);
                        }
                        if (NumOpt > 0)
                        {
                            rowTable.CreateCell(NumOpt).SetCellType(CellType.Numeric);
                            rowTable.GetCell(NumOpt).SetCellValue(pr.Cost);
                        }
                        if (NumSpec > 0)
                        {
                            if (pr.SuperCost > 0)
                            {
                                rowTable.CreateCell(NumSpec).SetCellType(CellType.Numeric);
                                rowTable.GetCell(NumSpec).SetCellValue(pr.SuperCost);
                            }
                            else
                            {
                                rowTable.CreateCell(NumSpec).SetCellType(CellType.String);
                                rowTable.GetCell(NumSpec).SetCellValue("");
                            }
                        }
                        if (NumOsob > 0)
                        {
                            rowTable.CreateCell(NumOsob).SetCellType(CellType.Numeric);
                            rowTable.GetCell(NumOsob).SetCellValue(pr.CostOsob);
                        }
                        if (NumRozn > 0)
                        {
                            rowTable.CreateCell(NumRozn).SetCellType(CellType.Numeric);
                            rowTable.GetCell(NumRozn).SetCellValue(pr.CostRozn);
                        }
                        if (NumSpecR > 0)
                        {
                            if (pr.SuperCostR > 0)
                            {
                                rowTable.CreateCell(NumSpecR).SetCellType(CellType.Numeric);
                                rowTable.GetCell(NumSpecR).SetCellValue(pr.SuperCostR);
                            }
                            else
                            {
                                rowTable.CreateCell(NumSpecR).SetCellType(CellType.String);
                                rowTable.GetCell(NumSpecR).SetCellValue("");
                            }
                        }
                        int s_number = 0;
                        foreach (double OstSkl in pr.OstatokSkl)
                        {
                            rowTable.CreateCell(s_number + NumSklFirst).SetCellType((ShowOstatok == 1 ? CellType.Numeric : CellType.String));
                            if (ShowOstatok == 1)
                            {
                                if (pr.OstatokSkl[s_number] > 0)
                                {
                                    rowTable.CreateCell(s_number + NumSklFirst).SetCellType(CellType.Numeric);
                                    rowTable.GetCell(s_number + NumSklFirst).SetCellValue(pr.OstatokSkl[s_number]);
                                }
                                else
                                {
                                    rowTable.CreateCell(s_number + NumSklFirst).SetCellType(CellType.String);
                                    rowTable.GetCell(s_number + NumSklFirst).SetCellValue("");
                                }
                            }
                            else
                            {
                                rowTable.CreateCell(s_number + NumSklFirst).SetCellType(CellType.String);
                                rowTable.GetCell(s_number + NumSklFirst).SetCellValue(GetOstatokAsMark(pr.OstatokSkl[s_number], pr.OstatokPrihod, pr.NeNadoZakup));
                            }
                            s_number++;
                        }
                        //rowTable.CreateCell(s_number + NumSklFirst).SetCellType(CellType.String);
                        //rowTable.GetCell(s_number + NumSklFirst).SetCellValue(GetOstatokAsMark(pr.Ostatok));
                        if (zakaz == 1)
                        {
                            if (pr.OstatokPrihod > 0)
                            {
                                rowTable.CreateCell(s_number + NumSklFirst).SetCellType(CellType.Numeric);
                                rowTable.GetCell(s_number + NumSklFirst).SetCellValue(pr.OstatokPrihod);
                            }
                            else
                            {
                                rowTable.CreateCell(s_number + NumSklFirst).SetCellType(CellType.String);
                                rowTable.GetCell(s_number + NumSklFirst).SetCellValue("");
                            }
                        }
                        bool used = UsedID.Contains(pr.ID);
                        for (int i = 1; i <= SklNum + NumSklFirst + zakaz - 1; i++)
                        {
                            ICell cell = rowTable.GetCell(i);
                            if (cell.CellType == CellType.Numeric)
                            {
                                if (((i == NumSpec) | (i == NumSpecR)) && (cell.NumericCellValue > 0))
                                    cell.CellStyle = styleValueSp;
                                else
                                    cell.CellStyle = styleValueNum;
                            }
                            else
                            {
                                if (i >= SklNum + NumSklFirst + zakaz - 1)
                                    cell.CellStyle = styleValueRight;
                                else
                                    cell.CellStyle = (used ? styleValueUsed : styleValue);
                            }
                        }
                    }
                    row++;
                }
            }
            catch (Exception e)
            {
                string s = e.Message;
            }

            //using (var fs = new FileStream(@"f:\tmp\tdf.xlsx", FileMode.Create, FileAccess.Write))
            //{
            //    workbook.Write(fs);
            //}
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }


        private static string GetOstatokAsMark(double Ostatok, double OstatokPrihod, bool NeNadoZakup)
        {
            if (Ostatok > 0)
                return "в наличии";
            else if ((OstatokPrihod > 0) && !NeNadoZakup)
                return "в пути";
            else if (!NeNadoZakup)
                return "ожидается";
            else
                return "под заказ";
        }
    }

}