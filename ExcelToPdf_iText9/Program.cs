using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Layout.Borders;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.IO.Font;


namespace ExcelToPdf_iText9
{
    internal class Program
    {
        private static readonly string NormalFontPath = @"C:\Windows\Fonts\msjh.ttc,0";   // 微軟正黑體 Regular
        private static readonly string BoldFontPath = @"C:\Windows\Fonts\msjhbd.ttc,0"; // 微軟正黑體 Bold

        private static iText.Kernel.Font.PdfFont normalFont;
        private static iText.Kernel.Font.PdfFont boldFont;



        static void Main(string[] args)
        {
            foreach (var file in Directory.GetFiles(@"C:\Windows\Fonts"))
            {
                Console.WriteLine(file);
            }
            // 讀取字型檔
            var fontProgram = FontProgramFactory.CreateFont(NormalFontPath, true);
            var normalFont = PdfFontFactory.CreateFont(fontProgram, PdfEncodings.IDENTITY_H,
                                                       PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
            // 正確載入 (強制嵌入，避免 NullReference)
            normalFont = PdfFontFactory.CreateFont(NormalFontPath, PdfEncodings.IDENTITY_H,
                                                   iText.Kernel.Font.PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
            boldFont = PdfFontFactory.CreateFont(BoldFontPath, PdfEncodings.IDENTITY_H,
                                                   iText.Kernel.Font.PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);

            Console.WriteLine("字型載入成功");

            string excelPath = @"C:\Users\Evan\Desktop\採購單\採購單範本.xlsx";
            string pdfPath = @"C:\Users\Evan\Desktop\採購單\採購單範本_out.pdf";

            ConvertExcelToPdf(excelPath, pdfPath);
            Console.WriteLine("PDF 已完成: " + pdfPath);
        }
        static void ConvertExcelToPdf(string excelPath, string pdfPath)
        {
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook wb = WorkbookFactory.Create(fs);
                ISheet sheet = wb.GetSheetAt(0);

                using (PdfWriter writer = new PdfWriter(pdfPath))
                using (PdfDocument pdf = new PdfDocument(writer))
                using (Document doc = new Document(pdf))
                {
                    int maxCols = GetMaxCols(sheet, 200);

                    // Excel 欄寬 → PDF 欄寬
                    float[] colWidths = new float[maxCols];
                    for (int c = 0; c < maxCols; c++)
                    {
                        int excelWidth = sheet.GetColumnWidth(c); // 單位 1/256
                        colWidths[c] = excelWidth / 256f * 7f;
                        if (colWidths[c] < 20f) colWidths[c] = 20f;
                    }

                    Table table = new Table(iText.Layout.Properties.UnitValue.CreatePointArray(colWidths));

                    // 合併儲存格
                    var mergedRegions = new List<CellRangeAddress>();
                    for (int i = 0; i < sheet.NumMergedRegions; i++)
                        mergedRegions.Add(sheet.GetMergedRegion(i));

                    int maxRows = Math.Min(sheet.LastRowNum + 1, 200);

                    for (int r = 0; r < maxRows; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        float rowHeightPt = 15f;
                        if (row != null && row.Height > 0)
                            rowHeightPt = row.Height / 20f;

                        for (int c = 0; c < maxCols; c++)
                        {
                            CellRangeAddress merged = IsMergedRegion(r, c, mergedRegions);
                            if (merged != null)
                            {
                                if (merged.FirstRow == r && merged.FirstColumn == c)
                                {
                                    ICell excelCell = row?.GetCell(c);
                                    string text = FormatExcelCell(excelCell, wb);
                                    var pdfCell = CreatePdfCell(text, excelCell, wb,
                                        merged.LastRow - merged.FirstRow + 1,
                                        merged.LastColumn - merged.FirstColumn + 1);

                                    pdfCell.SetHeight(rowHeightPt);
                                    table.AddCell(pdfCell);
                                }
                            }
                            else
                            {
                                ICell excelCell = row?.GetCell(c);
                                string text = FormatExcelCell(excelCell, wb);
                                var pdfCell = CreatePdfCell(text, excelCell, wb, 1, 1);
                                pdfCell.SetHeight(rowHeightPt);
                                table.AddCell(pdfCell);
                            }
                        }
                    }

                    doc.Add(new Paragraph("Excel 轉 PDF (iText 9.2.0, PdfFont 控制粗體)")
                            .SetFont(boldFont).SetFontSize(16));
                    doc.Add(table);

                    // Excel 圖片 (XSSF)
                    if (sheet is XSSFSheet xsheet)
                    {
                        foreach (var rel in xsheet.GetRelations())
                        {
                            if (rel is NPOI.XSSF.UserModel.XSSFDrawing drawing)
                            {
                                foreach (var shape in drawing.GetShapes())
                                {
                                    if (shape is NPOI.XSSF.UserModel.XSSFPicture pic)
                                    {
                                        var pdata = pic.PictureData;
                                        byte[] bytes = pdata.Data;
                                        var img = new Image(ImageDataFactory.Create(bytes));
                                        img.ScaleToFit(200, 200);
                                        doc.Add(new Paragraph("Excel 內嵌圖片:").SetFont(normalFont));
                                        doc.Add(img);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Excel 數值格式化
        static string FormatExcelCell(ICell cell, IWorkbook wb)
        {
            if (cell == null) return "";

            var evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            if (cell.CellType == CellType.Formula)
                cell = evaluator.EvaluateInCell(cell);

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    else
                    {
                        short fmtIdx = cell.CellStyle.DataFormat;
                        string fmt = wb.CreateDataFormat().GetFormat(fmtIdx);
                        if (!string.IsNullOrEmpty(fmt) && fmt.Contains("%"))
                            return (cell.NumericCellValue * 100).ToString("0.##") + "%";
                        return cell.NumericCellValue.ToString("0.##");
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.String:
                    return cell.StringCellValue;
                default:
                    return cell.ToString();
            }
        }

        // 建立 PDF Cell
        static iText.Layout.Element.Cell CreatePdfCell(string text, ICell excelCell, IWorkbook wb, int rowspan, int colspan)
        {
            var paragraph = new Paragraph(text).SetFontSize(10).SetFont(normalFont);

            var cell = new iText.Layout.Element.Cell(rowspan, colspan)
                .Add(paragraph)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(0.5f))
                .SetPadding(3);

            if (excelCell != null)
            {
                ICellStyle style = excelCell.CellStyle;
                if (style != null)
                {
                    // 水平對齊
                    switch (style.Alignment)
                    {
                        case NPOI.SS.UserModel.HorizontalAlignment.Center:
                            paragraph.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
                            break;
                        case NPOI.SS.UserModel.HorizontalAlignment.Right:
                            paragraph.SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT);
                            break;
                        default:
                            paragraph.SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT);
                            break;
                    }

                    // 垂直對齊
                    switch (style.VerticalAlignment)
                    {
                        case NPOI.SS.UserModel.VerticalAlignment.Top:
                            cell.SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.TOP);
                            break;
                        case NPOI.SS.UserModel.VerticalAlignment.Bottom:
                            cell.SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.BOTTOM);
                            break;
                        default:
                            cell.SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE);
                            break;
                    }

                    // 判斷是否粗體
                    IFont font = wb.GetFontAt(style.FontIndex);
                    if (font != null && font.IsBold)
                        paragraph.SetFont(boldFont);

                    // 背景色
                    if (style.FillPattern == FillPattern.SolidForeground)
                    {
                        if (style.FillForegroundColorColor is XSSFColor xcolor && xcolor.RGB != null)
                        {
                            var rgb = xcolor.RGB;
                            cell.SetBackgroundColor(new DeviceRgb(rgb[0], rgb[1], rgb[2]));
                        }
                    }
                }
            }
            return cell;
        }

        // 判斷合併儲存格
        static CellRangeAddress IsMergedRegion(int row, int col, List<CellRangeAddress> mergedRegions)
        {
            foreach (var range in mergedRegions)
            {
                if (row >= range.FirstRow && row <= range.LastRow &&
                    col >= range.FirstColumn && col <= range.LastColumn)
                    return range;
            }
            return null;
        }

        static int GetMaxCols(ISheet sheet, int maxRowsToCheck)
        {
            int max = 0;
            int rows = Math.Min(sheet.LastRowNum + 1, maxRowsToCheck);
            for (int r = 0; r < rows; r++)
            {
                var row = sheet.GetRow(r);
                if (row != null) max = Math.Max(max, row.LastCellNum);
            }
            return max;
        }
    }
}
