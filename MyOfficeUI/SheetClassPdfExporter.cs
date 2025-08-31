using MyOffice;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using PdfSharp.Fonts;
using System.Collections.Generic;
using PdfSharp;
using System.Linq;
using System.Drawing;

public static class SheetClassPdfExporter
{
    /// <summary>
    /// 儲存多個 SheetClass 為 PDF 檔案
    /// </summary>
    public static void SaveToPDF(this List<SheetClass> sheets,
                                  string filePath,
                                  PageSize pageSize = PageSize.A4,
                                  PageOrientation orientation = PageOrientation.Portrait,
                                  double marginMm = 10,
                                  double textPaddingHorizontalMm = 0.5,
                                  double textPaddingVerticalMm = 0.3,
                                  bool enableLog = true)
    {
        byte[] pdfBytes = sheets.SaveToPDF(pageSize, orientation, marginMm,
                                           textPaddingHorizontalMm, textPaddingVerticalMm,
                                           enableLog);

        File.WriteAllBytes(filePath, pdfBytes);
    }

    /// <summary>
    /// 產生 PDF 並輸出為 byte[] (多個 SheetClass = 多頁 PDF)
    /// </summary>
    public static byte[] SaveToPDF(this List<SheetClass> sheets,
                                   PageSize pageSize = PageSize.A4,
                                   PageOrientation orientation = PageOrientation.Portrait,
                                   double marginMm = 10,
                                   double textPaddingHorizontalMm = 0.5,
                                   double textPaddingVerticalMm = 0.3,
                                   bool enableLog = true)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        using (PdfDocument doc = new PdfDocument())
        {
            foreach (var sheet in sheets)
            {
                PdfPage page = doc.AddPage();
                page.Size = pageSize;
                page.Orientation = orientation;
                GlobalFontSettings.FontResolver = new CustomFontResolver();
                XGraphics g = XGraphics.FromPdfPage(page);

                // mm → pt
                double mmToPt(double mm) => mm / 25.4 * 72.0;
                double margin = mmToPt(marginMm);

                double usableWidth = page.Width - margin * 2;
                double usableHeight = page.Height - margin * 2;

                // 內容原始大小
                int rawWidth = sheet.ColumnsWidth.Sum(w => (int)((w + 5) / 256.0 * 7));
                int rawHeight = sheet.Rows.Sum(r => (int)(r.Height / 15.0 * 20 * 96.0 / 72.0));

                // 建立 XForm (虛擬畫布)
                XForm form = new XForm(doc, usableWidth, usableHeight);
                using (XGraphics gForm = XGraphics.FromForm(form))
                {
                    double scaleX = usableWidth / rawWidth;
                    double scaleY = usableHeight / rawHeight;
                    double scaleText = Math.Min(scaleX, scaleY); // ✅ 字體等比例縮放

                    sheet.DrawToGraphics(gForm, scaleX, scaleY, scaleText,
                                         textPaddingHorizontalMm, textPaddingVerticalMm);
                }

                g.DrawImage(form, margin, margin, usableWidth, usableHeight);
            }

            using (MemoryStream ms = new MemoryStream())
            {
                doc.Save(ms, false);

                sw.Stop();
                if (enableLog)
                {
                    Console.WriteLine($"SaveToPDF 多工作表耗時: {sw.ElapsedMilliseconds} ms, 共 {sheets.Count} 頁");
                }

                return ms.ToArray();
            }
        }
    }

    public static void SaveToPDF(this SheetClass sheet,
                                 string filePath,
                                 PageSize pageSize = PageSize.A4,
                                 PageOrientation orientation = PageOrientation.Portrait,
                                 double marginMm = 10,
                                 double textPaddingHorizontalMm = 0.5, // 左右文字邊距
                                 double textPaddingVerticalMm = 0.3)   // 上下文字邊距
    {
        // 呼叫共用方法產生 bytes
        byte[] pdfBytes = sheet.SaveToPDF(pageSize, orientation, marginMm,
                                          textPaddingHorizontalMm, textPaddingVerticalMm);

        // 存檔
        File.WriteAllBytes(filePath, pdfBytes);
    }

    /// <summary>
    /// 產生 PDF 並輸出為 byte[]
    /// </summary>
    public static byte[] SaveToPDF(this SheetClass sheet,
                                   PageSize pageSize = PageSize.A4,
                                   PageOrientation orientation = PageOrientation.Portrait,
                                   double marginMm = 10,
                                   double textPaddingHorizontalMm = 0.5,
                                   double textPaddingVerticalMm = 0.3,
                                   bool enableLog = true)   // 👈 控制是否顯示耗時
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        using (PdfDocument doc = new PdfDocument())
        {
            PdfPage page = doc.AddPage();
            page.Size = pageSize;
            page.Orientation = orientation;
            GlobalFontSettings.FontResolver = new CustomFontResolver();
            XGraphics g = XGraphics.FromPdfPage(page);

            // mm → pt
            double mmToPt(double mm) => mm / 25.4 * 72.0;
            double margin = mmToPt(marginMm);

            double usableWidth = page.Width - margin * 2;
            double usableHeight = page.Height - margin * 2;

            // 內容原始大小
            int rawWidth = sheet.ColumnsWidth.Sum(w => (int)((w + 5) / 256.0 * 7));
            int rawHeight = sheet.Rows.Sum(r => (int)(r.Height / 15.0 * 20 * 96.0 / 72.0));

            // 建立 XForm (虛擬畫布)
            XForm form = new XForm(doc, usableWidth, usableHeight);
            using (XGraphics gForm = XGraphics.FromForm(form))
            {
                double scaleX = usableWidth / rawWidth;
                double scaleY = usableHeight / rawHeight;
                double scaleText = Math.Min(scaleX, scaleY); // ✅ 字體等比例縮放

                sheet.DrawToGraphics(gForm, scaleX, scaleY, scaleText,
                                     textPaddingHorizontalMm, textPaddingVerticalMm);
            }

            g.DrawImage(form, margin, margin, usableWidth, usableHeight);

            // ✅ 存到記憶體
            using (MemoryStream ms = new MemoryStream())
            {
                doc.Save(ms, false);

                sw.Stop();
                if (enableLog)
                {
                    Console.WriteLine($"SaveToPDF 耗時: {sw.ElapsedMilliseconds} ms");
                }

                return ms.ToArray();
            }
        }
    }



    /// <summary>
    /// 繪製 SheetClass 到指定 XGraphics
    /// - 背景、圖片、邊框 → 照 ScaleX / ScaleY 拉伸
    /// - 文字 → 等比例縮放，並可設定水平/垂直 padding (mm)
    /// </summary>
    public static void DrawToGraphics(this SheetClass sheet, XGraphics g,
                                       double scaleX = 1.0, double scaleY = 1.0, double scaleText = 1.0,
                                       double textPaddingHorizontalMm = 0.5,
                                       double textPaddingVerticalMm = 0.3)
    {
        // mm → pt
        double mmToPt(double mm) => mm / 25.4 * 72.0;
        double padX = mmToPt(textPaddingHorizontalMm) * scaleText;
        double padY = mmToPt(textPaddingVerticalMm) * scaleText;

        // === Excel 換算公式 ===
        int ColumnWidthToPx(int excelWidth) => (int)((excelWidth + 5) / 256.0 * 7);
        int RowHeightToPx(int excelHeight) => (int)(excelHeight / 15.0 * 20 * 96.0 / 72);
        double EmuToPt(double emu) => emu / 914400.0 * 72.0;

        // 預先算欄/列座標 (未縮放)
        int[] colX = new int[sheet.ColumnsWidth.Count + 1];
        colX[0] = 0;
        for (int i = 0; i < sheet.ColumnsWidth.Count; i++)
            colX[i + 1] = colX[i] + ColumnWidthToPx(sheet.ColumnsWidth[i]);

        int[] rowY = new int[sheet.Rows.Count + 1];
        rowY[0] = 0;
        for (int i = 0; i < sheet.Rows.Count; i++)
            rowY[i + 1] = rowY[i] + RowHeightToPx(sheet.Rows[i].Height);

        // === 1️⃣ 畫背景 + 文字 ===
        foreach (var cell in sheet.CellValues)
        {
            if (cell.Slave) continue;

            double x = colX[cell.ColStart] * scaleX;
            double y = rowY[cell.RowStart] * scaleY;

            int colEndSafe = Math.Min(cell.ColEnd + 1, colX.Length - 1);
            int rowEndSafe = Math.Min(cell.RowEnd + 1, rowY.Length - 1);

            double w = (colX[colEndSafe] - colX[cell.ColStart]) * scaleX;
            double h = (rowY[rowEndSafe] - rowY[cell.RowStart]) * scaleY;

            var style = sheet.MyCellStyles[cell.CellStyle_index];
            XRect rect = new XRect(x, y, w, h);

            // 背景
            if (style.FillForegroundColor != 0)
                g.DrawRectangle(new XSolidBrush(style.FillForegroundColor.ToXColor()), rect);

            // 文字
            string text = (cell.Text ?? "").Trim().TrimEnd('_');
            if (!string.IsNullOrEmpty(text))
            {
                XFontStyleEx fontStyle = XFontStyleEx.Regular;
                if (style.IsBold) fontStyle |= XFontStyleEx.Bold;
                if (style.IsItalic) fontStyle |= XFontStyleEx.Italic;
                if (style.IsStrikeout) fontStyle |= XFontStyleEx.Strikeout;

                double fontSize = (style.FontHeightInPoints > 0 ? style.FontHeightInPoints : 12) * scaleText;
                XFont font = new XFont(style.FontName ?? "微軟正黑體", fontSize, fontStyle);

                XStringFormat format = new XStringFormat();

                bool flag_Distributed = false;
                // --- 水平對齊 ---
                switch (style.Alignment)
                {
                    case HorizontalAlignment.Center:
                        format.Alignment = XStringAlignment.Center;
                        break;
                    case HorizontalAlignment.Right:
                        format.Alignment = XStringAlignment.Far;         
                        // ✅ 保留右邊留白
                        rect = new XRect(rect.X, rect.Y, rect.Width - padX, rect.Height);
                        break;
                    case HorizontalAlignment.Distributed: // ✅ 水平分散對齊
                        DrawDistributedTextHorizontal(g, text, font, XBrushes.Black, rect);
                        flag_Distributed = true;
                        break;
                    default: // Left
                        format.Alignment = XStringAlignment.Near;
                        // ✅ 保留左邊留白
                        rect = new XRect(rect.X + padX, rect.Y, rect.Width - padX, rect.Height);
                        break;
                }

                // --- 垂直對齊 ---
                switch (style.VerticalAlignment)
                {
                    case VerticalAlignment.Top:
                        format.LineAlignment = XLineAlignment.Near;
                        rect = new XRect(rect.X, rect.Y + padY, rect.Width, rect.Height - padY);
                        break;
                    case VerticalAlignment.Bottom:
                        format.LineAlignment = XLineAlignment.Far;
                        rect = new XRect(rect.X, rect.Y, rect.Width, rect.Height - padY);
                        break;
                    case VerticalAlignment.Distributed: // ✅ 垂直分散對齊
                        DrawDistributedTextVertical(g, text, font, XBrushes.Black, rect);
                        flag_Distributed = true;
                        break;
                    default: // Middle
                        format.LineAlignment = XLineAlignment.Center;        
                        break;
                }
                if(flag_Distributed == false)
                {
                    g.DrawString(text, font, XBrushes.Black, rect, format);
                }
         
            }
        }

        // === 2️⃣ 圖片 ===
        if (sheet.Pictures != null)
        {
            foreach (var pic in sheet.Pictures)
            {
                try
                {
                    byte[] imgBytes = Convert.FromBase64String(pic.Base64);
                    using (var ms = new MemoryStream(imgBytes))
                    using (var img = XImage.FromStream(ms))
                    {
                        double x1 = (colX[pic.ColStart] + EmuToPt(pic.Dx1)) * scaleX;
                        double y1 = (rowY[pic.RowStart] + EmuToPt(pic.Dy1)) * scaleY;
                        double x2 = (colX[pic.ColEnd] + EmuToPt(pic.Dx2)) * scaleX;
                        double y2 = (rowY[pic.RowEnd] + EmuToPt(pic.Dy2)) * scaleY;

                        double w = x2 - x1;
                        double h = y2 - y1;
                        if (w <= 0 || h <= 0) continue;

                        double scale = Math.Min(w / img.PixelWidth, h / img.PixelHeight);
                        double drawW = img.PixelWidth * scale;
                        double drawH = img.PixelHeight * scale;
                        double drawX = x1 + (w - drawW) / 2;
                        double drawY = y1 + (h - drawH) / 2;

                        g.DrawImage(img, drawX, drawY, drawW, drawH);
                    }
                }
                catch { }
            }
        }

        // === 3️⃣ 畫邊框 ===
        foreach (var cell in sheet.CellValues)
        {
            if (cell.Slave) continue;

            if (cell.RowEnd + 1 >= rowY.Length) continue;
            if (cell.ColEnd + 1 >= colX.Length) continue;

            double x1 = colX[cell.ColStart] * scaleX;
            double y1 = rowY[cell.RowStart] * scaleY;
            double x2 = colX[cell.ColEnd + 1] * scaleX;
            double y2 = rowY[cell.RowEnd + 1] * scaleY;

            var style = sheet.MyCellStyles[cell.CellStyle_index];

            double GetLineWidth(BorderStyle border) =>
                border == BorderStyle.Thick ? 2.5 :
                border == BorderStyle.Medium ? 1.5 :
                border == BorderStyle.Dashed ? 0.75 :
                border == BorderStyle.Hair ? 0.25 :
                border == BorderStyle.Thin ? 0.25 :
                (border == BorderStyle.None ? 0.0 : 0.5);

            if (style.BorderTop != BorderStyle.None)
                g.DrawLine(new XPen(XColors.Black, GetLineWidth(style.BorderTop)),
                           new XPoint(x1, y1), new XPoint(x2, y1));

            if (style.BorderBottom != BorderStyle.None)
                g.DrawLine(new XPen(XColors.Black, GetLineWidth(style.BorderBottom)),
                           new XPoint(x1, y2), new XPoint(x2, y2));

            if (style.BorderLeft != BorderStyle.None)
                g.DrawLine(new XPen(XColors.Black, GetLineWidth(style.BorderLeft)),
                           new XPoint(x1, y1), new XPoint(x1, y2));

            if (style.BorderRight != BorderStyle.None)
                g.DrawLine(new XPen(XColors.Black, GetLineWidth(style.BorderRight)),
                           new XPoint(x2, y1), new XPoint(x2, y2));
        }
    }
    /// <summary>
    /// 在矩形範圍內，繪製「水平分散對齊」文字
    /// </summary>
    public static void DrawDistributedTextHorizontal(XGraphics g, string text, XFont font, XBrush brush, XRect rect)
    {
        if (string.IsNullOrEmpty(text)) return;

        // 逐字量測寬度
        double totalTextWidth = 0;
        double[] charWidths = new double[text.Length];
        XSize size;
        for (int i = 0; i < text.Length; i++)
        {
            size = g.MeasureString(text[i].ToString(), font);
            charWidths[i] = size.Width;
            totalTextWidth += size.Width;
        }
        size = g.MeasureString("測".ToString(), font);
        totalTextWidth += size.Width;
        // 計算間距
        double spacing = 0;
        if (text.Length > 1)
            spacing = (rect.Width - totalTextWidth) / (text.Length - 1);

        // 計算基準 Y (垂直置中)
        double textHeight = g.MeasureString(text, font).Height;
        double curY = rect.Top + (rect.Height - textHeight) / 2;

        // 從左開始繪製
        double curX = rect.Left;
        curX += size.Width;
        for (int i = 0; i < text.Length; i++)
        {
            string text_ = text[i].ToString();
            g.DrawString(text_.ToString(), font, brush, new XPoint(curX, curY + textHeight / 2), XStringFormats.Center);
            curX += charWidths[i] + spacing;
        }
    }
    /// <summary>
    /// 在矩形範圍內，繪製「垂直分散對齊」文字（多行平均分布）
    /// </summary>
    public static void DrawDistributedTextVertical(XGraphics g, string text, XFont font, XBrush brush, XRect rect)
    {
        if (string.IsNullOrEmpty(text)) return;

        // 切行 (Excel 通常以換行符號分行)
        string[] lines = text.Split(new[] { '\n' }, StringSplitOptions.None);

        if (lines.Length == 1)
        {
            // 單行 → 視為置中
            g.DrawString(text, font, brush, rect, XStringFormats.Center);
            return;
        }

        // 計算行高
        double lineHeight = g.MeasureString("測", font).Height; // 用字體實際高度
        double spacing = (rect.Height - lineHeight) / (lines.Length - 1);

        // 基準 X (水平置中)
        double baseX = rect.Left + rect.Width / 2;

        // 繪製每行
        for (int i = 0; i < lines.Length; i++)
        {
            double y = rect.Top + i * spacing + lineHeight / 2;
            g.DrawString(lines[i], font, brush,
                new XPoint(baseX, y), XStringFormats.Center);
        }
    }



    // 顏色轉換 NPOI_Color → XColor
    public static XColor ToXColor(this NPOI_Color color)
    {
        var sysColor = color.ToColor(); // 這裡假設你已有 ToColor() -> System.Drawing.Color
        return XColor.FromArgb(sysColor.A, sysColor.R, sysColor.G, sysColor.B);
    }
    public static XColor ToXColor(this short colorIndex)
    {
        switch (colorIndex)
        {
            case 0: return XColors.Black;
            case 1: return XColors.White;
            case 2: return XColors.Red;
            case 3: return XColors.Green;
            case 4: return XColors.Blue;
            // 你可以依 NPOI 顏色表擴充
            default: return XColors.White;
        }
    }
    public class CustomFontResolver : IFontResolver
    {
        private readonly Dictionary<string, string> _fontMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // ✅ 修改成你拆出來的 TTF 檔案路徑
            { "標楷體", @"C:\Windows\Fonts\kaiu.ttf" },
            { "新細明體", @"C:\Windows\Fonts\ttf_export\mingliu_0.ttf" },
            { "微軟正黑體", @"C:\Windows\Fonts\ttf_export\msjh_0.ttf" }
        };

        public string DefaultFontName => "標楷體"; // 預設字型

        public byte[] GetFont(string faceName)
        {
            if (_fontMapping.TryGetValue(faceName, out string path) && File.Exists(path))
            {
                return File.ReadAllBytes(path);
            }

            Console.WriteLine($"⚠ 找不到字型: {faceName}，改用 {DefaultFontName}");
            return File.ReadAllBytes(_fontMapping[DefaultFontName]);
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (_fontMapping.ContainsKey(familyName))
            {
                return new FontResolverInfo(familyName);
            }

            // fallback
            return new FontResolverInfo(DefaultFontName);
        }
    }
}

