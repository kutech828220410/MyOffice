﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.IO;
using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using Basic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.ComponentModel;

namespace MyOffice
{
    public enum NPOI_Color
    {
        Basic = 0,
        BLACK = 8,
        BROWN = 60,
        OLIVE_GREEN = 59,
        DARK_GREEN = 58,
        DARK_TEAL = 56,
        DARK_BLUE = 18,
        INDIGO = 62,
        GREY_80_PERCENT = 63,
        DARK_RED = 16,
        ORANGE = 53,
        DARK_YELLOW = 19,
        GREEN = 17,
        TEAL = 21,
        BLUE = 12,
        BLUE_GREY = 54,
        GREY_50_PERCENT = 23,
        RED = 10,
        LIGHT_ORANGE = 52,
        LIME = 50,
        SEA_GREEN = 57,
        AQUA = 49,
        LIGHT_BLUE = 48,
        VIOLET = 20,
        GREY_40_PERCENT = 55,
        PINK = 14,
        GOLD = 51,
        YELLOW = 13,
        BRIGHT_GREEN = 11,
        TURQUOISE = 15,
        SKY_BLUE = 40,
        PLUM = 61,
        GREY_25_PERCENT = 22,
        ROSE = 45,
        TAN = 47,
        LIGHT_YELLOW = 43,
        LIGHT_GREEN = 42,
        LIGHT_TURQUOISE = 41,
        PALE_BLUE = 44,
        LAVENDER = 46,
        WHITE = 9,
        CORNFLOWER_BLUE = 24,
        LEMON_CHIFFON = 26,
        MAROON = 25,
        ORCHID = 28,
        CORAL = 29,
        ROYAL_BLUE = 30,
        LIGHT_CORNFLOWER_BLUE = 31,
        AUTOMATIC = 64,

        BLACK2 = 32767,
    }
    public enum H_Alignment
    {
        Left,
        Center,
        Right,
    }
    public enum V_Alignment
    {
        Top,
        Center,
        Bottom,
    }
    public enum Excel_Type
    {
        xls,
        xlsx,
    }
    [Serializable]
    public class SheetClass
    {
        public SheetClass()
        {

        }
        public SheetClass(string Name)
        {
            this.Name = Name;
        }
        public class Row
        {
            public int Count
            {
                get
                {
                    return cell.Count;
                }
            }
            public int Height
            {
                get
                {
                    int temp = 0;
                    for (int i = 0; i < Cell.Count; i++)
                    {
                        if (Cell[i].Height > temp) temp = Cell[i].Height;
                    }
                    return temp / 20;
                }
                set
                {
                    for (int i = 0; i < Cell.Count; i++)
                    {
                        Cell[i].Height = (short)value;
                    }
                }
            }
            private List<CellValue> cell = new List<CellValue>();
            public List<CellValue> Cell { get => cell; set => cell = value; }
        }
        public List<ICellStyle> cellStyles = new List<ICellStyle>();
        public List<Row> Rows
        {
            get
            {
                List<Row> rows = new List<Row>();
                List<CellValue> cellValues_dist = cellValues.Distinct(new Distinct_CellValueRowStart()).ToList();
                cellValues_dist.Sort(new Icp_CellValue());
                for (int i = 0; i < cellValues_dist.Count; i++)
                {
                    int row_index = cellValues_dist[i].RowStart;
                    List<CellValue> cellValues_buf = new List<CellValue>();
                    cellValues_buf = (from value in cellValues
                                      where value.RowStart == row_index
                                      select value).ToList();
                    Row row = new Row();
                    row.Cell.LockAdd(cellValues_buf);
                    rows.Add(row);
                }
                return rows;
            }
        }
        public int Height
        {
            get
            {
                int temp = 0;
                for (int i = 0; i < Rows.Count; i++)
                {
                    temp += (int)(Rows[i].Height);
                }
                return temp ;
            }
        }
        public int Width
        {
            get
            {
                int temp = 0;
                for (int i = 0; i < columnsWidth.Count; i++)
                {
                    temp += (int)(columnsWidth[i] / 256 + 0.71);
                }
                return temp;
            }
        }
        public double Scale_X = 7.5;
        public double Scale_Y = 4D / 3D;

        private string name = "";
        private List<CellValue> cellValues = new List<CellValue>();
        private List<MyCellStyle> myCellStyles = new List<MyCellStyle>();

        private List<CellValue> cellValues_buffer = new List<CellValue>();
        private List<MyCellStyle> myCellStyles_buffer = new List<MyCellStyle>();

        private List<int> columnsWidth = new List<int>();

        public List<CellValue> CellValues { get => cellValues; set => cellValues = value; }
        public List<MyCellStyle> MyCellStyles { get => myCellStyles; set => myCellStyles = value; }
        public List<int> ColumnsWidth { get => columnsWidth; set => columnsWidth = value; }
        public string Name { get => name; set => name = value; }

        public int GetWidth(int index)
        {
            return (int)(columnsWidth[index] / 256 + 0.71);
        }
        public int GetHeight(int index_start, int index_end)
        {
            if (index_start > Rows.Count) return 0;
            if (index_end > Rows.Count) return 0;
            int temp = 0;
            for (int i = index_start; i < index_end; i++)
            {
                temp += (int)(Rows[i].Height);
            }
            return temp;
        }

        public CellValue SortCellValue(int RowStart, int RowEnd ,int ColStart,int ColEnd)
        {
            lock(CellValues)
            {
                List<CellValue> CellValues_buf = new List<CellValue>();
                CellValues_buf = (from value in CellValues
                                  where value.RowStart == RowStart
                                  where value.RowEnd == RowEnd
                                  where value.ColStart == ColStart
                                  where value.ColEnd == ColEnd
                                  select value).ToList();
                if (CellValues_buf.Count == 0) return null;
                return CellValues_buf[0];
            }
          
        }
        public bool Add(CellValue cellValue, MyCellStyle myCellStyle)
        {
            lock(CellValues)
            {
             
                List<CellValue> CellValues_buf = new List<CellValue>();
                CellValues_buf = (from value in CellValues
                                  where value.RowStart == cellValue.RowStart
                                  where value.RowEnd == cellValue.RowEnd
                                  where value.ColStart == cellValue.ColStart
                                  where value.ColEnd == cellValue.ColEnd
                                  select value).ToList();
                if (CellValues_buf.Count == 0)
                {
                    int index = this.Add(myCellStyle);
                    cellValue.CellStyle_index = index;
                    this.CellValues.Add(cellValue);
                    return true;
                }
                return false;
            }        
        }
        public int Add(MyCellStyle myCellStyle)
        {
            for (int i = 0; i < MyCellStyles.Count; i++)
            {
                bool flag_ok = true;
                if (myCellStyle.VerticalAlignment != MyCellStyles[i].VerticalAlignment) flag_ok = false;
                if (myCellStyle.Alignment != MyCellStyles[i].Alignment) flag_ok = false;
                if (myCellStyle.BorderTop != MyCellStyles[i].BorderTop) flag_ok = false;
                if (myCellStyle.BorderBottom != MyCellStyles[i].BorderBottom) flag_ok = false;
                if (myCellStyle.BorderLeft != MyCellStyles[i].BorderLeft) flag_ok = false;
                if (myCellStyle.BorderRight != MyCellStyles[i].BorderRight) flag_ok = false;
                if (myCellStyle.FillForegroundColor != MyCellStyles[i].FillForegroundColor) flag_ok = false;


                if (myCellStyle.FontName != MyCellStyles[i].FontName) flag_ok = false;
                if (myCellStyle.FontHeight != MyCellStyles[i].FontHeight) flag_ok = false;
                if (myCellStyle.FontHeightInPoints != MyCellStyles[i].FontHeightInPoints) flag_ok = false;
                if (myCellStyle.IsItalic != MyCellStyles[i].IsItalic) flag_ok = false;
                if (myCellStyle.IsStrikeout != MyCellStyles[i].IsStrikeout) flag_ok = false;
                if (myCellStyle.Color != MyCellStyles[i].Color) flag_ok = false;
                if (myCellStyle.TypeOffset != MyCellStyles[i].TypeOffset) flag_ok = false;
                if (myCellStyle.Underline != MyCellStyles[i].Underline) flag_ok = false;
                if (myCellStyle.Charset != MyCellStyles[i].Charset) flag_ok = false;
                if (myCellStyle.Index != MyCellStyles[i].Index) flag_ok = false;
                if (myCellStyle.Boldweight != MyCellStyles[i].Boldweight) flag_ok = false;
                if (myCellStyle.IsBold != MyCellStyles[i].IsBold) flag_ok = false;
                if(flag_ok)
                {
                    return i;
                }
            }
            MyCellStyles.Add(myCellStyle);
            return MyCellStyles.Count - 1;
        }
        public void Init(NPOI.SS.UserModel.IWorkbook workbook)
        {
            for (int i = 0; i < MyCellStyles.Count; i++)
            {
                ICellStyle cellStyle = workbook.CreateCellStyle();

                cellStyle.Alignment = MyCellStyles[i].Alignment;
                cellStyle.VerticalAlignment = MyCellStyles[i].VerticalAlignment;
                cellStyle.BorderTop = MyCellStyles[i].BorderTop;
                cellStyle.BorderBottom = MyCellStyles[i].BorderBottom;
                cellStyle.BorderLeft = MyCellStyles[i].BorderLeft;
                cellStyle.BorderRight = MyCellStyles[i].BorderRight;
                cellStyle.WrapText = true;
                IFont myFont = workbook.CreateFont();
                myFont.FontName = MyCellStyles[i].FontName;
                myFont.FontHeight = MyCellStyles[i].FontHeight;
                myFont.FontHeightInPoints = MyCellStyles[i].FontHeightInPoints;
                myFont.IsItalic = MyCellStyles[i].IsItalic;
                myFont.IsStrikeout = MyCellStyles[i].IsStrikeout;
                myFont.Color = MyCellStyles[i].Color;
                myFont.TypeOffset = MyCellStyles[i].TypeOffset;
                myFont.Underline = MyCellStyles[i].Underline;
                myFont.Charset = MyCellStyles[i].Charset;
                myFont.Boldweight = MyCellStyles[i].Boldweight;
                myFont.IsBold = MyCellStyles[i].IsBold;

                cellStyle.SetFont(myFont);
                cellStyles.Add(cellStyle);
            }
            for (int i = 0; i < CellValues.Count; i++)
            {
                short fontHeight = (short)this.GetICellStyle(CellValues[i].CellStyle_index).GetFont(workbook).FontHeight;
                if (CellValues[i].Height < fontHeight) CellValues[i].Height = fontHeight;
                
            }
        }
        public ICellStyle GetICellStyle(int index)
        {
            return cellStyles[index];
        }


        public Rectangle GetCellSize(CellValue cellValue)
        {
            return this.GetCellSize(cellValue, 0);
        }
        public Rectangle GetCellSize(CellValue cellValue , int rowIndex_start)
        {
            Rectangle rect = new Rectangle();

            int col_len = cellValue.ColEnd - cellValue.ColStart;
            int row_len = cellValue.RowEnd - cellValue.RowStart;
            col_len++;
            row_len++;
            for (int i = 0; i < cellValue.ColStart; i++)
            {
                rect.X += GetWidth(i);
            }
            for (int i = rowIndex_start; i < cellValue.RowStart; i++)
            {
                rect.Y += Rows[i].Height;
            }
            for (int i = cellValue.ColStart; i <= cellValue.ColEnd; i++)
            {
                rect.Width += GetWidth(i);
            }
            for (int i = cellValue.RowStart; i <= cellValue.RowEnd; i++)
            {
                rect.Height += Rows[i].Height;
            }
            return rect;
        }
        public Size ToPixcel(Size size)
        {
            size.Width = (int)(size.Width * Scale_X);
            size.Height = (int)(size.Height * Scale_Y);
            return size;
        }
        public SizeF ToPixcel(SizeF size)
        {
            size.Width = (float)(size.Width * Scale_X);
            size.Height = (float)(size.Height * Scale_Y);
            return size;
        }
        public PointF ToPixcel(PointF pointF)
        {
            pointF.X = (int)(pointF.X * Scale_X);
            pointF.Y = (int)(pointF.Y * Scale_Y);
            return pointF;
        }
        public Rectangle ToPixcel(Rectangle rect)
        {
            rect.X = (int)(rect.X * Scale_X);
            rect.Y = (int)(rect.Y * Scale_Y);
            rect.Width = (int)(rect.Width * Scale_X);
            rect.Height = (int)(rect.Height * Scale_Y);

            return rect;
        }
        public Bitmap GetBitmap()
        {
            Rectangle rectangle = new Rectangle();
            return this.GetBitmap(0, Rows.Count, ref rectangle);
        }
        public Bitmap GetBitmap(int width, int height, double Scale , H_Alignment horizontalAlignment, V_Alignment verticalAlignment, int pad_left = 0, int pad_top = 0)
        {
            Rectangle rectangle = new Rectangle();
            Bitmap bitmap = new Bitmap(width, height);
            using (Bitmap bitmap_sheet = this.GetBitmap(0, Rows.Count, ref rectangle))
            {
                Graphics g = Graphics.FromImage(bitmap);
                rectangle.Width = (int)(rectangle.Width * Scale);
                rectangle.Height = (int)(rectangle.Height * Scale);

                if(horizontalAlignment == H_Alignment.Left)
                {
                    rectangle.X = pad_left;
                }
                if (horizontalAlignment == H_Alignment.Center)
                {
                    rectangle.X = (width - rectangle.Width) / 2;
                }
                if (horizontalAlignment == H_Alignment.Right)
                {
                    rectangle.X = (width - rectangle.Width);
                }

                if(verticalAlignment == V_Alignment.Top)
                {
                    rectangle.Y = pad_top;
                }
                if (verticalAlignment == V_Alignment.Center)
                {
                    rectangle.Y = (height - rectangle.Height) / 2;
                }
                if (verticalAlignment == V_Alignment.Bottom)
                {
                    rectangle.Y = (height - rectangle.Height);
                }
                g.DrawImage(bitmap_sheet, rectangle);
                g.Dispose();
            }
            return bitmap;
        }
        public Bitmap GetBitmap(int RowIndex, ref Rectangle rectangle_Area)
        {
            return this.GetBitmap(RowIndex, RowIndex + 1, ref rectangle_Area);
        }
        public Bitmap GetBitmap(int RowIndex)
        {
            Rectangle rectangle = new Rectangle();
            return this.GetBitmap(RowIndex, RowIndex + 1 , ref rectangle);
        }
        public Bitmap GetBitmap(int Rowindex_start, int Rowindex_end , ref Rectangle rectangle_Area)
        {
            int width = this.Width;
            int height = GetHeight(Rowindex_start, Rowindex_end);
            Size size = ToPixcel(new Size(width, height));
            Bitmap bitmap = new Bitmap(size.Width + 1, size.Height + 1);
            Graphics g = Graphics.FromImage(bitmap);
            g.SmoothingMode = SmoothingMode.HighQuality; //使繪圖質量最高，即消除鋸齒
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

            if (Rowindex_start > Rows.Count) return null;
            if (Rowindex_end > Rows.Count) return null;
            for (int r = Rowindex_start; r < Rowindex_end; r++)
            {           
                for (int i = 0; i < Rows[r].Cell.Count; i++)
                {
                    if (r == Rowindex_start && i == 0)
                    {
                        rectangle_Area = GetCellSize(Rows[r].Cell[i]);
                        rectangle_Area.Width = this.Width + 1;
                        rectangle_Area.Height++;
                        rectangle_Area = ToPixcel(rectangle_Area);
                     
                    }
                    if (Rows[r].Cell[i].Slave == true) continue;
                    Rectangle rect = GetCellSize(Rows[r].Cell[i] , Rowindex_start);
                    rect = ToPixcel(rect);
                    MyCellStyle myCellStyle = MyCellStyles[Rows[r].Cell[i].CellStyle_index];

                    this.DrawBorder(g, rect, myCellStyle);

                    FontStyle fontStyle = new FontStyle();
                    if (myCellStyle.IsBold) fontStyle |= FontStyle.Bold;
                    if (myCellStyle.IsItalic) fontStyle |= FontStyle.Italic;
                    if (myCellStyle.IsStrikeout) fontStyle |= FontStyle.Strikeout;
                    if (myCellStyle.Underline == FontUnderlineType.Single) fontStyle |= FontStyle.Underline;
                    Font font = new Font(myCellStyle.FontName, (float)myCellStyle.FontHeightInPoints, fontStyle);
                    Color fore_color = ((NPOI_Color)myCellStyle.Color).ToColor();
                    Color background_color = ((NPOI_Color)myCellStyle.FillForegroundColor).ToColor();
                    if (fore_color == Color.White && background_color == Color.White)
                    {
                        fore_color = Color.Black;
                    }
                    g.FillRectangle(new SolidBrush(background_color), rect);
                    SizeF sizeF_font = g.MeasureString(Rows[r].Cell[i].Text, font, new Size(rect.Width, rect.Height), StringFormat.GenericDefault);
                    //sizeF_font = ToPixcel(sizeF_font);
                    PointF pointF_font = new PointF();

                    if (myCellStyle.Alignment == HorizontalAlignment.Center || myCellStyle.Alignment == HorizontalAlignment.General)
                    {
                        pointF_font.X = ((rect.Width - sizeF_font.Width) / 2) + rect.X;
                    }
                    else if (myCellStyle.Alignment == HorizontalAlignment.Left)
                    {
                        pointF_font.X = rect.X;
                    }
                    else if (myCellStyle.Alignment == HorizontalAlignment.Right)
                    {
                        pointF_font.X = ((rect.Width - sizeF_font.Width)) + rect.X;
                    }

                    if (myCellStyle.VerticalAlignment == VerticalAlignment.Center || myCellStyle.VerticalAlignment == VerticalAlignment.None)
                    {
                        pointF_font.Y = ((rect.Height - sizeF_font.Height) / 2) + rect.Y;
                    }
                    else if (myCellStyle.VerticalAlignment == VerticalAlignment.Top)
                    {
                        pointF_font.Y = rect.Y;
                    }
                    else if (myCellStyle.VerticalAlignment == VerticalAlignment.Bottom)
                    {
                        pointF_font.Y = ((rect.Height - sizeF_font.Height)) + rect.Y;
                    }

                    g.DrawString(Rows[r].Cell[i].Text, font, new SolidBrush(fore_color), new RectangleF(pointF_font.X, pointF_font.Y, rect.Width, rect.Height), StringFormat.GenericDefault);

                    this.DrawBorder(g, rect, myCellStyle);
                }
            }

            g.Dispose();
            rectangle_Area.Width = bitmap.Width;
            rectangle_Area.Height = bitmap.Height;
            return bitmap;
        }
        public void AddNewCell(int row, int col, string text, Font font, int height, BorderStyle BS_top, BorderStyle BS_bottom, BorderStyle BS_left, BorderStyle BS_right)
        {
            this.AddNewCell(row, row, col, col, text, font, NPOI_Color.BLACK, height,HorizontalAlignment.Center,VerticalAlignment.Center, BS_top, BS_bottom, BS_left, BS_right);
        }
        public void AddNewCell(int row , int col , string text, Font font)
        {
            this.AddNewCell(row, row, col, col, text, font, NPOI_Color.BLACK);
        }
        public void AddNewCell(int row, int col, string text, Font font, int height)
        {
            this.AddNewCell(row, row, col, col, text, font, NPOI_Color.BLACK, height);
        }
        public void AddNewCell(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, Font font)
        {
            this.AddNewCell(RowStart, RowEnd, ColStart, ColEnd, text, font, NPOI_Color.BLACK);
        }
        public void AddNewCell(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, Font font,int height)
        {
            this.AddNewCell(RowStart, RowEnd, ColStart, ColEnd, text, font, NPOI_Color.BLACK, height);
        }
        public void AddNewCell(int row, int col, string text, Font font, NPOI_Color foreColor, int height)
        {
            this.AddNewCell(row, row, col, col, text, font, foreColor, height);
        }
        public void AddNewCell(int row, int col, string text, Font font, NPOI_Color foreColor)
        {
            this.AddNewCell(row, row, col, col, text, font, foreColor);
        }    
        public void AddNewCell(int row, int col, string text, Font font, NPOI_Color foreColor, HorizontalAlignment horizontalAlignment)
        {
            this.AddNewCell(row, row, col, col, text, font, foreColor, 0, horizontalAlignment, VerticalAlignment.Center, BorderStyle.None, BorderStyle.None, BorderStyle.None, BorderStyle.None);
        }
        public void AddNewCell(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, Font font, NPOI_Color foreColor, HorizontalAlignment horizontalAlignment)
        {
            this.AddNewCell(RowStart, RowEnd, ColStart, ColEnd, text, font, foreColor, 0, horizontalAlignment,  VerticalAlignment.Center, BorderStyle.None, BorderStyle.None, BorderStyle.None, BorderStyle.None);
        }
        public void AddNewCell(int row, int col, string text, Font font, NPOI_Color foreColor, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            this.AddNewCell(row, row, col, col, text, font, foreColor, 0, horizontalAlignment, verticalAlignment, BorderStyle.None, BorderStyle.None, BorderStyle.None, BorderStyle.None);
        }
        public void AddNewCell(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, Font font, NPOI_Color foreColor, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            this.AddNewCell(RowStart, RowEnd, ColStart, ColEnd, text, font, foreColor, 0, horizontalAlignment, verticalAlignment, BorderStyle.None, BorderStyle.None, BorderStyle.None, BorderStyle.None);
        }
        public void AddNewCell(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, Font font, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS_top = BorderStyle.Thin, BorderStyle BS_bottom = BorderStyle.Thin, BorderStyle BS_left = BorderStyle.Thin, BorderStyle BS_right = BorderStyle.Thin, bool flag_Slave = false)
        {
            CellValue cellValue = new CellValue();
            MyCellStyle myCellStyle = new MyCellStyle();
            cellValue.RowStart = RowStart;
            cellValue.RowEnd = RowEnd;
            cellValue.ColStart = ColStart;
            cellValue.ColEnd = ColEnd;
            cellValue.Text = text;
            cellValue.Height = (short)height;
            cellValue.Slave = flag_Slave;

            myCellStyle.FontName = font.Name;
            myCellStyle.FontHeight = font.Height;
            myCellStyle.FontHeightInPoints = font.Size;
            myCellStyle.IsItalic = font.Italic;
            myCellStyle.IsStrikeout = font.Strikeout;
            myCellStyle.Color = (short)(foreColor);
            myCellStyle.Charset = font.GdiCharSet;
            myCellStyle.IsBold = font.Bold;
            myCellStyle.Alignment = horizontalAlignment;
            myCellStyle.VerticalAlignment = verticalAlignment;

            myCellStyle.BorderTop = BS_top;
            myCellStyle.BorderBottom = BS_bottom;
            myCellStyle.BorderLeft = BS_left;
            myCellStyle.BorderRight = BS_right;

            this.Add(cellValue, myCellStyle);
        }

        public void AddNewCell_Webapi(int Row, int Col, string text, string FontName, float FontHeightInPoints, bool IsBold, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS = BorderStyle.Thin)
        {
            this.AddNewCell_Webapi(Row, Row, Col, Col, text, FontName, FontHeightInPoints, IsBold, foreColor, height, horizontalAlignment, verticalAlignment, BS, BS, BS, BS);
        }
        public CellValue AddNewCell_Webapi(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, string FontName, float FontHeightInPoints, bool IsBold, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS = BorderStyle.Thin)
        {
           return this.AddNewCell_Webapi(RowStart, RowEnd, ColStart, ColEnd, text, FontName, FontHeightInPoints, IsBold, foreColor, height, horizontalAlignment, verticalAlignment, BS, BS, BS, BS);
        }

        public CellValue AddNewCell_Webapi(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, string FontName , float FontHeightInPoints , bool IsBold, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS_top = BorderStyle.Thin, BorderStyle BS_bottom = BorderStyle.Thin, BorderStyle BS_left = BorderStyle.Thin, BorderStyle BS_right = BorderStyle.Thin)
        {
            CellValue cellValue = new CellValue();
            MyCellStyle myCellStyle = new MyCellStyle();
            cellValue.RowStart = RowStart;
            cellValue.RowEnd = RowEnd;
            cellValue.ColStart = ColStart;
            cellValue.ColEnd = ColEnd;
            cellValue.Text = text;
            cellValue.Height = (short)height;

            myCellStyle.FontName = FontName;
            myCellStyle.FontHeight = FontHeightInPoints;
            myCellStyle.FontHeightInPoints = FontHeightInPoints;

            myCellStyle.Color = (short)(foreColor);
            myCellStyle.IsBold = IsBold;
            myCellStyle.Alignment = horizontalAlignment;
            myCellStyle.VerticalAlignment = verticalAlignment;

            myCellStyle.BorderTop = BS_top;
            myCellStyle.BorderBottom = BS_bottom;
            myCellStyle.BorderLeft = BS_left;
            myCellStyle.BorderRight = BS_right;

            this.Add(cellValue, myCellStyle);
            return cellValue;
        }
        public void NewCell_Webapi_Buffer_Clear()
        {
            cellValues_buffer.Clear();
            myCellStyles_buffer.Clear();
        }
        public void AddNewCell_Webapi_Buffer(int Row, int Col, string text, string FontName, float FontHeightInPoints, bool IsBold, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS = BorderStyle.Thin)
        {
            this.AddNewCell_Webapi_Buffer(Row, Row, Col, Col, text, FontName, FontHeightInPoints, IsBold, foreColor, height, horizontalAlignment, verticalAlignment, BS, BS, BS, BS);
        }
        public void AddNewCell_Webapi_Buffer(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, string FontName, float FontHeightInPoints, bool IsBold, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS = BorderStyle.Thin)
        {
            this.AddNewCell_Webapi_Buffer(RowStart, RowEnd, ColStart, ColEnd, text, FontName, FontHeightInPoints, IsBold, foreColor, height, horizontalAlignment, verticalAlignment, BS, BS, BS, BS);
        }

        public void AddNewCell_Webapi_Buffer(int RowStart, int RowEnd, int ColStart, int ColEnd, string text, string FontName, float FontHeightInPoints, bool IsBold, NPOI_Color foreColor, int height = 0, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center, BorderStyle BS_top = BorderStyle.Thin, BorderStyle BS_bottom = BorderStyle.Thin, BorderStyle BS_left = BorderStyle.Thin, BorderStyle BS_right = BorderStyle.Thin)
        {
            CellValue cellValue = new CellValue();
            MyCellStyle myCellStyle = new MyCellStyle();
            cellValue.RowStart = RowStart;
            cellValue.RowEnd = RowEnd;
            cellValue.ColStart = ColStart;
            cellValue.ColEnd = ColEnd;
            cellValue.Text = text;
            cellValue.Height = (short)height;

            myCellStyle.FontName = FontName;
            myCellStyle.FontHeight = FontHeightInPoints;
            myCellStyle.FontHeightInPoints = FontHeightInPoints;

            myCellStyle.Color = (short)(foreColor);
            myCellStyle.IsBold = IsBold;
            myCellStyle.Alignment = horizontalAlignment;
            myCellStyle.VerticalAlignment = verticalAlignment;

            myCellStyle.BorderTop = BS_top;
            myCellStyle.BorderBottom = BS_bottom;
            myCellStyle.BorderLeft = BS_left;
            myCellStyle.BorderRight = BS_right;

            int index = this.Add(myCellStyle);
            cellValue.CellStyle_index = index;
            cellValues_buffer.Add(cellValue);
         
        }
        public void NewCell_Webapi_Buffer_Caculate()
        {
            cellValues_buffer = cellValues_buffer.Distinct(new Distinct_CellValue()).ToList();
            for (int i = 0; i < cellValues_buffer.Count; i++)
            {
                cellValues.Add(cellValues_buffer[i]);
            }
        }
        public void SetSlave(int Row, int Col, bool flag_Slave)
        {
            this.Rows[Row].Cell[Col].Slave = flag_Slave;
        }
        public void ReplaceCell(int Row, int Col, string Text)
        {
            this.Rows[Row].Cell[Col].Text = Text;
        }

        private void DrawBorder(Graphics g, RectangleF rectangleF, MyCellStyle myCellStyle )
        {
            PointF pointF_TOP_Start = new PointF(rectangleF.X, rectangleF.Y);
            PointF pointF_TOP_End = new PointF(rectangleF.X + rectangleF.Width, rectangleF.Y);

            PointF pointF_Bottom_Start = new PointF(rectangleF.X, rectangleF.Y + rectangleF.Height);
            PointF pointF_Bottom_End = new PointF(rectangleF.X + rectangleF.Width, rectangleF.Y + rectangleF.Height);

            PointF pointF_Left_Start = new PointF(rectangleF.X, rectangleF.Y);
            PointF pointF_Left_End = new PointF(rectangleF.X , rectangleF.Y + rectangleF.Height);

            PointF pointF_Right_Start = new PointF(rectangleF.X + rectangleF.Width, rectangleF.Y);
            PointF pointF_Right_End = new PointF(rectangleF.X + rectangleF.Width, rectangleF.Y + rectangleF.Height);


            this.DrawBorder(g, pointF_TOP_Start, pointF_TOP_End, myCellStyle.BorderTop, Color.Black);
            this.DrawBorder(g, pointF_Bottom_Start, pointF_Bottom_End, myCellStyle.BorderBottom, Color.Black);
            this.DrawBorder(g, pointF_Left_Start, pointF_Left_End, myCellStyle.BorderLeft, Color.Black);
            this.DrawBorder(g, pointF_Right_Start, pointF_Right_End, myCellStyle.BorderRight, Color.Black);

        }
        private void DrawBorder(Graphics g, PointF pointF_start, PointF pointF_end, BorderStyle borderStyle , Color color)
        {
            if (borderStyle == BorderStyle.Thin)
            {
                Pen _pen = new Pen(color, 1);
                g.DrawLine(_pen, pointF_start, pointF_end);
            }
            if (borderStyle == BorderStyle.Medium)
            {
                Pen _pen = new Pen(color, 2);
                g.DrawLine(_pen, pointF_start, pointF_end);
            }
            else if (borderStyle == BorderStyle.Thick)
            {
                Pen _pen = new Pen(color, 3);
                g.DrawLine(_pen, pointF_start, pointF_end);
            }
        }
        private class Icp_CellValue : IComparer<CellValue>
        {
            public int Compare(CellValue x, CellValue y)
            {
                return x.RowStart.CompareTo(y.RowStart);
            }
        }
        private class Distinct_CellValueRowStart : IEqualityComparer<CellValue>
        {
            public bool Equals(CellValue x, CellValue y)
            {
                return (x.RowStart == y.RowStart);
            }

            public int GetHashCode(CellValue obj)
            {
                return 1;
            }
        }
        private class Distinct_CellValue : IEqualityComparer<CellValue>
        {
            public bool Equals(CellValue x, CellValue y)
            {
                return (x.RowStart == y.RowStart) && (x.RowEnd == y.RowEnd) && (x.ColStart == y.ColStart) && (x.ColEnd == y.ColEnd);
            }

            public int GetHashCode(CellValue obj)
            {
                return 1;
            }
        }
        public static Bitmap ScaleImage(Bitmap pBmp, int pWidth, int pHeight)
        {
            try
            {
                Bitmap tmpBmp = new Bitmap(pWidth, pHeight);
                Graphics tmpG = Graphics.FromImage(tmpBmp);

                //tmpG.InterpolationMode = InterpolationMode.HighQualityBicubic;

                tmpG.DrawImage(pBmp,
                                           new Rectangle(0, 0, pWidth, pHeight),
                                           new Rectangle(0, 0, pBmp.Width, pBmp.Height),
                                           GraphicsUnit.Pixel);
                tmpG.Dispose();
                return tmpBmp;
            }
            catch
            {
                return null;
            }
        }

     

    }
    [Serializable]
    public class CellValue
    {
        static public class ColorSerializationHelper
        {
            static public Color FromString(string value)
            {
                var parts = value.Split(':');

                int A = 0;
                int R = 0;
                int G = 0;
                int B = 0;
                int.TryParse(parts[0], out A);
                int.TryParse(parts[1], out R);
                int.TryParse(parts[2], out G);
                int.TryParse(parts[3], out B);
                return Color.FromArgb(A, R, G, B);
            }
            static public string ToString(Color color)
            {
                return color.A + ":" + color.R + ":" + color.G + ":" + color.B;

            }
        }
        [TypeConverter(typeof(FontConverter))]
        static public class FontSerializationHelper
        {
            static public Font FromString(string value)
            {
                var parts = value.Split(':');
                return new Font(
                    parts[0],                                                   // FontFamily.Name
                    float.Parse(parts[1]),                                      // Size
                    EnumSerializationHelper.FromString<FontStyle>(parts[2]),    // Style
                    EnumSerializationHelper.FromString<GraphicsUnit>(parts[3]), // Unit
                    byte.Parse(parts[4]),                                       // GdiCharSet
                    bool.Parse(parts[5])                                        // GdiVerticalFont
                );
            }
            static public string ToString(Font font)
            {
                return font.FontFamily.Name
                        + ":" + font.Size
                        + ":" + font.Style
                        + ":" + font.Unit
                        + ":" + font.GdiCharSet
                        + ":" + font.GdiVerticalFont
                        ;
            }
        }
        [TypeConverter(typeof(EnumConverter))]
        static public class EnumSerializationHelper
        {
            static public T FromString<T>(string value)
            {
                return (T)Enum.Parse(typeof(T), value, true);
            }
        }
        

        private string text = "";
        private int rowStart = 0;
        private int rowEnd = 0;
        private int colStart = 0;
        private int colEnd = 0;
        private bool slave = false;
        private short height = 0;
        private bool isDouble = false;

        private int cellStyle_index;
        public string Text { get => text; set => text = value; }
        public int RowStart { get => rowStart; set => rowStart = value; }
        public int RowEnd { get => rowEnd; set => rowEnd = value; }
        public int ColStart { get => colStart; set => colStart = value; }
        public int ColEnd { get => colEnd; set => colEnd = value; }
        public int CellStyle_index { get => cellStyle_index; set => cellStyle_index = value; }
        public bool Slave { get => slave; set => slave = value; }
        public short Height { get => height; set => height = value; }
        public bool IsDouble { get => isDouble; set => isDouble = value; }
    }
   
    [Serializable]
    public class MyCellStyle
    {
        public VerticalAlignment VerticalAlignment { get; set; }
        public HorizontalAlignment Alignment { get; set; }
        public BorderStyle BorderBottom { get; set; }
        public BorderStyle BorderTop { get; set; }
        public BorderStyle BorderRight { get; set; }
        public BorderStyle BorderLeft { get; set; }

        //Font
        public string FontName { get; set; }
        public double FontHeight { get; set; }
        public double FontHeightInPoints { get; set; }
        public bool IsItalic { get; set; }
        public bool IsStrikeout { get; set; }
        public short Color { get; set; }
        public FontSuperScript TypeOffset { get; set; }
        public FontUnderlineType Underline { get; set; }
        public short Charset { get; set; }
        public short Index { get; set; }
        public short Boldweight { get; set; }
        public bool IsBold { get; set; }
        public short FillForegroundColor { get; set; }
        public static MyCellStyle ToMyCellStyle(NPOI.SS.UserModel.IWorkbook workbook ,ICellStyle cellStyle)
        {
            MyCellStyle myCellStyle = new MyCellStyle();

            myCellStyle.VerticalAlignment = cellStyle.VerticalAlignment;
            myCellStyle.Alignment = cellStyle.Alignment;
            myCellStyle.BorderBottom = cellStyle.BorderBottom;
            myCellStyle.BorderTop = cellStyle.BorderTop;
            myCellStyle.BorderRight = cellStyle.BorderRight;
            myCellStyle.BorderLeft = cellStyle.BorderLeft;
            myCellStyle.FillForegroundColor = cellStyle.FillForegroundColor;
     

            IFont font = cellStyle.GetFont(workbook);
            myCellStyle.FontName = font.FontName;
            myCellStyle.FontHeight = font.FontHeight;
            myCellStyle.FontHeightInPoints = font.FontHeightInPoints;
            myCellStyle.IsItalic = font.IsItalic;
            myCellStyle.IsStrikeout = font.IsStrikeout;
            myCellStyle.Color = font.Color;
            myCellStyle.TypeOffset = font.TypeOffset;
            myCellStyle.Underline = font.Underline;
            myCellStyle.Charset = font.Charset;
            myCellStyle.Index = font.Index;
            myCellStyle.Boldweight = font.Boldweight;
            myCellStyle.IsBold = font.IsBold;

            return myCellStyle;
        }
    
    }
    static public class ExcelClass
    {
        public static System.Data.DataTable LoadFile(string FileName)
        {
           
            System.Data.DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                dt.TableName = sheetName;
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }
     
        public static void SaveFile(this System.Data.DataTable dataTable, string filepath)
        {
            try
            {
                SaveFile(dataTable, filepath, new int[] { });
            }
            catch
            {
                Console.WriteLine("找不到Excel安裝資訊!");
            }
    
        }
        public static void SaveFile(this System.Data.DataTable dataTable, string filepath ,int[] ColunmWidth )
        {
            int[] colunmWidth = new int[dataTable.Columns.Count];
            for (int i = 0; i < colunmWidth.Length; i++)
            {
                colunmWidth[i] = 10;
            }
            for(int i = 0 ; i < ColunmWidth.Length ; i ++)
            {
                if (i < colunmWidth.Length) colunmWidth[i] = ColunmWidth[i];
            }
            try
            {
                //建立Excel應用程式類的一個例項，相當於從電腦開始選單開啟Excel
                Microsoft.Office.Interop.Excel.ApplicationClass xlsxapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                //新建一張Excel工作簿
                Microsoft.Office.Interop.Excel.Workbook wbook = xlsxapp.Workbooks.Add(true);
                //第一個sheet頁
                Microsoft.Office.Interop.Excel.Worksheet wsheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets.get_Item(1);
                //將DataTable的列名顯示在Excel表第一行
                wsheet.Cells.NumberFormatLocal = "@";

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    //注意Excel表的行和列的索引都是從1開始的
                    ((Microsoft.Office.Interop.Excel.Range)wsheet.Columns[Convert_Colums(i), System.Type.Missing]).ColumnWidth = colunmWidth[i];
                    wsheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
                }
                //遍歷DataTable，給Excel賦值
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        //從第二行第一列開始寫入資料
                        wsheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                    }
                }

                try
                {
                    //儲存檔案
                    wbook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                    null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                    false, false, null, null, null);
                }
                catch
                {
                    Console.WriteLine("Excel 存檔失敗!");
                }
                finally
                {
                    string ProcessName = "WINWORD";//換成想要結束的進程名字
                    System.Diagnostics.Process[] MyProcess = System.Diagnostics.Process.GetProcessesByName(ProcessName);
                    for (int i = 0; i < MyProcess.Length; i++)
                    {
                        MyProcess[i].Kill();
                    }
                    ProcessName = "EXCEL";//換成想要結束的進程名字
                    MyProcess = System.Diagnostics.Process.GetProcessesByName(ProcessName);
                    for (int i = 0; i < MyProcess.Length; i++)
                    {
                        MyProcess[i].Kill();
                    }
                    //釋放資源
                    xlsxapp.Quit();
                }
            }
            catch
            {
                Console.WriteLine("找不到Excel安裝資訊!");
            }

         
      
        }
        public static byte[] NPOI_GetBytes(this System.Data.DataTable dt)
        {
            return NPOI_GetBytes(dt, Excel_Type.xls);
        }
        public static void NPOI_GetBytes(this System.Data.DataTable dt, Excel_Type excel_Type, Enum[] EnumcolName_ary)
        {
            List<string> colName_ary = new List<string>();
            for (int i = 0; i < EnumcolName_ary.Length; i++)
            {
                colName_ary.Add(EnumcolName_ary[i].GetEnumName());
            }
            NPOI_GetBytes(dt, excel_Type, colName_ary.ToArray());
        }
        public static void NPOI_GetBytes(this System.Data.DataTable dt, Excel_Type excel_Type, string[] colName_ary)
        {
            List<int> list_int = new List<int>();
            for (int i = 0; i < colName_ary.Length; i++)
            {
                int temp = dt.Columns.IndexOf(colName_ary[i]);
                if (temp > 0)
                {
                    list_int.Add(temp);
                }
            }
            NPOI_GetBytes(dt, excel_Type, list_int.ToArray());
        }
        public static byte[] NPOI_GetBytes(this System.Data.DataTable dt, Excel_Type excel_Type, params int[] int_col_ary)
        {
            NPOI.SS.UserModel.IWorkbook workbook;
            if (excel_Type == Excel_Type.xlsx)
            {
                workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            }
            else if (excel_Type == Excel_Type.xls)
            {
                workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            }
            else
            {
                return null;
            }

            NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            // 表头  
            NPOI.SS.UserModel.IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                NPOI.SS.UserModel.ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            // 数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                NPOI.SS.UserModel.IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    NPOI.SS.UserModel.ICell cell = dataRow.CreateCell(j);

                    // 处理指定的数值列
                    bool isNumericColumn = int_col_ary.Contains(j);
                    if (isNumericColumn)
                    {
                        double numericValue;
                        if (double.TryParse(dt.Rows[i][j]?.ToString(), out numericValue))
                        {
                            cell.SetCellValue(numericValue);
                        }
                        else
                        {
                            cell.SetCellValue(0); // 处理无效数值或空值
                        }
                    }
                    else
                    {
                        string cellValue = dt.Rows[i][j]?.ToString() ?? string.Empty;
                        cell.SetCellValue(cellValue);
                    }
                }
            }

            // 写入MemoryStream
            using (MemoryStream stream = new MemoryStream())
            {
                try
                {
                    workbook.Write(stream);
                }
                catch (Exception ex)
                {
                    // 捕获写入异常并打印错误信息
                    Console.WriteLine($"Exception during workbook write: {ex.Message}");
                    return null;
                }

                byte[] buf = stream.ToArray();
                Console.WriteLine($"Generated byte length: {buf.Length}");

                return buf;
            }
        }
        public static byte[] NPOI_LoadToBytes(string filename)
        {
            byte[] fileBytes = File.ReadAllBytes(filename);
            return fileBytes;
        }
   
        public static byte[] NPOI_GetBytes(List<DataTable> dts, params int[] int_col_ary)
        {
            return NPOI_GetBytes(dts, Excel_Type.xls, int_col_ary);
        }
        public static byte[] NPOI_GetBytes(List<DataTable> dts, Excel_Type excel_Type, Enum[] EnumcolName_ary)
        {
            for (int i = 0; i < EnumcolName_ary.Length; i++)
            {
                foreach(DataTable dt in dts)
                {
                    if (dt.Columns.IndexOf(EnumcolName_ary[i].GetEnumName()) > 0)
                    {
                        dt.Columns[EnumcolName_ary[i].GetEnumName()].Caption = "double";
                    }
                }
            }
            return NPOI_GetBytes(dts, excel_Type);
        }
        public static byte[] NPOI_GetBytes(List<DataTable> dts, Excel_Type excel_Type, string[] colName_ary)
        {
            for (int i = 0; i < colName_ary.Length; i++)
            {
                foreach (DataTable dt in dts)
                {
                    if (dt.Columns.IndexOf(colName_ary[i]) > 0)
                    {
                        dt.Columns[colName_ary[i]].Caption = "double";
                    }
                }
            }
            return NPOI_GetBytes(dts, excel_Type);
        }
        public static byte[] NPOI_GetBytes(List<DataTable> dts, Excel_Type excel_Type, params int[] int_col_ary)
        {
            NPOI.SS.UserModel.IWorkbook workbook;
            if (excel_Type == Excel_Type.xlsx) { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (excel_Type == Excel_Type.xls) { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return null; }
            for (int m = 0; m < dts.Count; m++)
            {
                DataTable dt = dts[m];
                NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

                //表头  
                NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    NPOI.SS.UserModel.ICell cell = row.CreateCell(i);
                    cell.SetCellValue(dt.Columns[i].ColumnName);
                }

                //数据  
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
  
                        NPOI.SS.UserModel.ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        bool flag_is_double = false;
                        if(dt.Columns[j].Caption == "double")
                        {
                            if (dt.Rows[i][j].ObjectToString().StringIsEmpty() == false)
                            {
                                cell.SetCellValue(dt.Rows[i][j].ObjectToString().StringToDouble());
                                flag_is_double = true;
                            }
                            else
                            {
                                cell.SetCellValue(0);
                                flag_is_double = true;
                            }
                        }
                        else
                        {
                            for (int k = 0; k < int_col_ary.Length; k++)
                            {
                                if (int_col_ary[k] == j)
                                {
                                    if (dt.Rows[i][j].ObjectToString().StringIsEmpty() == false)
                                    {
                                        cell.SetCellValue(dt.Rows[i][j].ObjectToString().StringToDouble());
                                        flag_is_double = true;
                                        break;
                                    }
                                    else
                                    {
                                        cell.SetCellValue(0);
                                        flag_is_double = true;
                                        break;
                                    }

                                }
                            }
                        }
                       
                        if (!flag_is_double)
                        {
                            
                            if (j == 0)
                            {
                                if (dt.Rows[i][j].ObjectToString().StringIsInt32())
                                {
                                    cell.SetCellValue(dt.Rows[i][j].ObjectToString().StringToInt32().ToString("00000"));
                                }
                                else
                                {
                                    cell.SetCellValue(dt.Rows[i][j].ToString());
                                }                       
                            }
                            else
                            {
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                    }
                }
            }

          
            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            stream.Dispose();
            return buf;
        }
   
        public static byte[] NPOI_GetBytes(this string json)
        {
            return NPOI_GetBytes(json, Excel_Type.xls);
        }
        public static byte[] NPOI_GetBytes(this string json , Excel_Type excel_Type)
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();
            SheetClass sheetClass = json.JsonDeserializet<SheetClass>();
            if (sheetClass == null) return null;

            NPOI.SS.UserModel.IWorkbook workbook;
            if (excel_Type == Excel_Type.xlsx) { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (excel_Type == Excel_Type.xls) { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }

            if (workbook == null) { return null; }
            sheetClass.Init(workbook);


            NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty("Sheet1") ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet("Sheet1");
            for (int i = 0; i < sheetClass.ColumnsWidth.Count; i++)
            {
                sheet.SetColumnWidth(i, sheetClass.ColumnsWidth[i]);
            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
                if (sheet.GetRow(cellValue.RowStart) == null) sheet.CreateRow(cellValue.RowStart);
                if (sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart) == null) sheet.GetRow(cellValue.RowStart).CreateCell(cellValue.ColStart);

            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
                ICell cell = sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart);
                cell.SetCellValue(cellValue.Text);
                cell.CellStyle = sheetClass.GetICellStyle(cellValue.CellStyle_index);

            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
                sheet.GetRow(cellValue.RowStart).Height = cellValue.Height;
                if (!cellValue.Slave)
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd));

                }
            }
            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            stream.Dispose();


            Console.WriteLine($"存檔耗時{myTimerBasic.ToString()}");
            return buf;
        }
        public static byte[] NPOI_GetBytes(this SheetClass sheetClass, Excel_Type excel_Type)
        {
            return NPOI_GetBytes(sheetClass.JsonSerializationt(), excel_Type);
        }
        public static byte[] NPOI_GetBytes(this SheetClass sheetClass)
        {
            return NPOI_GetBytes(sheetClass.JsonSerializationt());
        }
        public static byte[] NPOI_GetBytes(this List<SheetClass> sheetClasses)
        {
            return NPOI_GetBytes(sheetClasses, Excel_Type.xls);
        }
        public static byte[] NPOI_GetBytes(this List<SheetClass> sheetClasses, Excel_Type excel_Type)
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();




            NPOI.SS.UserModel.IWorkbook workbook;
            if (excel_Type == Excel_Type.xlsx) { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (excel_Type == Excel_Type.xls) { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return null; }

            for (int p = 0; p < sheetClasses.Count; p++)
            {
                SheetClass sheetClass = sheetClasses[p];
                sheetClass.Init(workbook);

                sheetClass.Name = sheetClass.Name.Replace("*", "");
                NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty($"{sheetClass.Name}") ? workbook.CreateSheet($"Sheet-{p}") : workbook.CreateSheet($"{sheetClass.Name}-{p}");
                for (int i = 0; i < sheetClass.ColumnsWidth.Count; i++)
                {
                    sheet.SetColumnWidth(i, sheetClass.ColumnsWidth[i]);
                }
                for (int i = 0; i < sheetClass.CellValues.Count; i++)
                {
                    CellValue cellValue = sheetClass.CellValues[i];
                    if (sheet.GetRow(cellValue.RowStart) == null) sheet.CreateRow(cellValue.RowStart);
                    if (sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart) == null) sheet.GetRow(cellValue.RowStart).CreateCell(cellValue.ColStart);

                }
                for (int i = 0; i < sheetClass.CellValues.Count; i++)
                {
                    CellValue cellValue = sheetClass.CellValues[i];
                    ICell cell = sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart);
                    if(cellValue.IsDouble)
                    {
                        cell.SetCellValue(cellValue.Text.StringToDouble());
                    }
                    else
                    {
                        cell.SetCellValue(cellValue.Text);
                    }
                   
                    cell.CellStyle = sheetClass.GetICellStyle(cellValue.CellStyle_index);

                }
                for (int i = 0; i < sheetClass.CellValues.Count; i++)
                {
                    CellValue cellValue = sheetClass.CellValues[i];
                    sheet.GetRow(cellValue.RowStart).Height = cellValue.Height;
                    if (!cellValue.Slave)
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd));

                    }
                }
                //转为字节数组  

            }
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            stream.Dispose();

            return buf;
            Console.WriteLine($"存檔耗時{myTimerBasic.ToString()}");
        }

        public static SheetClass NPOI_GetSheetClass(this System.Data.DataTable dt, Enum[] EnumcolName_ary)
        {
            List<string> colName_ary = new List<string>();
            for (int i = 0; i < EnumcolName_ary.Length; i++)
            {
                colName_ary.Add(EnumcolName_ary[i].GetEnumName());
            }
            return NPOI_GetSheetClass(dt, colName_ary.ToArray());
        }
        public static SheetClass NPOI_GetSheetClass(this System.Data.DataTable dt, string[] colName_ary)
        {
            List<int> list_int = new List<int>();
            for (int i = 0; i < colName_ary.Length; i++)
            {
                int temp = dt.Columns.IndexOf(colName_ary[i]);
                if (temp > 0)
                {
                    list_int.Add(temp);
                }
            }
            return NPOI_GetSheetClass(dt, list_int.ToArray());
        }
        public static SheetClass NPOI_GetSheetClass(this System.Data.DataTable dt, params int[] int_col_ary)
        {
            return NPOI_GetSheetClass(dt, 3000 , int_col_ary);
        }
        public static SheetClass NPOI_GetSheetClass(this System.Data.DataTable dt , int col_width, params int[] int_col_ary)
        {
            SheetClass sheetClass = new SheetClass();
            int row_index = 0;
            for(int i = 0; i < dt.Columns.Count; i++)
            {
            }
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                CellValue cellValue = sheetClass.AddNewCell_Webapi(row_index, row_index, c, c, $"{dt.Columns[c].ColumnName}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
            
            }
            row_index++;
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    CellValue cellValue = sheetClass.AddNewCell_Webapi(row_index, row_index, c, c, $"{dt.Rows[r][c].ToString()}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                    for (int k = 0; k < int_col_ary.Length; k++)
                    {
                        if (int_col_ary[k] == c)
                        {
                            if (cellValue.Text.ObjectToString().StringIsEmpty() == false)
                            {
                                cellValue.IsDouble = true;
                                break;
                            }
                            else
                            {
                                cellValue.IsDouble = true;
                                cellValue.Text = "0";
                                break;
                            }

                        }
                    }
                }
                row_index++;
            }

            return sheetClass;
        }

        public static void NPOI_SaveFile(this System.Data.DataTable dt, string filepath, Enum[] EnumcolName_ary)
        {
            List<string> colName_ary = new List<string>();
            for (int i = 0; i < EnumcolName_ary.Length; i++)
            {
                colName_ary.Add(EnumcolName_ary[i].GetEnumName());
            }
            NPOI_SaveFile(dt, filepath, colName_ary.ToArray());
        }
        public static void NPOI_SaveFile(this System.Data.DataTable dt, string filepath, string[] colName_ary)
        {
            List<int> list_int = new List<int>();
            for (int i = 0; i < colName_ary.Length; i++)
            {
                int temp = dt.Columns.IndexOf(colName_ary[i]);
                if(temp > 0)
                {
                    list_int.Add(temp);
                }
            }
            NPOI_SaveFile(dt, filepath, list_int.ToArray());
        }
        public static void NPOI_SaveFile(this System.Data.DataTable dt, string filepath, params int[] int_col_ary)
        {
            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(filepath).ToLower();
            if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            //表头  
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                NPOI.SS.UserModel.ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    NPOI.SS.UserModel.ICell cell = row1.CreateCell(j);
                    bool flag_is_double = false;
                    for (int k = 0; k < int_col_ary.Length; k++)
                    {
                        if (int_col_ary[k] == j)
                        {
                            if (dt.Rows[i][j].ObjectToString().StringIsEmpty() == false)
                            {
                                cell.SetCellValue(dt.Rows[i][j].ObjectToString().StringToDouble());
                                flag_is_double = true;
                                break;
                            }
                            else
                            {
                                cell.SetCellValue(0);
                                flag_is_double = true;
                                break;
                            }

                        }
                    }
                    if (!flag_is_double)
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        //if (j == 0)
                        //{
                        //    cell.SetCellValue(dt.Rows[i][j].ObjectToString().StringToInt32().ToString("00000"));
                        //}
                        //else
                        //{
                        //    cell.SetCellValue(dt.Rows[i][j].ToString());
                        //}

                    }
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }


        }
        public static void NPOI_SaveFile(this SheetClass sheetClass, string file)
        {
            NPOI_SaveFile(sheetClass.JsonSerializationt(), file);
        }
        public static void NPOI_SaveFile(this List<SheetClass> sheetClasses, string file)
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

         
           

            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }

            for (int p = 0; p < sheetClasses.Count; p++)
            {
                SheetClass sheetClass = sheetClasses[p];
                sheetClass.Init(workbook);


                NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty($"{sheetClass.Name}") ? workbook.CreateSheet($"Sheet{p}") : workbook.CreateSheet($"{sheetClass.Name}");
                for (int i = 0; i < sheetClass.ColumnsWidth.Count; i++)
                {
                    sheet.SetColumnWidth(i, sheetClass.ColumnsWidth[i]);
                }
                for (int i = 0; i < sheetClass.CellValues.Count; i++)
                {
                    CellValue cellValue = sheetClass.CellValues[i];
                    if (sheet.GetRow(cellValue.RowStart) == null) sheet.CreateRow(cellValue.RowStart);
                    if (sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart) == null) sheet.GetRow(cellValue.RowStart).CreateCell(cellValue.ColStart);

                }
                for (int i = 0; i < sheetClass.CellValues.Count; i++)
                {
                    CellValue cellValue = sheetClass.CellValues[i];
                    ICell cell = sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart);
                    cell.SetCellValue(cellValue.Text);
                    cell.CellStyle = sheetClass.GetICellStyle(cellValue.CellStyle_index);

                }
                for (int i = 0; i < sheetClass.CellValues.Count; i++)
                {
                    CellValue cellValue = sheetClass.CellValues[i];
                    sheet.GetRow(cellValue.RowStart).Height = cellValue.Height;
                    if (!cellValue.Slave)
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd));

                    }
                }
                //转为字节数组  
             
            }
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            try
            {
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message} {DateTime.Now.ToDateTimeString()}");
            }

            Console.WriteLine($"存檔耗時{myTimerBasic.ToString()}");
        }
        public static void NPOI_SaveFile(this string json, string file)
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();
            SheetClass sheetClass = json.JsonDeserializet<SheetClass>();
            if (sheetClass == null) return;
             
            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            sheetClass.Init(workbook);
    

            NPOI.SS.UserModel.ISheet sheet = string.IsNullOrEmpty("Sheet1") ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet("Sheet1");
            for (int i = 0; i < sheetClass.ColumnsWidth.Count; i++)
            {
                sheet.SetColumnWidth(i, sheetClass.ColumnsWidth[i]);
            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];         
                if (sheet.GetRow(cellValue.RowStart) == null) sheet.CreateRow(cellValue.RowStart);
                if (sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart) == null) sheet.GetRow(cellValue.RowStart).CreateCell(cellValue.ColStart);
              
            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
                ICell cell = sheet.GetRow(cellValue.RowStart).GetCell(cellValue.ColStart);
                cell.SetCellValue(cellValue.Text);
                cell.CellStyle = sheetClass.GetICellStyle(cellValue.CellStyle_index);
              
            }
            for (int i = 0; i < sheetClass.CellValues.Count; i++)
            {
                CellValue cellValue = sheetClass.CellValues[i];
                sheet.GetRow(cellValue.RowStart).Height = cellValue.Height;
                if (!cellValue.Slave)
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd));
                    
                }         
            }
            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            try
            {
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message} {DateTime.Now.ToDateTimeString()}");
            }
            //保存为Excel文件  
          
 
            Console.WriteLine($"存檔耗時{myTimerBasic.ToString()}");
        }
        public static List<SheetClass> NPOI_LoadToSheetClasses(this string file)
        {
            List<SheetClass> sheetClasses = NPOI_LoadSheetsToJson(file).JsonDeserializet<List<SheetClass>>();
            return sheetClasses;
        }
        public static SheetClass NPOI_LoadToSheetClass(this string file)
        {
            SheetClass sheetClass = NPOI_LoadSheetToJson(file).JsonDeserializet<SheetClass>();
            return sheetClass;
        }
        public static string NPOI_LoadSheetsToJson(this string file)
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

            string result = "";
            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            try
            {
                FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                List<SheetClass> sheetClasses = new List<SheetClass>();
                for (int num = 0; num < workbook.NumberOfSheets; num++)
                {
                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(num);
                    SheetClass sheetClass = new SheetClass(sheet.SheetName);
                    List<ICell> cells = new List<ICell>();

                    for (int r = 0; r <= sheet.LastRowNum; r++)
                    {
                        for (int c = 0; c < sheet.GetRow(r).LastCellNum; c++)
                        {

                            if (r == 0)
                            {
                                sheetClass.ColumnsWidth.Add(sheet.GetColumnWidth(c));
                            }
                            CellValue cellValue = new CellValue();
                            ICell cell = sheet.GetRow(r).GetCell(c);

                            object obj = NPOI_GetValueType(cell);
                            if (cell == null || obj == null) continue;
                            if (obj != null)
                            {
                                cellValue.Text = obj.ObjectToString();
                            }
                            cellValue.Height = cell.Row.Height;
                            bool flag_IsMergedCell = cell.IsMergedCell;

                            if (flag_IsMergedCell)
                            {
                                sheet.NPOI_IsMergeCell(r, c, ref cellValue);
                            }
                            else
                            {
                                cellValue.RowStart = r;
                                cellValue.RowEnd = r;
                                cellValue.ColStart = c;
                                cellValue.ColEnd = c;
                                cellValue.Slave = false;
                            }
                            CellValue cellValue_buf = sheetClass.SortCellValue(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd);
                            if (cellValue_buf == null && flag_IsMergedCell == true)
                            {

                                ICell cell_end = sheet.GetRow(cellValue.RowEnd).GetCell(cellValue.ColEnd);
                                cell.CellStyle.BorderRight = cell_end.CellStyle.BorderRight;
                                cell.CellStyle.BorderBottom = cell_end.CellStyle.BorderBottom;
                                cellValue.Slave = false;
                            }
                            else if (cellValue_buf != null && flag_IsMergedCell == true)
                            {
                                cellValue.RowStart = r;
                                cellValue.RowEnd = r;
                                cellValue.ColStart = c;
                                cellValue.ColEnd = c;
                                cellValue.Slave = true;
                            }


                            MyCellStyle myCellStyle = MyCellStyle.ToMyCellStyle(workbook, cell.CellStyle);
                            sheetClass.Add(cellValue, myCellStyle);

                        }
                    }
                    sheetClasses.Add(sheetClass);

                }
                result = sheetClasses.JsonSerializationt(false);
                //Console.WriteLine($"{result}");
                fs.Close();
                fs.Dispose();
                workbook.Close();
                Console.WriteLine($"讀檔耗時{myTimerBasic.ToString()}");

            }
            catch
            {
                Console.WriteLine($"NPOI_LoadHeader 檔案已開啟!無法讀取! , 位置 : {file}");
                return "[]";
            }
            finally
            {

            }


            return result;

        }
        public static string NPOI_LoadSheetsToJson(byte[] bytes, string fileExt = ".xlsx")
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

            string result = "";
            NPOI.SS.UserModel.IWorkbook workbook;
            try
            {
                MemoryStream fs = new MemoryStream(bytes);
                if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                List<SheetClass> sheetClasses = new List<SheetClass>();
                for (int num = 0; num < workbook.NumberOfSheets; num++)
                {
                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(num);
                    SheetClass sheetClass = new SheetClass(sheet.SheetName);
                    List<ICell> cells = new List<ICell>();

                    for (int r = 0; r <= sheet.LastRowNum; r++)
                    {
                        for (int c = 0; c < sheet.GetRow(r).LastCellNum; c++)
                        {

                            if (r == 0)
                            {
                                sheetClass.ColumnsWidth.Add(sheet.GetColumnWidth(c));
                            }
                            CellValue cellValue = new CellValue();
                            ICell cell = sheet.GetRow(r).GetCell(c);

                            object obj = NPOI_GetValueType(cell);
                            if (obj != null)
                            {
                                cellValue.Text = obj.ObjectToString();
                            }
                            cellValue.Height = cell.Row.Height;
                            bool flag_IsMergedCell = cell.IsMergedCell;

                            if (flag_IsMergedCell)
                            {
                                sheet.NPOI_IsMergeCell(r, c, ref cellValue);
                            }
                            else
                            {
                                cellValue.RowStart = r;
                                cellValue.RowEnd = r;
                                cellValue.ColStart = c;
                                cellValue.ColEnd = c;
                                cellValue.Slave = false;
                            }
                            CellValue cellValue_buf = sheetClass.SortCellValue(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd);
                            if (cellValue_buf == null && flag_IsMergedCell == true)
                            {

                                ICell cell_end = sheet.GetRow(cellValue.RowEnd).GetCell(cellValue.ColEnd);
                                cell.CellStyle.BorderRight = cell_end.CellStyle.BorderRight;
                                cell.CellStyle.BorderBottom = cell_end.CellStyle.BorderBottom;
                                cellValue.Slave = false;
                            }
                            else if (cellValue_buf != null && flag_IsMergedCell == true)
                            {
                                cellValue.RowStart = r;
                                cellValue.RowEnd = r;
                                cellValue.ColStart = c;
                                cellValue.ColEnd = c;
                                cellValue.Slave = true;
                            }


                            MyCellStyle myCellStyle = MyCellStyle.ToMyCellStyle(workbook, cell.CellStyle);
                            sheetClass.Add(cellValue, myCellStyle);

                        }
                    }
                    sheetClasses.Add(sheetClass);

                }
                result = sheetClasses.JsonSerializationt(false);
                //Console.WriteLine($"{result}");
                fs.Close();
                fs.Dispose();
                workbook.Close();
                Console.WriteLine($"讀檔耗時{myTimerBasic.ToString()}");

            }
            catch
            {
                Console.WriteLine($"NPOI_LoadHeader 檔案已開啟!無法讀取!");
                return "[]";
            }
            finally
            {

            }


            return result;

        }
        public static string NPOI_LoadSheetToJson(this string file)
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

            string result = "";
            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            try
            {
                FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
                
                SheetClass sheetClass = new SheetClass(sheet.SheetName);
                List<ICell> cells = new List<ICell>();
                
                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    for (int c = 0; c < sheet.GetRow(r).LastCellNum; c++)
                    {

                        if (r == 0)
                        {
                            sheetClass.ColumnsWidth.Add(sheet.GetColumnWidth(c));
                        }
                        CellValue cellValue = new CellValue();
                        ICell cell = sheet.GetRow(r).GetCell(c);
                     
                        object obj = NPOI_GetValueType(cell);
                        if (obj != null)
                        {
                            cellValue.Text = obj.ObjectToString();
                        }
                        if (cell == null) continue;
                        cellValue.Height = cell.Row.Height;
                        bool flag_IsMergedCell = cell.IsMergedCell;

                        if (flag_IsMergedCell)
                        {
                            sheet.NPOI_IsMergeCell(r, c, ref cellValue);
                        }
                        else
                        {
                            cellValue.RowStart = r;
                            cellValue.RowEnd = r;
                            cellValue.ColStart = c;
                            cellValue.ColEnd = c;
                            cellValue.Slave = false;
                        }
                        CellValue cellValue_buf = sheetClass.SortCellValue(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd);
                        if (cellValue_buf == null && flag_IsMergedCell == true)
                        {

                            ICell cell_end = sheet.GetRow(cellValue.RowEnd).GetCell(cellValue.ColEnd);
                            cell.CellStyle.BorderRight = cell_end.CellStyle.BorderRight;
                            cell.CellStyle.BorderBottom = cell_end.CellStyle.BorderBottom;
                            cellValue.Slave = false;
                        }
                        else if (cellValue_buf != null && flag_IsMergedCell == true)
                        {
                            cellValue.RowStart = r;
                            cellValue.RowEnd = r;
                            cellValue.ColStart = c;
                            cellValue.ColEnd = c;
                            cellValue.Slave = true;
                        }

      
                        MyCellStyle myCellStyle = MyCellStyle.ToMyCellStyle(workbook, cell.CellStyle);
                        sheetClass.Add(cellValue, myCellStyle);

                    }
                }
                result = sheetClass.JsonSerializationt(false);
                //Console.WriteLine($"{result}");
                fs.Close();
                fs.Dispose();
                workbook.Close();
                Console.WriteLine($"讀檔耗時{myTimerBasic.ToString()}");
            }
            catch
            {
                Console.WriteLine($"NPOI_LoadHeader 檔案已開啟!無法讀取! , 位置 : {file}");
                return "[]";
            }
            finally
            {

            }
 
    
            return result;
        }
        public static string NPOI_LoadSheetToJson(byte[] bytes, string fileExt = ".xlsx")
        {
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

            string result = "";
            NPOI.SS.UserModel.IWorkbook workbook;
            try
            {
                MemoryStream fs = new MemoryStream(bytes);
                if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);

                SheetClass sheetClass = new SheetClass(sheet.SheetName);
                List<ICell> cells = new List<ICell>();

                for (int r = 0; r <= sheet.LastRowNum; r++)
                {
                    for (int c = 0; c < sheet.GetRow(r).LastCellNum; c++)
                    {

                        if (r == 0)
                        {
                            sheetClass.ColumnsWidth.Add(sheet.GetColumnWidth(c));
                        }
                        CellValue cellValue = new CellValue();
                        ICell cell = sheet.GetRow(r).GetCell(c);

                        object obj = NPOI_GetValueType(cell);
                        if (obj != null)
                        {
                            cellValue.Text = obj.ObjectToString();
                        }
                        if (cell == null) continue;
                        cellValue.Height = cell.Row.Height;
                        bool flag_IsMergedCell = cell.IsMergedCell;

                        if (flag_IsMergedCell)
                        {
                            sheet.NPOI_IsMergeCell(r, c, ref cellValue);
                        }
                        else
                        {
                            cellValue.RowStart = r;
                            cellValue.RowEnd = r;
                            cellValue.ColStart = c;
                            cellValue.ColEnd = c;
                            cellValue.Slave = false;
                        }
                        CellValue cellValue_buf = sheetClass.SortCellValue(cellValue.RowStart, cellValue.RowEnd, cellValue.ColStart, cellValue.ColEnd);
                        if (cellValue_buf == null && flag_IsMergedCell == true)
                        {

                            ICell cell_end = sheet.GetRow(cellValue.RowEnd).GetCell(cellValue.ColEnd);
                            cell.CellStyle.BorderRight = cell_end.CellStyle.BorderRight;
                            cell.CellStyle.BorderBottom = cell_end.CellStyle.BorderBottom;
                            cellValue.Slave = false;
                        }
                        else if (cellValue_buf != null && flag_IsMergedCell == true)
                        {
                            cellValue.RowStart = r;
                            cellValue.RowEnd = r;
                            cellValue.ColStart = c;
                            cellValue.ColEnd = c;
                            cellValue.Slave = true;
                        }


                        MyCellStyle myCellStyle = MyCellStyle.ToMyCellStyle(workbook, cell.CellStyle);
                        sheetClass.Add(cellValue, myCellStyle);

                    }
                }
                result = sheetClass.JsonSerializationt(false);
                //Console.WriteLine($"{result}");
                fs.Close();
                fs.Dispose();
                workbook.Close();
                Console.WriteLine($"讀檔耗時{myTimerBasic.ToString()}");
            }
            catch
            {
                Console.WriteLine($"NPOI_LoadHeader 檔案已開啟!無法讀取!");
                return "[]";
            }
            finally
            {

            }


            return result;
        }


        public static DataTable NPOI_LoadFile(byte[] bytes , string fileExt = ".xlsx")
        {
            try
            {
                DataTable dt = new DataTable();
                NPOI.SS.UserModel.IWorkbook workbook;
                using (MemoryStream fs = new MemoryStream(bytes))
                {
                    //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                    if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                    if (workbook == null) { return null; }
                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
                    //表头
                    NPOI.SS.UserModel.IRow header = sheet.GetRow(sheet.FirstRowNum);
                    List<int> columns = new List<int>();
                    for (int i = 0; i < header.LastCellNum; i++)
                    {
                        object obj = NPOI_GetValueType(header.GetCell(i));
                        if (obj == null || obj.ToString() == string.Empty)
                        {
                            dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        }
                        columns.Add(i);
                    }
                    //数据
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        foreach (int j in columns)
                        {
                            dr[j] = NPOI_GetValueType(sheet.GetRow(i).GetCell(j));
                            string temp = "";
                            if (sheet.GetRow(i).GetCell(j) != null)
                            {
                                temp = sheet.GetRow(i).GetCell(j).ToString();
                                if (temp.Check_Date_String())
                                {
                                    temp = temp.StringToDateTime().ToDateTimeString();
                                }
                            }
                          
                            dr[j] = temp;
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                        if (hasValue)
                        {
                            dt.Rows.Add(dr);
                        }
                    }
                }
                return dt;
            }
            catch
            {
                return null;
            }
        }
        public static DataTable NPOI_LoadFile(this string file)
        {
            try
            {
                DataTable dt = new DataTable();
                NPOI.SS.UserModel.IWorkbook workbook;
                string fileExt = Path.GetExtension(file).ToLower();
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                    if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                    if (workbook == null) { return null; }
                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
                    //表头
                    NPOI.SS.UserModel.IRow header = sheet.GetRow(sheet.FirstRowNum);
                    List<int> columns = new List<int>();
                    for (int i = 0; i < header.LastCellNum; i++)
                    {
                        object obj = NPOI_GetValueType(header.GetCell(i));
                        if (obj == null || obj.ToString() == string.Empty)
                        {
                            dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        }
                        columns.Add(i);
                    }
                    //数据
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        foreach (int j in columns)
                        {
                            dr[j] = NPOI_GetValueType(sheet.GetRow(i).GetCell(j));
                            string temp = "";
                            if (sheet.GetRow(i).GetCell(j) != null)
                            {
                                temp = sheet.GetRow(i).GetCell(j).ToString();
                                if (temp.Check_Date_String())
                                {
                                    temp = temp.StringToDateTime().ToDateTimeString();
                                }
                            }

                            dr[j] = temp;
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                        if (hasValue)
                        {
                            dt.Rows.Add(dr);
                        }
                    }
                }
                return dt;
            }
            catch
            {
                return null;
            }        
        }
        public static List<DataTable> NPOI_LoadFile2DataTables(this string file)
        {
            List<DataTable> dataTables = new List<DataTable>();
            Basic.MyTimerBasic myTimerBasic = new Basic.MyTimerBasic(100000);
            myTimerBasic.StartTickTime();

            NPOI.SS.UserModel.IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            try
            {
                FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                if (fileExt == ".xlsx") { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                for (int num = 0; num < workbook.NumberOfSheets; num++)
                {
                    DataTable dt = new DataTable();
                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(num);
                    List<ICell> cells = new List<ICell>();
                    dt.TableName = sheet.SheetName;
                    //表头
                    NPOI.SS.UserModel.IRow header = sheet.GetRow(sheet.FirstRowNum);
                    List<int> columns = new List<int>();
                    for (int i = 0; i < header.LastCellNum; i++)
                    {
                        object obj = NPOI_GetValueType(header.GetCell(i));
                        if (obj == null || obj.ToString() == string.Empty)
                        {
                            dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        }
                        columns.Add(i);
                    }
                    //数据
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        foreach (int j in columns)
                        {
                            dr[j] = NPOI_GetValueType(sheet.GetRow(i).GetCell(j));
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                        if (hasValue)
                        {
                            dt.Rows.Add(dr);
                        }
                    }
                    dataTables.Add(dt);

                }
                fs.Close();
                fs.Dispose();
                workbook.Close();
                Console.WriteLine($"讀檔耗時{myTimerBasic.ToString()}");

            }
            catch(Exception ex)
            {
                Console.WriteLine($"NPOI_LoadHeader 檔案已開啟!無法讀取! , 位置 : {file} , Exception : {ex.Message}");
                return dataTables;
            }
            finally
            {

            }
            return dataTables;
        }
        public static void NPOI_SaveFiles2Folder(this string file, string path, string fileExt = ".xls", params int[] int_col_ary)
        {
            List<DataTable> dataTables = NPOI_LoadFile2DataTables(file);
            NPOI_SaveDataTables2Folder(dataTables, path, fileExt, int_col_ary);
        }
        public static void NPOI_SaveDataTables2Folder(this List<DataTable> dataTables, string path, string fileExt = ".xls", params int[] int_col_ary)
        {
            for (int i = 0; i < dataTables.Count; i++)
            {
                string name = dataTables[i].TableName;
                string filename = $@"{path}\{name}{fileExt}";
                dataTables[i].NPOI_SaveFile(filename, int_col_ary);
            }
        }


        private static object NPOI_GetValueType(this NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case NPOI.SS.UserModel.CellType.Blank: //BLANK:  
                    return null;
                case NPOI.SS.UserModel.CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case NPOI.SS.UserModel.CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case NPOI.SS.UserModel.CellType.String: //STRING:  
                    return cell.StringCellValue;
                case NPOI.SS.UserModel.CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case NPOI.SS.UserModel.CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }
        private static bool NPOI_IsMergeCell(this ISheet sheet, int rowIndex, int colIndex, ref CellValue cellValue)
        {
            bool result = false;
            cellValue.RowStart = 0;
            cellValue.RowEnd = 0;
            cellValue.ColStart = 0;
            cellValue.ColEnd = 0;
            if ((rowIndex < 0) || (colIndex < 0)) return result;
            int regionsCount = sheet.NumMergedRegions;
            for (int i = 0; i < regionsCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                //sheet.IsMergedRegion(range); 
                if (rowIndex >= range.FirstRow && rowIndex <= range.LastRow && colIndex >= range.FirstColumn && colIndex <= range.LastColumn)
                {
                    cellValue.RowStart = range.FirstRow;
                    cellValue.RowEnd = range.LastRow;
                    cellValue.ColStart = range.FirstColumn;
                    cellValue.ColEnd = range.LastColumn;
                    result = true;
                    break;
                }
            }
            return result;
        }

        private static string Convert_Colums(int index)
        {
            string str = "";
            int value_00 = -1;
            int value_01 = -1;
            value_00 = index % 26;
            if (index >= 26)
            {
                value_01 = index / 26;
                value_01--;
            }
            if (value_01 != -1)
            {
                str += (char)(value_01 + 65);
            }
            if (value_00 != -1)
            {
                str += (char)(value_00 + 65);
            }
            
            return str;
        }

        private static Color HexToColor(this string hex)
        {
            string[] str_array = hex.Split(':');
            if (str_array.Length == 3)
            {
                int R = str_array[0].StringHexToint() / 256;
                int G = str_array[1].StringHexToint() / 256;
                int B = str_array[2].StringHexToint() / 256;

                return Color.FromArgb(R, G, B);
            }
            return Color.Black;
        }
        public static Color ToColor(this NPOI_Color nPOI_Color)
        {

            if (nPOI_Color == NPOI_Color.BLACK)
            {
                NPOI.HSSF.Util.HSSFColor.Black hSSFColor = new NPOI.HSSF.Util.HSSFColor.Black();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.Basic)
            {
                NPOI.HSSF.Util.HSSFColor.White hSSFColor = new NPOI.HSSF.Util.HSSFColor.White();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.BLACK2)
            {
                NPOI.HSSF.Util.HSSFColor.Black hSSFColor = new NPOI.HSSF.Util.HSSFColor.Black();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.BROWN)
            {
                NPOI.HSSF.Util.HSSFColor.Brown hSSFColor = new NPOI.HSSF.Util.HSSFColor.Brown();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.OLIVE_GREEN)
            {
                NPOI.HSSF.Util.HSSFColor.OliveGreen hSSFColor = new NPOI.HSSF.Util.HSSFColor.OliveGreen();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.DARK_GREEN)
            {
                NPOI.HSSF.Util.HSSFColor.DarkGreen hSSFColor = new NPOI.HSSF.Util.HSSFColor.DarkGreen();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.DARK_TEAL)
            {
                NPOI.HSSF.Util.HSSFColor.DarkTeal hSSFColor = new NPOI.HSSF.Util.HSSFColor.DarkTeal();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.DARK_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.DarkBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.DarkBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.INDIGO)
            {
                NPOI.HSSF.Util.HSSFColor.Indigo hSSFColor = new NPOI.HSSF.Util.HSSFColor.Indigo();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.GREY_80_PERCENT)
            {
                NPOI.HSSF.Util.HSSFColor.Grey80Percent hSSFColor = new NPOI.HSSF.Util.HSSFColor.Grey80Percent();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.DARK_RED)
            {
                NPOI.HSSF.Util.HSSFColor.DarkRed hSSFColor = new NPOI.HSSF.Util.HSSFColor.DarkRed();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.ORANGE)
            {
                NPOI.HSSF.Util.HSSFColor.Orange hSSFColor = new NPOI.HSSF.Util.HSSFColor.Orange();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.DARK_YELLOW)
            {
                NPOI.HSSF.Util.HSSFColor.DarkYellow hSSFColor = new NPOI.HSSF.Util.HSSFColor.DarkYellow();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }

            else if (nPOI_Color == NPOI_Color.GREEN)
            {
                NPOI.HSSF.Util.HSSFColor.Green hSSFColor = new NPOI.HSSF.Util.HSSFColor.Green();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.TEAL)
            {
                NPOI.HSSF.Util.HSSFColor.Teal hSSFColor = new NPOI.HSSF.Util.HSSFColor.Teal();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.Blue hSSFColor = new NPOI.HSSF.Util.HSSFColor.Blue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.BLUE_GREY)
            {
                NPOI.HSSF.Util.HSSFColor.BlueGrey hSSFColor = new NPOI.HSSF.Util.HSSFColor.BlueGrey();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.GREY_50_PERCENT)
            {
                NPOI.HSSF.Util.HSSFColor.Grey50Percent hSSFColor = new NPOI.HSSF.Util.HSSFColor.Grey50Percent();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.RED)
            {
                NPOI.HSSF.Util.HSSFColor.Red hSSFColor = new NPOI.HSSF.Util.HSSFColor.Red();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIGHT_ORANGE)
            {
                NPOI.HSSF.Util.HSSFColor.LightOrange hSSFColor = new NPOI.HSSF.Util.HSSFColor.LightOrange();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIME)
            {
                NPOI.HSSF.Util.HSSFColor.Lime hSSFColor = new NPOI.HSSF.Util.HSSFColor.Lime();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.SEA_GREEN)
            {
                NPOI.HSSF.Util.HSSFColor.SeaGreen hSSFColor = new NPOI.HSSF.Util.HSSFColor.SeaGreen();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.AQUA)
            {
                NPOI.HSSF.Util.HSSFColor.Aqua hSSFColor = new NPOI.HSSF.Util.HSSFColor.Aqua();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIGHT_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.LightBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.LightBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.VIOLET)
            {
                NPOI.HSSF.Util.HSSFColor.Violet hSSFColor = new NPOI.HSSF.Util.HSSFColor.Violet();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.GREY_40_PERCENT)
            {
                NPOI.HSSF.Util.HSSFColor.Grey40Percent hSSFColor = new NPOI.HSSF.Util.HSSFColor.Grey40Percent();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.PINK)
            {
                NPOI.HSSF.Util.HSSFColor.Pink hSSFColor = new NPOI.HSSF.Util.HSSFColor.Pink();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.GOLD)
            {
                NPOI.HSSF.Util.HSSFColor.Gold hSSFColor = new NPOI.HSSF.Util.HSSFColor.Gold();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.YELLOW)
            {
                NPOI.HSSF.Util.HSSFColor.Yellow hSSFColor = new NPOI.HSSF.Util.HSSFColor.Yellow();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.BRIGHT_GREEN)
            {
                NPOI.HSSF.Util.HSSFColor.BrightGreen hSSFColor = new NPOI.HSSF.Util.HSSFColor.BrightGreen();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.TURQUOISE)
            {
                NPOI.HSSF.Util.HSSFColor.Turquoise hSSFColor = new NPOI.HSSF.Util.HSSFColor.Turquoise();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.SKY_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.SkyBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.SkyBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.PLUM)
            {
                NPOI.HSSF.Util.HSSFColor.Plum hSSFColor = new NPOI.HSSF.Util.HSSFColor.Plum();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.GREY_25_PERCENT)
            {
                NPOI.HSSF.Util.HSSFColor.Grey25Percent hSSFColor = new NPOI.HSSF.Util.HSSFColor.Grey25Percent();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.ROSE)
            {
                NPOI.HSSF.Util.HSSFColor.Rose hSSFColor = new NPOI.HSSF.Util.HSSFColor.Rose();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.TAN)
            {
                NPOI.HSSF.Util.HSSFColor.Tan hSSFColor = new NPOI.HSSF.Util.HSSFColor.Tan();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIGHT_YELLOW)
            {
                NPOI.HSSF.Util.HSSFColor.LightYellow hSSFColor = new NPOI.HSSF.Util.HSSFColor.LightYellow();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIGHT_GREEN)
            {
                NPOI.HSSF.Util.HSSFColor.LightGreen hSSFColor = new NPOI.HSSF.Util.HSSFColor.LightGreen();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIGHT_TURQUOISE)
            {
                NPOI.HSSF.Util.HSSFColor.LightTurquoise hSSFColor = new NPOI.HSSF.Util.HSSFColor.LightTurquoise();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.PALE_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.PaleBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.PaleBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LAVENDER)
            {
                NPOI.HSSF.Util.HSSFColor.Lavender hSSFColor = new NPOI.HSSF.Util.HSSFColor.Lavender();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.WHITE)
            {
                NPOI.HSSF.Util.HSSFColor.White hSSFColor = new NPOI.HSSF.Util.HSSFColor.White();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.CORNFLOWER_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.CornflowerBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.CornflowerBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LEMON_CHIFFON)
            {
                NPOI.HSSF.Util.HSSFColor.LemonChiffon hSSFColor = new NPOI.HSSF.Util.HSSFColor.LemonChiffon();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.MAROON)
            {
                NPOI.HSSF.Util.HSSFColor.Maroon hSSFColor = new NPOI.HSSF.Util.HSSFColor.Maroon();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.ORCHID)
            {
                NPOI.HSSF.Util.HSSFColor.Orchid hSSFColor = new NPOI.HSSF.Util.HSSFColor.Orchid();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.CORAL)
            {
                NPOI.HSSF.Util.HSSFColor.Coral hSSFColor = new NPOI.HSSF.Util.HSSFColor.Coral();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.ROYAL_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.RoyalBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.RoyalBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.LIGHT_CORNFLOWER_BLUE)
            {
                NPOI.HSSF.Util.HSSFColor.LightCornflowerBlue hSSFColor = new NPOI.HSSF.Util.HSSFColor.LightCornflowerBlue();
                string hex = hSSFColor.GetHexString();
                return hex.HexToColor();
            }
            else if (nPOI_Color == NPOI_Color.AUTOMATIC)
            {
                NPOI.HSSF.Util.HSSFColor.Automatic hSSFColor = new NPOI.HSSF.Util.HSSFColor.Automatic();
                string hex = hSSFColor.GetHexString();
                return Color.White;
            }
            return Color.Black;
        }

        public static byte[] ConvertXlsToXlsx(this byte[] xlsBytes)
        {
            using (MemoryStream xlsStream = new MemoryStream(xlsBytes))
            {
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook(xlsStream);

                // 创建一个新的 .xlsx 工作簿
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

                // 用于存储样式的映射，避免重复创建相同样式
                var styleMap = new Dictionary<ICellStyle, ICellStyle>();

                for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
                {
                    ISheet oldSheet = hssfWorkbook.GetSheetAt(i);
                    ISheet newSheet = xssfWorkbook.CreateSheet(oldSheet.SheetName);

                    CopySheet(oldSheet, newSheet, styleMap);
                }

                using (MemoryStream xlsxStream = new MemoryStream())
                {
                    xssfWorkbook.Write(xlsxStream);
                    return xlsxStream.ToArray();
                }
            }
        }

        private static void CopySheet(ISheet oldSheet, ISheet newSheet, Dictionary<ICellStyle, ICellStyle> styleMap)
        {
            for (int i = 0; i <= oldSheet.LastRowNum; i++)
            {
                IRow oldRow = oldSheet.GetRow(i);
                IRow newRow = newSheet.CreateRow(i);

                if (oldRow != null)
                {
                    CopyRow(oldRow, newRow, styleMap);
                }
            }
        }

        private static void CopyRow(IRow oldRow, IRow newRow, Dictionary<ICellStyle, ICellStyle> styleMap)
        {
            newRow.Height = oldRow.Height;

            for (int i = 0; i < oldRow.LastCellNum; i++)
            {
                ICell oldCell = oldRow.GetCell(i);
                ICell newCell = newRow.CreateCell(i);

                if (oldCell != null)
                {
                    newCell.SetCellValue(oldCell.ToString());

                    // 复制样式
                    if (oldCell.CellStyle != null)
                    {
                        if (!styleMap.TryGetValue(oldCell.CellStyle, out var newCellStyle))
                        {
                            newCellStyle = CopyCellStyle(oldCell.CellStyle, newCell.Sheet.Workbook);
                            styleMap[oldCell.CellStyle] = newCellStyle;
                        }
                        newCell.CellStyle = newCellStyle;
                    }

                    // 复制公式
                    if (oldCell.CellType == CellType.Formula)
                    {
                        newCell.SetCellFormula(oldCell.CellFormula);
                    }
                }
            }
        }
        private static ICellStyle CopyCellStyle(ICellStyle oldCellStyle, IWorkbook newWorkbook)
        {
            ICellStyle newCellStyle = newWorkbook.CreateCellStyle();

            // 复制基本样式属性
            newCellStyle.Alignment = oldCellStyle.Alignment;
            newCellStyle.VerticalAlignment = oldCellStyle.VerticalAlignment;
            newCellStyle.BorderBottom = oldCellStyle.BorderBottom;
            newCellStyle.BorderLeft = oldCellStyle.BorderLeft;
            newCellStyle.BorderRight = oldCellStyle.BorderRight;
            newCellStyle.BorderTop = oldCellStyle.BorderTop;
            newCellStyle.FillForegroundColor = oldCellStyle.FillForegroundColor;
            newCellStyle.FillPattern = oldCellStyle.FillPattern;
            newCellStyle.FillBackgroundColor = oldCellStyle.FillBackgroundColor;

            // 如果有其他需要复制的属性，继续添加复制代码
            // ...

            return newCellStyle;
        }
    }
}
