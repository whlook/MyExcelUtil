////////////////////////////////////////////////>
// 该Excel库是基于NPOI的Excel简单操作的封装。
// 直接使用行列指定操作的单元格对象，
// 不需要考虑单元格是否已经创建，只需要
// 提供已经创建的工作簿或工作表。
// 这个库是我从项目中整理出来的，因为有些地方
// 和项目联系较多，所以代码看起来封装的不是很
// 好，在代码中也存在一些不安全的地方，这在
// 今后将逐步改进。
////////////////////////////////////////////////>
// 作者：吴辉
// 日期：2017-5-27
////////////////////////////////////////////////>

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using NPOI.SS.Format;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.Util.Collections;



namespace ExcelTool
{
    public class ExcelUtils
    {

        #region 将值写入单元格
        /// <summary>
        /// 在单元格中写入值,自动判断是否是数字
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <param name="val">值string</param>
        /// <param name="sheet">目标表格</param>
        /// <returns></returns>
        public static bool WriteCell(int row, int col, string val, ref HSSFSheet sheet)
        {

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);
                if (isNumber(val))
                {

                    double valNum = double.Parse(val);

                    t_cell.SetCellValue(valNum);
                }
                else
                    t_cell.SetCellValue(val);
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    if (isNumber(val))
                    {

                        double valNum = double.Parse(val);
                        t_cell.SetCellValue(valNum);
                    }
                    else
                        t_cell.SetCellValue(val);
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    if (isNumber(val))
                    {

                        double valNum = double.Parse(val);
                        t_cell.SetCellValue(valNum);
                    }
                    else
                        t_cell.SetCellValue(val);
                }
            }


            return true;
        }
        /// <summary>
        /// 写入小数
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="val"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static bool WriteCellDot(int row, int col, string val, ref HSSFSheet sheet)
        {


            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);
                if (isNumber(val))
                {
                    SetCellFormat(row, col, CellFormat.Point2, ref myExcelWork);
                    double valNum = double.Parse(val);

                    t_cell.SetCellValue(valNum);
                }
                else
                    t_cell.SetCellValue(val);
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    if (isNumber(val))
                    {
                        SetCellFormat(row, col, CellFormat.Point2, ref myExcelWork);
                        double valNum = double.Parse(val);
                        t_cell.SetCellValue(valNum);
                    }
                    else
                        t_cell.SetCellValue(val);
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    if (isNumber(val))
                    {
                        SetCellFormat(row, col, CellFormat.Point2, ref myExcelWork);
                        double valNum = double.Parse(val);
                        t_cell.SetCellValue(valNum);
                    }
                    else
                        t_cell.SetCellValue(val);
                }
            }


            return true;
        }
        /// <summary>
        /// 判断字符串是否为数字
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        private static bool isNumber(string val)
        {
            string temp = val.Trim();
            double i;
            return double.TryParse(temp,out i);
        }
        /// <summary>
        /// 向单元格写入日期
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="date"></param>
        /// <param name="wb"></param>
        public static void WriteCellDate(int row, int col, ref HSSFWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            IDataFormat dataFormat = wb.CreateDataFormat();
            ISheet sheet = wb.GetSheetAt(0);

            DateTime date = DateTime.Now;

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);

                cellStyle.DataFormat = dataFormat.GetFormat("yyyy年mm月dd日");

                t_cell.CellStyle = cellStyle;

                t_cell.SetCellValue(date);

            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.DataFormat = dataFormat.GetFormat("yyyy年mm月dd日");

                    t_cell.CellStyle = cellStyle;

                    t_cell.SetCellValue(date);
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.DataFormat = dataFormat.GetFormat("yyyy年mm月dd日");

                    t_cell.CellStyle = cellStyle;

                    t_cell.SetCellValue(date);
                }
            }
        }

        #endregion

        #region 复制单元格
        /// <summary>
        /// 将一个源文件的某个单元格的值复制到目标文件的指定单元格
        /// 以富文本的形式，保留原格式
        /// </summary>
        /// <param name="dst_row">目标行</param>
        /// <param name="dst_col">目标列</param>
        /// <param name="src_row">源行</param>
        /// <param name="src_col">源列</param>
        /// <param name="dst">目标文件</param>
        /// <param name="src">源文件</param>
        /// <returns></returns>
        public static bool CopyCell(int dst_row,int dst_col,int src_row,int src_col,ref HSSFSheet dst,ref HSSFSheet src )
        {

            if(src.GetRow(src_row-1)==null)
            {
                return false;
            }
            else
            {
                HSSFRow t_src_row = (HSSFRow)src.GetRow(src_row - 1);
                if(t_src_row.GetCell(src_col-1) == null)
                {
                    return false;
                }
                else
                {
                    HSSFCell t_src_cell = (HSSFCell)t_src_row.GetCell(src_col - 1);

                    if (dst.GetRow(dst_row - 1) == null)
                    {
                        HSSFRow t_row = (HSSFRow)dst.CreateRow(dst_row - 1);

                        HSSFCell t_cell = (HSSFCell)t_row.CreateCell(dst_col - 1);


                        t_cell.CellStyle = t_src_cell.CellStyle;
                        IRichTextString t = t_src_cell.RichStringCellValue;
                        t_cell.SetCellValue(t);
                    }
                    else
                    {
                        HSSFRow t_row = (HSSFRow)dst.GetRow(dst_row - 1);

                        if (t_row.GetCell(dst_col-1) == null)
                        {
                            HSSFCell t_cell = (HSSFCell)t_row.CreateCell(dst_col - 1);
                            t_cell.CellStyle = t_src_cell.CellStyle;
                            IRichTextString t = t_src_cell.RichStringCellValue;
                            t_cell.SetCellValue(t);
                        }
                        else
                        {
                            HSSFCell t_cell = (HSSFCell)t_row.GetCell(dst_col - 1);

                            t_cell.CellStyle = t_src_cell.CellStyle;
                            IRichTextString t = t_src_cell.RichStringCellValue;
                            t_cell.SetCellValue(t);
                        }
                    }

                }
            }


            return true;
        }
        #endregion

        #region 读取单元格
        /// <summary>
        /// 读取指定位置的单元格的值string
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static string ReadCell(int row,int col,ref HSSFSheet sheet)
        {
            if(sheet.GetRow(row-1) == null)
            {
                return "null";
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if(t_row.GetCell(col-1) == null)
                {
                    return "null";
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    return t_cell.ToString();
                }
            }
        }
        #endregion

        #region 合并单元格
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="start_row">起始行</param>
        /// <param name="end_row">结束行</param>
        /// <param name="start_col">起始列</param>
        /// <param name="end_col">结束列</param>
        /// <param name="sheet">需要设置的表格</param>
        public static void MergeCells(int start_row, int end_row, int start_col, int end_col, ref HSSFSheet sheet)
        {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(start_row-1, end_row-1, start_col-1, end_col-1));
        }
        #endregion

        #region 设置单元格的格式
        /// <summary>
        /// 设置行高，height = height*20
        /// </summary>
        /// <param name="row"></param>
        /// <param name="height"></param>
        /// <param name="sheet"></param>
        public static void SetRowHeight(int row, short height, ref HSSFSheet sheet)
        {
            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                t_row.Height = height;
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                t_row.Height = height;
            }
        }
        /// <summary>
        /// 设置列宽 width = width*256
        /// </summary>
        /// <param name="col"></param>
        /// <param name="width"></param>
        /// <param name="sheet"></param>
        public static void SetColWidth(int col, int width, ref HSSFSheet sheet)
        {
            sheet.SetColumnWidth(col - 1, width);
        }
        /// <summary>
        /// 设置单元格是否居中对齐，默认居左和底部
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="vertical">是否垂直居中</param>
        /// <param name="horizontal">是否水平居中</param>
        /// <param name="wb"></param>
        public static void SetCellAlignmentCenter(int row, int col, bool vertical, bool horizontal, ref HSSFWorkbook wb)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)wb.CreateCellStyle();

            HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);


            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);

                if (vertical) // 垂直
                {
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
                else
                {
                    cellStyle.VerticalAlignment = VerticalAlignment.Bottom;
                }
                if (horizontal) // 水平
                {
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                }
                else
                {
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                }

                t_cell.CellStyle = cellStyle;
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    if (vertical) // 垂直
                    {
                        cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    }
                    else
                    {
                        cellStyle.VerticalAlignment = VerticalAlignment.Bottom;
                    }
                    if (horizontal) // 水平
                    {
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    }
                    else
                    {
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                    }

                    t_cell.CellStyle = cellStyle;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    if (vertical) // 垂直
                    {
                        cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    }
                    else
                    {
                        cellStyle.VerticalAlignment = VerticalAlignment.Bottom;
                    }
                    if (horizontal) // 水平
                    {
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    }
                    else
                    {
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                    }

                    t_cell.CellStyle = cellStyle;
                }
            }


        }

        public enum CellFontColor
        {
            white,
            black,
            yellow,
            blue,
            green,
            red,
        }
        public enum CellFontSize
        {
            s10,
            s11,
            s12,
            s14,
            s16,
            s18,
            s20,
            s24
        }
        public enum CellFontName
        {
            SongTi,
            TimesNewRoman,
        }
        /// <summary>
        /// 设置单元格的字体的颜色和大小和字体格式
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="color">颜色</param>
        /// <param name="size">大小</param>
        /// <param name="wb"></param>
        public static void SetCellFont(int row, int col, CellFontColor color, CellFontSize size,CellFontName fontName, ref HSSFWorkbook wb)
        {
            HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);
            ICellStyle Style = wb.CreateCellStyle();
            IFont font = wb.CreateFont();
            switch (color)
            {
                case CellFontColor.black: font.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;
                    break;
                case CellFontColor.blue: font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;
                    break;
                case CellFontColor.green: font.Color = NPOI.HSSF.Util.HSSFColor.Green.Index;
                    break;
                case CellFontColor.red: font.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
                    break;
                case CellFontColor.white: font.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
                    break;
                case CellFontColor.yellow: font.Color = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                    break;
            }
            switch (size)
            {
                case CellFontSize.s10: font.FontHeightInPoints = 10;
                    break;
                case CellFontSize.s11: font.FontHeightInPoints = 11;
                    break;
                case CellFontSize.s12: font.FontHeightInPoints = 12;
                    break;
                case CellFontSize.s14: font.FontHeightInPoints = 14;
                    break;
                case CellFontSize.s16: font.FontHeightInPoints = 16;
                    break;
                case CellFontSize.s18: font.FontHeightInPoints = 18;
                    break;
                case CellFontSize.s20: font.FontHeightInPoints = 20;
                    break;
                case CellFontSize.s24: font.FontHeightInPoints = 24;
                    break;
            }

            switch (fontName)
            {
                case CellFontName.SongTi: font.FontName = "宋体"; break;
                case CellFontName.TimesNewRoman: font.FontName = "Times New Roman"; break;
            }

            // font.Boldweight = 700;  // 设置粗体

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                Style.CloneStyleFrom(t_cell.CellStyle);

                Style.SetFont(font);

                t_cell.CellStyle = Style;
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    Style.CloneStyleFrom(t_cell.CellStyle);

                    Style.SetFont(font);

                    t_cell.CellStyle = Style;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    Style.CloneStyleFrom(t_cell.CellStyle);

                    Style.SetFont(font);

                    t_cell.CellStyle = Style;
                }
            }

        }


        public enum CellFormat
        {
            Date,//日期格式
            Point2,//小数点保留两位
        }
        /// <summary>
        /// 设置单元格的格式，如日期、小数
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="format"></param>
        /// <param name="wb"></param>

        public static void SetCellFormat(int row, int col, CellFormat format, ref HSSFWorkbook wb)
        {
            HSSFCellStyle Style =(HSSFCellStyle)wb.CreateCellStyle();

            IDataFormat dataFormat = wb.CreateDataFormat();

            HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                Style.CloneStyleFrom(t_cell.CellStyle);

                switch (format)
                {
                    case CellFormat.Date: Style.DataFormat = dataFormat.GetFormat("yyyy年m月d日");
                        break;
                    case CellFormat.Point2: Style.DataFormat = dataFormat.GetFormat("0.00");
                        break;
                }

                t_cell.CellStyle = Style;
            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    Style.CloneStyleFrom(t_cell.CellStyle);

                    switch (format)
                    {
                        case CellFormat.Date: Style.DataFormat = dataFormat.GetFormat("yyyy年m月d日");
                            break;
                        case CellFormat.Point2: Style.DataFormat = dataFormat.GetFormat("0.00");
                            break;
                    }

                    t_cell.CellStyle = Style;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    Style.CloneStyleFrom(t_cell.CellStyle);

                    switch (format)
                    {
                        case CellFormat.Date: Style.DataFormat = dataFormat.GetFormat("yyyy年m月d日");
                            break;
                        case CellFormat.Point2: Style.DataFormat = dataFormat.GetFormat("0.00");
                            break;
                    }

                    t_cell.CellStyle = Style;
                }
            }

        }
        #endregion

        #region 设置边框样式
        /// <summary>
        /// 设置左边框样式
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borderStyle"></param>
        /// <param name="wb"></param>
        public static void SetCellBorderLeft(int row, int col, NPOI.SS.UserModel.BorderStyle borderStyle, ref HSSFWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            ISheet sheet = wb.GetSheetAt(0);

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);

                cellStyle.BorderLeft = borderStyle;

                t_cell.CellStyle = cellStyle;

            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderLeft = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderLeft = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
            }
        }
        /// <summary>
        /// 设置右边框样式
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borderStyle"></param>
        /// <param name="wb"></param>
        public static void SetCellBorderRight(int row, int col, NPOI.SS.UserModel.BorderStyle borderStyle, ref HSSFWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            ISheet sheet = wb.GetSheetAt(0);

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);

                cellStyle.BorderRight = borderStyle;

                t_cell.CellStyle = cellStyle;

            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderRight = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderRight = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
            }
        }
        /// <summary>
        /// 设置顶边框样式
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borderStyle"></param>
        /// <param name="wb"></param>
        public static void SetCellBorderTop(int row, int col, NPOI.SS.UserModel.BorderStyle borderStyle, ref HSSFWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            ISheet sheet = wb.GetSheetAt(0);

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);

                cellStyle.BorderTop = borderStyle;

                t_cell.CellStyle = cellStyle;

            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderTop = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderTop = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
            }
        }
        /// <summary>
        /// 设置底边框样式
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borderStyle"></param>
        /// <param name="wb"></param>
        public static void SetCellBorderBottom(int row, int col, NPOI.SS.UserModel.BorderStyle borderStyle, ref HSSFWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            ISheet sheet = wb.GetSheetAt(0);

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);

                cellStyle.BorderBottom = borderStyle;

                t_cell.CellStyle = cellStyle;

            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderBottom = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);

                    cellStyle.CloneStyleFrom(t_cell.CellStyle);

                    cellStyle.BorderBottom = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
            }
        }
        /// <summary>
        /// 设置四边边框样式
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="borderStyle"></param>
        /// <param name="wb"></param>
        public static void SetCellBorderAll(int row, int col, NPOI.SS.UserModel.BorderStyle borderStyle, ref HSSFWorkbook wb)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)wb.CreateCellStyle();

            HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);

            if (sheet.GetRow(row - 1) == null)
            {
                HSSFRow t_row = (HSSFRow)sheet.CreateRow(row - 1);

                HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);

                cellStyle.CloneStyleFrom(t_cell.CellStyle);
                cellStyle.BorderBottom = borderStyle;
                cellStyle.BorderTop = borderStyle;
                cellStyle.BorderLeft = borderStyle;
                cellStyle.BorderRight = borderStyle;

                t_cell.CellStyle = cellStyle;

            }
            else
            {
                HSSFRow t_row = (HSSFRow)sheet.GetRow(row - 1);

                if (t_row.GetCell(col - 1) == null)
                {
                    HSSFCell t_cell = (HSSFCell)t_row.CreateCell(col - 1);


                    cellStyle.CloneStyleFrom(t_cell.CellStyle);
                    cellStyle.BorderBottom = borderStyle;
                    cellStyle.BorderTop = borderStyle;
                    cellStyle.BorderLeft = borderStyle;
                    cellStyle.BorderRight = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
                else
                {
                    HSSFCell t_cell = (HSSFCell)t_row.GetCell(col - 1);


                    cellStyle.CloneStyleFrom(t_cell.CellStyle);
                    cellStyle.BorderBottom = borderStyle;
                    cellStyle.BorderTop = borderStyle;
                    cellStyle.BorderLeft = borderStyle;
                    cellStyle.BorderRight = borderStyle;

                    t_cell.CellStyle = cellStyle;
                }
            }
        }
        #endregion
		
    } // ExcelUtils
} // namespace
