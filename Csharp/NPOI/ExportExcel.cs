using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.IO;
using System.Data;

namespace ExportExcel
{
    public class CellStyle
    {
        /// <summary>
        /// 设定单元格样式

        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="myFont"></param>
        /// <param name="myCell"></param>
        /// <param name="myAlign"></param>
        /// <param name="bgcolor"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static ICellStyle GetCellStyle(IWorkbook workbook, IFont myFont, MyCell myCell, MyAlign myAlign, short bgcolor, string format)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            IDataFormat dataFormat = workbook.CreateDataFormat();

            //边框样式
            cellStyle.BorderTop = myCell.bordertop;
            cellStyle.BorderLeft = myCell.borderleft;
            cellStyle.BorderRight = myCell.borderright;
            cellStyle.BorderBottom = myCell.borderbottom;
            //边框颜色
            cellStyle.TopBorderColor = myCell.topborderColor;
            cellStyle.LeftBorderColor = myCell.leftborderColor;
            cellStyle.RightBorderColor = myCell.rightborderColor;
            cellStyle.BottomBorderColor = myCell.bottomborderColor;
            //对齐
            cellStyle.Alignment = myAlign.horizontalAlignment;
            cellStyle.VerticalAlignment = myAlign.verticalAlignment;
            //换行
            cellStyle.WrapText = myAlign.wrapText;
            //字体
            cellStyle.SetFont(myFont);
            //数据格式化

            if (format != "")
            {
                cellStyle.DataFormat = dataFormat.GetFormat(format);
            }
            //背景色

            cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            cellStyle.FillForegroundColor = bgcolor;

            return cellStyle;
        }
        /// <summary>
        /// Create font style
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="myfont"></param>
        /// <returns></returns>
        public static IFont GetFont(IWorkbook workbook, MyFont myfont)
        {
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = myfont.fontpoints;
            font.FontName = myfont.fontname;
            font.Color = myfont.fontcolor;
            font.IsItalic = myfont.isItalic;
            font.Underline = myfont.underline;
            font.IsStrikeout = myfont.isStrikeout;
            return font;
        }
    }
    /// <summary>
    /// 字体属性

    /// </summary>
    #region MyFont 字体属性

    public class MyFont
    {
        public short fontpoints { get; set; }   //字体大小
        public string fontname { get; set; }    //字体类型 宋体。。。

        public short fontcolor { get; set; }    //字体颜色 HSSFColor..index
        public bool isItalic { get; set; }      //斜体
        public byte underline { get; set; }     //下划线

        public bool isStrikeout { get; set; }   //删除线


        public MyFont()
        {
            fontpoints = (short)18;
            fontname = "微软雅黑";
            fontcolor = HSSFColor.BLACK.index;
            isItalic = false;
            underline = 0;
            isStrikeout = false;
        }

        public MyFont(short _fontpoints, string _fontname, short _fontcolor, bool _isItalic, byte _underline, bool _isStrikeout)
        {
            fontpoints = _fontpoints;
            fontname = _fontname;
            fontcolor = _fontcolor;
            isItalic = _isItalic;
            underline = _underline;
            isStrikeout = _isStrikeout;
        }
    }
    #endregion
    /// <summary>
    /// 边框样式与颜色

    /// </summary>
    #region MyCell 边框样式与颜色

    public class MyCell
    {
        //边框样式
        public BorderStyle bordertop { get; set; }
        public BorderStyle borderleft { get; set; }
        public BorderStyle borderright { get; set; }
        public BorderStyle borderbottom { get; set; }
        //边框颜色 HSSFColor..index
        public short topborderColor { get; set; }
        public short leftborderColor { get; set; }
        public short rightborderColor { get; set; }
        public short bottomborderColor { get; set; }

        public MyCell()
        {
            borderbottom = BorderStyle.NONE;
            borderleft = BorderStyle.NONE;
            borderright = BorderStyle.NONE;
            bordertop = BorderStyle.NONE;

            bottomborderColor = HSSFColor.BLACK.index;
            topborderColor = HSSFColor.BLACK.index;
            leftborderColor = HSSFColor.BLACK.index;
            rightborderColor = HSSFColor.BLACK.index;
        }
        public MyCell(BorderStyle _bordertop, BorderStyle _borderleft, BorderStyle _borderright, BorderStyle _borderbottom, short _topborderColor, short _leftborderColor, short _rightborderColor, short _bottomborderColor)
        {
            bordertop = _bordertop;
            borderleft = _borderleft;
            borderright = _borderright;
            borderbottom = _borderbottom;

            topborderColor = _topborderColor;
            leftborderColor = _leftborderColor;
            rightborderColor = _rightborderColor;
            bottomborderColor = _bottomborderColor;
        }
    }
    #endregion
    /// <summary>
    /// 排列方式，换行

    /// </summary>
    #region MyAlign 排列方式，换行

    public class MyAlign
    {
        public HorizontalAlignment horizontalAlignment { get; set; }
        public VerticalAlignment verticalAlignment { get; set; }
        public bool wrapText { get; set; }

        public MyAlign()
        {
            horizontalAlignment = HorizontalAlignment.CENTER;
            verticalAlignment = VerticalAlignment.CENTER;
            wrapText = false;
        }
        public MyAlign(HorizontalAlignment _horizontalAlignment, VerticalAlignment _verticalAlignment, bool _wrapText)
        {
            horizontalAlignment = _horizontalAlignment;
            verticalAlignment = _verticalAlignment;
            wrapText = _wrapText;
        }
    }
    #endregion
    public class ExportToExcel
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="SourceTable">导出的数据</param>
        /// <param name="Reference">导出的样式和字段</param>
        /// <param name="Sheetname">sheet名字</param>
        /// <param name="startRowCol">导出第一个数据的位置, startRowCol[0]--起始行,startRowCol[1]--起始列</param>
        /// <param name="withHeader">是否有表头</param>
        /// <returns></returns>
        public static Stream DataTableToExcel_Flexible1111(DataTable SourceTable, DataTable Reference, DataTable Parameter)
        {
            //start参数处理
            string Sheetname = Parameter.Rows[0]["sheetname"].ToString();
            string templete = System.AppDomain.CurrentDomain.BaseDirectory.ToString() + Parameter.Rows[0]["templete_path"].ToString();
            int[] startRowCol = { Convert.ToInt32(Parameter.Rows[0]["start_position_row"]), Convert.ToInt32(Parameter.Rows[0]["start_position_col"]) };
            bool withHeader = Convert.ToBoolean(Parameter.Rows[0]["withheader"]);
            bool special = Convert.ToBoolean(Parameter.Rows[0]["special"]);

            //end参数处理

            IWorkbook workbook;
            ISheet sheet;

            //如果有模板,按模板第一个sheet创建workbook.
            if (templete.Length > 0)
            {
                using (FileStream file = new FileStream(templete, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(file);
                    sheet = workbook.GetSheet(workbook.GetSheetName(0));
                }
            }
            else
            {
                workbook = new HSSFWorkbook();
                sheet = workbook.CreateSheet(Sheetname);
            }

            #region CreateStyle
            //创建字体
            //标题字体
            MyFont header = new MyFont(12, "微软雅黑", HSSFColor.WHITE.index, false, 0, false);
            IFont fontheader = CellStyle.GetFont(workbook, header);
            //内容字体
            MyFont body = new MyFont(10, "微软雅黑", HSSFColor.BLACK.index, false, 0, false);
            IFont fontbody = CellStyle.GetFont(workbook, body);

            //创建边框颜色

            //无边框无颜色
            MyCell noborder = new MyCell();
            //黑色细边框
            MyCell slimborder = new MyCell(BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN,
                                HSSFColor.BLACK.index, HSSFColor.BLACK.index, HSSFColor.BLACK.index, HSSFColor.BLACK.index);
            //黑色粗边框
            MyCell thickborder = new MyCell(BorderStyle.THICK, BorderStyle.THICK, BorderStyle.THICK, BorderStyle.THICK,
                                HSSFColor.BLACK.index, HSSFColor.BLACK.index, HSSFColor.BLACK.index, HSSFColor.BLACK.index);

            //设定对齐方式

            //默认居中不换行
            MyAlign aligncenter = new MyAlign();
            //左对齐不换行
            MyAlign alignleft = new MyAlign(HorizontalAlignment.LEFT, VerticalAlignment.CENTER, false);
            //右对齐不换行
            MyAlign alignright = new MyAlign(HorizontalAlignment.RIGHT, VerticalAlignment.CENTER, false);


            //最终样式

            #region 标题样式
            //标题样式
            ICellStyle[] cellStyleHeader1 = new ICellStyle[10];
            cellStyleHeader1[1] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.SEA_GREEN.index, "");
            cellStyleHeader1[2] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLUE.index, "");
            cellStyleHeader1[3] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.ORANGE.index, "");
            cellStyleHeader1[4] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLACK.index, "");

            ICellStyle[] cellStyleHeader2 = new ICellStyle[10];
            cellStyleHeader2[1] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.SEA_GREEN.index, "");
            cellStyleHeader2[2] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLUE.index, "");
            cellStyleHeader2[3] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.ORANGE.index, "");
            cellStyleHeader2[4] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLACK.index, "");

            ICellStyle[] cellStyleHeader3 = new ICellStyle[10];
            cellStyleHeader3[1] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.SEA_GREEN.index, "");
            cellStyleHeader3[2] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLUE.index, "");
            cellStyleHeader3[3] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.ORANGE.index, "");
            cellStyleHeader3[4] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLACK.index, "");

            ICellStyle[] cellStyleHeader4 = new ICellStyle[10];
            cellStyleHeader4[1] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.SEA_GREEN.index, "");
            cellStyleHeader4[2] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLUE.index, "");
            cellStyleHeader4[3] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.ORANGE.index, "");
            cellStyleHeader4[4] = CellStyle.GetCellStyle(workbook, fontheader, slimborder, aligncenter, HSSFColor.BLACK.index, "");
            #endregion

            #region 文本样式
            //1.正常文本
            ICellStyle[] cellStyle1 = new ICellStyle[10];
            cellStyle1[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "");
            cellStyle1[1] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignleft, HSSFColor.WHITE.index, "");
            cellStyle1[2] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "");
            cellStyle1[3] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "");
            cellStyle1[4] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "");
            cellStyle1[5] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "");

            //2.日期
            ICellStyle[] cellStyle2 = new ICellStyle[10];
            cellStyle2[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "yyyy/mm/dd");
            cellStyle2[1] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm/dd/yyyy");
            cellStyle2[2] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm/dd");
            cellStyle2[3] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "yyyy-mm-dd");
            cellStyle2[4] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm-dd-yyyy");
            cellStyle2[5] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm-dd");
            cellStyle2[6] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "yyyy年m月d日");
            cellStyle2[7] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "yyyy年mm月dd日");

            //3.时间
            ICellStyle[] cellStyle3 = new ICellStyle[10];
            cellStyle3[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "yyyy/mm/dd hh:mm:ss");
            cellStyle3[1] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm/dd/yyyy hh:mm:ss");
            cellStyle3[2] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm/dd hh:mm:ss");
            cellStyle3[3] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "yyyy-mm-dd hh:mm:ss");
            cellStyle3[4] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm-dd-yyyy hh:mm:ss");
            cellStyle3[5] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "mm-dd hh:mm:ss");

            //4.货币
            ICellStyle[] cellStyle4 = new ICellStyle[10];
            cellStyle4[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "￥#,##0");
            cellStyle4[1] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "[DbNum2][$-804]0元");
            cellStyle4[2] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "￥#,##0");
            cellStyle4[3] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "￥#,##0");

            //5.百分比
            ICellStyle[] cellStyle5 = new ICellStyle[10];
            cellStyle5[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0%");
            cellStyle5[1] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.0%");
            cellStyle5[2] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.00%");
            cellStyle5[3] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.000%");
            cellStyle5[4] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.0000%");
            cellStyle5[5] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.00000%");

            //6.数字
            ICellStyle[] cellStyle6 = new ICellStyle[10];
            cellStyle6[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0");
            cellStyle6[1] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.0");
            cellStyle6[2] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.00");
            cellStyle6[3] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.000");
            cellStyle6[4] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.0000");
            cellStyle6[5] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.00000");
            cellStyle6[6] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, alignright, HSSFColor.WHITE.index, "0.000000");

            //7.科学计数法
            ICellStyle[] cellStyle7 = new ICellStyle[10];
            cellStyle7[0] = CellStyle.GetCellStyle(workbook, fontbody, slimborder, aligncenter, HSSFColor.WHITE.index, "0.00E+00");

            #endregion
            #endregion

            MemoryStream ms = new MemoryStream();

            if (special)
            {
                //第一行固定时间--特别要求--写死
                IRow headerRow0 = sheet.CreateRow(0);
                ICell cell0 = headerRow0.CreateCell(0);
                cell0.SetCellValue(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                ICell cell1 = headerRow0.CreateCell(1);
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 1));
            }

            #region Header

            if (withHeader)
            {
                IRow headerRow = sheet.CreateRow(startRowCol[0]);
                //表头
                for (int i = 0; i < SourceTable.Columns.Count; i++)
                {
                    string style = string.Empty;
                    int seq = 0;
                    ICell cell = headerRow.CreateCell(startRowCol[1] + i);
                    DataRow[] drColumns = Reference.Select("export_column ='" + SourceTable.Columns[i].ColumnName.ToLower() + "'");

                    if (drColumns.Length > 0)
                    {
                        cell.SetCellValue(drColumns[0]["column_display"].ToString());
                        style = drColumns[0]["header_style"].ToString().ToLower();
                        seq = Convert.ToInt32(style.Substring(style.Length - 1, 1));
                    }
                    else
                    {
                        cell.SetCellValue(SourceTable.Columns[i].ColumnName);
                    }

                    if (style.Contains("cellstyleheader1"))
                    {
                        cell.CellStyle = cellStyleHeader1[seq];
                    }
                    else if (style.Contains("cellstyleheader2"))
                    {
                        cell.CellStyle = cellStyleHeader2[seq];
                    }
                    else if (style.Contains("cellstyleheader3"))
                    {
                        cell.CellStyle = cellStyleHeader3[seq];
                    }
                    else if (style.Contains("cellstyleheader4"))
                    {
                        cell.CellStyle = cellStyleHeader4[seq];
                    }
                    else
                    {
                        cell.CellStyle = cellStyleHeader1[0];
                    }
                }
                headerRow = null;
            }
            #endregion

            #region body
            //数据
            for (int i = 0; i < SourceTable.Rows.Count; i++)
            {
                IRow row;
                if (withHeader)
                {
                    row = sheet.CreateRow(startRowCol[0] + i + 1);
                }
                else
                {
                    row = sheet.CreateRow(startRowCol[0] + i);
                }

                #region For
                for (int j = 0; j < SourceTable.Columns.Count; j++)
                {
                    string style = string.Empty;
                    int seq = 0;
                    int pos = startRowCol[1] + j;
                    Type dataType = SourceTable.Rows[i][j].GetType();
                    ICell cell = row.CreateCell(pos);
                    DataRow[] drColumns = Reference.Select("export_column ='" + SourceTable.Columns[j].ColumnName.ToLower() + "'");


                    if (drColumns.Length > 0)
                    {
                        style = drColumns[0]["column_style"].ToString().ToLower();
                        seq = Convert.ToInt32(style.Substring(style.Length - 1, 1));
                    }

                    if (style.Contains("cellstyle1") && style.Length > 0)
                    {
                        cell.CellStyle = cellStyle1[seq];
                        cell.SetCellValue(SourceTable.Rows[i][j].ToString());
                    }
                    else if (style.Contains("cellstyle2") && style.Length > 0 && dataType == typeof(DateTime))
                    {
                        cell.CellStyle = cellStyle2[seq];
                        cell.SetCellValue(Convert.ToDateTime(SourceTable.Rows[i][j]));
                    }
                    else if (style.Contains("cellstyle3") && style.Length > 0 && dataType == typeof(DateTime))
                    {
                        cell.CellStyle = cellStyle3[seq];
                        cell.SetCellValue(Convert.ToDateTime(SourceTable.Rows[i][j]));
                    }
                    else if (style.Contains("cellstyle4") && style.Length > 0 && (dataType == typeof(double) || dataType == typeof(decimal)) || dataType == typeof(int))
                    {
                        cell.CellStyle = cellStyle4[seq];
                        cell.SetCellValue(Convert.ToDouble(SourceTable.Rows[i][j]));
                    }
                    else if (style.Contains("cellstyle5") && style.Length > 0 && (dataType == typeof(double) || dataType == typeof(decimal)) || dataType == typeof(int))
                    {
                        cell.CellStyle = cellStyle5[seq];
                        cell.SetCellValue(Convert.ToDouble(SourceTable.Rows[i][j]));
                    }
                    else if (style.Contains("cellstyle6") && style.Length > 0 && (dataType == typeof(double) || dataType == typeof(decimal)) || dataType == typeof(int))
                    {
                        cell.CellStyle = cellStyle6[seq];
                        cell.SetCellValue(Convert.ToDouble(SourceTable.Rows[i][j]));
                    }
                    else if (style.Contains("cellstyle7") && style.Length > 0 && (dataType == typeof(double) || dataType == typeof(decimal)) || dataType == typeof(int))
                    {
                        cell.CellStyle = cellStyle7[seq];
                        cell.SetCellValue(Convert.ToInt32(SourceTable.Rows[i][j]));
                    }
                    else
                    {
                        cell.CellStyle = cellStyle1[0];
                        cell.SetCellValue(SourceTable.Rows[i][j].ToString());
                    }
                }
                #endregion
            }
            #endregion

            for (int i = 0; i < SourceTable.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            if (withHeader)
            {
                sheet.CreateFreezePane(startRowCol[1], startRowCol[0] + 1);
            }
            else//没有表头
            {
                sheet.CreateFreezePane(startRowCol[1], startRowCol[0]);
            }

            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            sheet = null;
            workbook = null;
            return ms;
        }
    }
}
