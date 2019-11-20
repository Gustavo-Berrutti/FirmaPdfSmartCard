using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FirmaPdfWpf.Npoi
{
    public abstract class ExcelGeneratorBase
    {
        public enum ExcelFormat
        {
            xls,
            xlsx
        }

        private const int WIDTH_MULTIPLIER = 256;
        protected const short SUB_TOTAL_BACK_COLOR_INDEX = 57;
        protected const short ALTERNATE_BACK_COLOR_INDEX = 58;
        protected const short DARK_BLUE_EXCEL_2007 = 59;

        protected Dictionary<short, IColor> customColors = new Dictionary<short, IColor>();

        protected HSSFPalette palette;

        protected IWorkbook wb;
        protected IFont bold16, bold11, bold12, bold14, bold11Italic, boldUnderlined11, normal11;
        protected ICellStyle wrapTextStyle, titleStyle, bold11RightStyle, bold11LeftStyle, rowSeparatorStyle, columnHeaderStyleUnderlined, columnHeaderStyle, tableTitleStyle, tableDateTitleStyle, textStyle, currencyStyle, currencyRightLineStyle, alternateTextStyle, alternateCurrencyStyle, alternateCurrencyRightLineStyle, intStyle, alternateIntStyle, intStyleWithoutThousandSeparator, alternateIntStyleWithoutThousandSeparator,
           dateStyle, dateTimeStyle, alternateDateTimeStyle, dateTimeAmPmStyle, alternateDateTimeAmPmStyle, alternateDateStyle, percentageStyle, alternatePercentageStyle, fullDateStyle, alternateFullDateStyle, subTotalTextStyle, subTotalIntStyle, subTotalCurrencyStyle, grandTotalTextStyle, grandTotalIntStyle, grandTotalCurrencyStyle, grandTotalCurrencyRightLineStyle, grandTotalPercentageStyle;

        protected IDataFormat format;

        protected int widthInChars(int cantChars)
        {
            return cantChars * WIDTH_MULTIPLIER;
        }

        protected void CreateWorkbook(String subject)
        {
            CreateWorkbook(subject, ExcelFormat.xls);
        }

        protected void CreateWorkbook(String subject, ExcelFormat excelFormat)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-UY");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("es-UY");
            if (excelFormat == ExcelFormat.xls)
                wb = new HSSFWorkbook();
            else
                wb = new XSSFWorkbook();


            AddDocumentInformation(wb, subject, excelFormat);
            LoadCustomColors(wb, excelFormat);

        }

        private void AddDocumentInformation(IWorkbook wb, string subject, ExcelFormat excelFormat)
        {
            if (excelFormat == ExcelFormat.xls)
            {
                var wbHssf = wb as HSSFWorkbook;

                
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "BSE";

                wbHssf.DocumentSummaryInformation = dsi;

                
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Subject = subject;
                si.Author = "BSE";
                wbHssf.SummaryInformation = si;
            }
            else
            {
                var wbXssf = wb as NPOI.XSSF.UserModel.XSSFWorkbook;
                var xmlProps = wbXssf.GetProperties();
                var coreProps = xmlProps.CoreProperties;
                coreProps.Subject = subject;
                coreProps.Creator = "BSE";
                var appProps = xmlProps.ExtendedProperties.GetUnderlyingProperties();
                appProps.Company = "BSE";
            }
        }

        private void LoadCustomColors(IWorkbook wb, ExcelFormat format)
        {
            if (format == ExcelFormat.xls)
            {
                var wbHssf = wb as HSSFWorkbook;
                palette = wbHssf.GetCustomPalette();
                palette.SetColorAtIndex(SUB_TOTAL_BACK_COLOR_INDEX, (byte)239, (byte)240, (byte)241);
                palette.SetColorAtIndex(ALTERNATE_BACK_COLOR_INDEX, (byte)210, (byte)210, (byte)210);
                palette.SetColorAtIndex(DARK_BLUE_EXCEL_2007, (byte)31, (byte)73, (byte)125);
                customColors.Add(SUB_TOTAL_BACK_COLOR_INDEX, palette.GetColor(SUB_TOTAL_BACK_COLOR_INDEX));
                customColors.Add(ALTERNATE_BACK_COLOR_INDEX, palette.GetColor(ALTERNATE_BACK_COLOR_INDEX));
                customColors.Add(DARK_BLUE_EXCEL_2007, palette.GetColor(DARK_BLUE_EXCEL_2007));
            }
            else
            {
                customColors.Add(SUB_TOTAL_BACK_COLOR_INDEX, new XSSFColor(System.Drawing.Color.FromArgb(239, 240, 241)));
                customColors.Add(ALTERNATE_BACK_COLOR_INDEX, new XSSFColor(System.Drawing.Color.FromArgb(210, 210, 210)));
                customColors.Add(DARK_BLUE_EXCEL_2007, new XSSFColor(System.Drawing.Color.FromArgb(31, 73, 125)));
            }


        }

        protected void setBackgroundCustomColor(ICellStyle style, IColor customColor)
        {
            if (customColor is XSSFColor)
            {
                var xssfStyle = style as XSSFCellStyle;
                xssfStyle.SetFillForegroundColor((XSSFColor)customColor);
            }
            else
                style.FillForegroundColor = customColor.Indexed;
        }

        protected void InitFontsStyles()
        {
            bold16 = wb.CreateFont();
            bold16.Color = IndexedColors.BlueGrey.Index;
            bold16.FontHeightInPoints = 16;
            bold16.Boldweight = (short)FontBoldWeight.Bold;

            bold11 = wb.CreateFont();
            bold11.FontHeightInPoints = 11;
            bold11.Boldweight = (short)FontBoldWeight.Bold;

            bold12 = wb.CreateFont();
            bold12.FontHeightInPoints = 12;
            bold12.Boldweight = (short)FontBoldWeight.Bold;

            bold14 = wb.CreateFont();
            bold14.FontHeightInPoints = 14;
            bold14.Boldweight = (short)FontBoldWeight.Bold;

            bold11Italic = wb.CreateFont();
            bold11Italic.FontHeightInPoints = 11;
            bold11Italic.Boldweight = (short)FontBoldWeight.Bold;
            bold11Italic.IsItalic = true;

            boldUnderlined11 = wb.CreateFont();
            boldUnderlined11.FontHeightInPoints = 11;
            boldUnderlined11.Underline = FontUnderlineType.Single;
            boldUnderlined11.Boldweight = (short)FontBoldWeight.Bold;

            normal11 = wb.CreateFont();
            normal11.FontHeightInPoints = 11;
            normal11.Color = IndexedColors.Black.Index;
            normal11.FontName = "Times New Roman";

            format = wb.CreateDataFormat();

            titleStyle = wb.CreateCellStyle();
            titleStyle.SetFont(bold16);
            titleStyle.Alignment = HorizontalAlignment.Center;

            wrapTextStyle = wb.CreateCellStyle();
            wrapTextStyle.WrapText = true;

            tableTitleStyle = wb.CreateCellStyle();
            tableTitleStyle.SetFont(bold14);

            tableDateTitleStyle = wb.CreateCellStyle();
            tableDateTitleStyle.SetFont(bold12);
            tableDateTitleStyle.DataFormat = format.GetFormat("dddd, MMM. dd, yyyy");
            tableDateTitleStyle.Alignment = HorizontalAlignment.Left;


            columnHeaderStyleUnderlined = wb.CreateCellStyle();
            columnHeaderStyleUnderlined.SetFont(bold11);
            columnHeaderStyleUnderlined.Alignment = HorizontalAlignment.Center;
            columnHeaderStyleUnderlined.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            columnHeaderStyleUnderlined.FillPattern = FillPattern.SolidForeground;
            columnHeaderStyleUnderlined.BorderBottom = BorderStyle.Medium;
            columnHeaderStyleUnderlined.WrapText = true;

            columnHeaderStyle = wb.CreateCellStyle();
            columnHeaderStyle.SetFont(bold11);
            columnHeaderStyle.Alignment = HorizontalAlignment.Center;
            columnHeaderStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            columnHeaderStyle.FillPattern = FillPattern.SolidForeground;
            columnHeaderStyle.WrapText = true;



            bold11RightStyle = wb.CreateCellStyle();
            bold11RightStyle.SetFont(bold11);
            bold11RightStyle.Alignment = HorizontalAlignment.Right;

            bold11LeftStyle = wb.CreateCellStyle();
            bold11LeftStyle.SetFont(bold11);
            bold11LeftStyle.Alignment = HorizontalAlignment.Left;

            textStyle = wb.CreateCellStyle();
            textStyle.SetFont(normal11);
            textStyle.VerticalAlignment = VerticalAlignment.Center;

            currencyStyle = wb.CreateCellStyle();
            currencyStyle.SetFont(normal11);
            currencyStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");
            currencyStyle.Alignment = HorizontalAlignment.Right;



            currencyRightLineStyle = wb.CreateCellStyle();
            currencyRightLineStyle.SetFont(normal11);
            currencyRightLineStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");
            currencyRightLineStyle.Alignment = HorizontalAlignment.Right;
            currencyRightLineStyle.BorderRight = BorderStyle.Medium;

            alternateTextStyle = wb.CreateCellStyle();
            alternateTextStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateTextStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateTextStyle.FillPattern = FillPattern.SolidForeground;

            alternateCurrencyStyle = wb.CreateCellStyle();
            alternateCurrencyStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateCurrencyStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateCurrencyStyle.FillPattern = FillPattern.SolidForeground;
            alternateCurrencyStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");
            alternateCurrencyStyle.Alignment = HorizontalAlignment.Right;

            alternateCurrencyRightLineStyle = wb.CreateCellStyle();
            alternateCurrencyRightLineStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateCurrencyRightLineStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateCurrencyRightLineStyle.FillPattern = FillPattern.SolidForeground;
            alternateCurrencyRightLineStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");
            alternateCurrencyRightLineStyle.Alignment = HorizontalAlignment.Right;
            alternateCurrencyRightLineStyle.BorderRight = BorderStyle.Medium;

            intStyle = wb.CreateCellStyle();
            intStyle.SetFont(normal11);
            intStyle.Alignment = HorizontalAlignment.Right;
            intStyle.DataFormat = format.GetFormat("#,##0");

            alternateIntStyle = wb.CreateCellStyle();
            alternateIntStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateIntStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateIntStyle.FillPattern = FillPattern.SolidForeground;
            alternateIntStyle.Alignment = HorizontalAlignment.Right;
            alternateIntStyle.DataFormat = format.GetFormat("#,##0");

            intStyleWithoutThousandSeparator = wb.CreateCellStyle();
            intStyleWithoutThousandSeparator.SetFont(normal11);
            intStyleWithoutThousandSeparator.Alignment = HorizontalAlignment.Right;
            intStyleWithoutThousandSeparator.DataFormat = format.GetFormat("#0");

            alternateIntStyleWithoutThousandSeparator = wb.CreateCellStyle();
            alternateIntStyleWithoutThousandSeparator.SetFont(normal11);
            setBackgroundCustomColor(alternateIntStyleWithoutThousandSeparator, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateIntStyleWithoutThousandSeparator.FillPattern = FillPattern.SolidForeground;
            alternateIntStyleWithoutThousandSeparator.Alignment = HorizontalAlignment.Right;
            alternateIntStyleWithoutThousandSeparator.DataFormat = format.GetFormat("#0");




            percentageStyle = wb.CreateCellStyle();
            percentageStyle.SetFont(normal11);
            percentageStyle.DataFormat = format.GetFormat("#,##0.00%");
            percentageStyle.Alignment = HorizontalAlignment.Right;

            alternatePercentageStyle = wb.CreateCellStyle();
            alternatePercentageStyle.SetFont(normal11);
            setBackgroundCustomColor(alternatePercentageStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternatePercentageStyle.FillPattern = FillPattern.SolidForeground;
            alternatePercentageStyle.DataFormat = format.GetFormat("#,##0.00%");
            alternatePercentageStyle.Alignment = HorizontalAlignment.Right;


            dateStyle = wb.CreateCellStyle();
            dateStyle.SetFont(normal11);
            dateStyle.Alignment = HorizontalAlignment.Center;
            dateStyle.DataFormat = format.GetFormat("dd/MM/yyyy");

            alternateDateStyle = wb.CreateCellStyle();
            alternateDateStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateDateStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateDateStyle.FillPattern = FillPattern.SolidForeground;
            alternateDateStyle.Alignment = HorizontalAlignment.Center;
            alternateDateStyle.DataFormat = format.GetFormat("dd/MM/yyyy");


            dateTimeStyle = wb.CreateCellStyle();
            dateTimeStyle.SetFont(normal11);
            dateTimeStyle.Alignment = HorizontalAlignment.Center;
            dateTimeStyle.DataFormat = format.GetFormat("dd/MM/yyyy HH:mm");

            alternateDateTimeStyle = wb.CreateCellStyle();
            alternateDateTimeStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateDateTimeStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateDateTimeStyle.FillPattern = FillPattern.SolidForeground;
            alternateDateTimeStyle.Alignment = HorizontalAlignment.Center;
            alternateDateTimeStyle.DataFormat = format.GetFormat("dd/MM/yyyy HH:mm");

            dateTimeAmPmStyle = wb.CreateCellStyle();
            dateTimeAmPmStyle.SetFont(normal11);
            dateTimeAmPmStyle.Alignment = HorizontalAlignment.Center;
            dateTimeAmPmStyle.DataFormat = format.GetFormat("dd/MM/yyyy hh:mm AM/PM");

            alternateDateTimeAmPmStyle = wb.CreateCellStyle();
            alternateDateTimeAmPmStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateDateTimeAmPmStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateDateTimeAmPmStyle.FillPattern = FillPattern.SolidForeground;
            alternateDateTimeAmPmStyle.Alignment = HorizontalAlignment.Center;
            alternateDateTimeAmPmStyle.DataFormat = format.GetFormat("dd/MM/yyyy HH:mm AM/PM");

            fullDateStyle = wb.CreateCellStyle();
            fullDateStyle.SetFont(normal11);
            fullDateStyle.Alignment = HorizontalAlignment.Center;
            fullDateStyle.DataFormat = format.GetFormat("dd-MMM-yyyy h:mm AM/PM");

            alternateFullDateStyle = wb.CreateCellStyle();
            alternateFullDateStyle.SetFont(normal11);
            setBackgroundCustomColor(alternateFullDateStyle, customColors[ALTERNATE_BACK_COLOR_INDEX]);
            alternateFullDateStyle.FillPattern = FillPattern.SolidForeground;
            alternateFullDateStyle.Alignment = HorizontalAlignment.Center;
            alternateFullDateStyle.DataFormat = format.GetFormat("dd-MMM-yyyy h:mm AM/PM");

            subTotalTextStyle = wb.CreateCellStyle();
            subTotalTextStyle.SetFont(bold11);
            subTotalTextStyle.Alignment = HorizontalAlignment.Center;
            subTotalTextStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            subTotalTextStyle.FillPattern = FillPattern.SolidForeground;
            subTotalTextStyle.BorderBottom = BorderStyle.Medium;
            subTotalTextStyle.WrapText = true;
            setBackgroundCustomColor(subTotalTextStyle, customColors[SUB_TOTAL_BACK_COLOR_INDEX]);
            subTotalTextStyle.BorderTop = BorderStyle.Medium;
            subTotalTextStyle.Alignment = HorizontalAlignment.Right;

            subTotalIntStyle = wb.CreateCellStyle();
            subTotalIntStyle.SetFont(bold11);
            subTotalIntStyle.Alignment = HorizontalAlignment.Center;
            subTotalIntStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            subTotalIntStyle.FillPattern = FillPattern.SolidForeground;
            subTotalIntStyle.BorderBottom = BorderStyle.Medium;
            subTotalIntStyle.WrapText = true;
            setBackgroundCustomColor(subTotalIntStyle, customColors[SUB_TOTAL_BACK_COLOR_INDEX]);
            subTotalIntStyle.BorderTop = BorderStyle.Medium;
            subTotalIntStyle.Alignment = HorizontalAlignment.Right;
            subTotalIntStyle.DataFormat = format.GetFormat("#,##0");

            subTotalCurrencyStyle = wb.CreateCellStyle();
            subTotalCurrencyStyle.SetFont(bold11);
            subTotalCurrencyStyle.Alignment = HorizontalAlignment.Center;
            subTotalCurrencyStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            subTotalCurrencyStyle.FillPattern = FillPattern.SolidForeground;
            subTotalCurrencyStyle.BorderBottom = BorderStyle.Medium;
            subTotalCurrencyStyle.WrapText = true;
            setBackgroundCustomColor(subTotalCurrencyStyle, customColors[SUB_TOTAL_BACK_COLOR_INDEX]);
            subTotalCurrencyStyle.BorderTop = BorderStyle.Medium;
            subTotalCurrencyStyle.Alignment = HorizontalAlignment.Right;
            subTotalCurrencyStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");

            grandTotalTextStyle = wb.CreateCellStyle();
            grandTotalTextStyle.SetFont(bold11);
            grandTotalTextStyle.Alignment = HorizontalAlignment.Center;
            grandTotalTextStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            grandTotalTextStyle.FillPattern = FillPattern.SolidForeground;
            grandTotalTextStyle.BorderBottom = BorderStyle.Medium;
            grandTotalTextStyle.WrapText = true;
            grandTotalTextStyle.BorderTop = BorderStyle.Medium;
            grandTotalTextStyle.Alignment = HorizontalAlignment.Right;


            grandTotalIntStyle = wb.CreateCellStyle();
            grandTotalIntStyle.SetFont(bold11);
            grandTotalIntStyle.Alignment = HorizontalAlignment.Center;
            grandTotalIntStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            grandTotalIntStyle.FillPattern = FillPattern.SolidForeground;
            grandTotalIntStyle.BorderBottom = BorderStyle.Medium;
            grandTotalIntStyle.WrapText = true;
            grandTotalIntStyle.BorderTop = BorderStyle.Medium;
            grandTotalIntStyle.Alignment = HorizontalAlignment.Right;
            grandTotalIntStyle.DataFormat = format.GetFormat("#,##0");

            grandTotalCurrencyStyle = wb.CreateCellStyle();
            grandTotalCurrencyStyle.SetFont(bold11);
            grandTotalCurrencyStyle.Alignment = HorizontalAlignment.Center;
            grandTotalCurrencyStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            grandTotalCurrencyStyle.FillPattern = FillPattern.SolidForeground;
            grandTotalCurrencyStyle.BorderBottom = BorderStyle.Medium;
            grandTotalCurrencyStyle.WrapText = true;
            grandTotalCurrencyStyle.BorderTop = BorderStyle.Medium;
            grandTotalCurrencyStyle.Alignment = HorizontalAlignment.Right;
            grandTotalCurrencyStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");

            grandTotalCurrencyRightLineStyle = wb.CreateCellStyle();
            grandTotalCurrencyRightLineStyle.SetFont(bold11);
            grandTotalCurrencyRightLineStyle.Alignment = HorizontalAlignment.Center;
            grandTotalCurrencyRightLineStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            grandTotalCurrencyRightLineStyle.FillPattern = FillPattern.SolidForeground;
            grandTotalCurrencyRightLineStyle.BorderBottom = BorderStyle.Medium;
            grandTotalCurrencyRightLineStyle.WrapText = true;
            grandTotalCurrencyRightLineStyle.BorderTop = BorderStyle.Medium;
            grandTotalCurrencyRightLineStyle.Alignment = HorizontalAlignment.Right;
            grandTotalCurrencyRightLineStyle.DataFormat = format.GetFormat("$ #,##0.00;$ -#,##0.00");
            grandTotalCurrencyRightLineStyle.BorderRight = BorderStyle.Medium;

            grandTotalPercentageStyle = wb.CreateCellStyle();
            grandTotalPercentageStyle.SetFont(bold11);
            grandTotalPercentageStyle.Alignment = HorizontalAlignment.Center;
            grandTotalPercentageStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            grandTotalPercentageStyle.FillPattern = FillPattern.SolidForeground;
            grandTotalPercentageStyle.BorderBottom = BorderStyle.Medium;
            grandTotalPercentageStyle.WrapText = true;
            grandTotalPercentageStyle.BorderTop = BorderStyle.Medium;
            grandTotalPercentageStyle.Alignment = HorizontalAlignment.Right;
            grandTotalPercentageStyle.DataFormat = format.GetFormat("#,##0.00%");


            rowSeparatorStyle = wb.CreateCellStyle();
            rowSeparatorStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            rowSeparatorStyle.FillPattern = FillPattern.ThinForwardDiagonals; ;
        }

        protected void AddHeaderCell(IRow row, int col, String text, int width = -1, bool underlined = true, int colSpan = 1, bool borderRight = false)
        {
            var cell = row.CreateCell(col);
            cell.SetCellValue(text);
            cell.CellStyle = underlined ? columnHeaderStyleUnderlined : columnHeaderStyle;
            if (width != -1)
                row.Sheet.SetColumnWidth(col, width);
            var range = new CellRangeAddress(row.RowNum, row.RowNum, col, col + colSpan - 1);
            if (colSpan > 1)
                row.Sheet.AddMergedRegion(range);
            if (borderRight)
                RegionUtil.SetBorderRight((int)BorderStyle.Medium, range, row.Sheet, row.Sheet.Workbook);
        }


        protected void AddTitleCell(IRow row, int col, int colSpan, string text)
        {
            var cell = row.CreateCell(col);

            cell.CellStyle = titleStyle;
            var rowIndex = row.RowNum;
            CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, col, col + colSpan - 1);
            row.Sheet.AddMergedRegion(region);
            cell.SetCellValue(text);
        }

        protected MemoryStream WriteToStream()
        {
            //Write the stream data of workbook
            using (MemoryStream file = new MemoryStream())
            {

                wb.Write(file);
                return file;
            }
        }

        protected byte[] WriteToArray()
        {
            using (MemoryStream file = new MemoryStream())
            {
                wb.Write(file);
                return file.ToArray();
            }
        }


        
        protected void copyTexttOfWrappedRegionInLastColumn(string cellText, int rowNum, int rangeWidth, ISheet sheet)
        {
            var row = sheet.GetRow(rowNum);
            var cell = row.CreateCell(255);
            sheet.SetColumnWidth(255, rangeWidth);
            cell.SetCellValue(cellText);
            cell.CellStyle = wrapTextStyle;
        }
    }
}
