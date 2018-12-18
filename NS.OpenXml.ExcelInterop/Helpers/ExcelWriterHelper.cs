/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.excelInterop - All Rights Reserved */
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;

namespace NS.OpenXml.ExcelInterop.Helpers
{
    /// <summary>
    /// The excel writer helper that allows to writes excel file in Open Xml format 
    /// </summary>
    internal class ExcelWriterHelper
    {
        #region internal methods

        /// <summary>
        /// Creates column data
        /// </summary>
        /// <param name="startColumnIndex">the start column index</param>
        /// <param name="endColumnIndex">the end column index</param>
        /// <param name="columnWidth">the column width</param>
        /// <returns>The column</returns>
        internal static Column CreateColumnData(uint startColumnIndex, uint endColumnIndex, double columnWidth)
        {
            return new Column
            {
                Min = startColumnIndex,
                Max = endColumnIndex,
                Width = columnWidth,
                CustomWidth = true
            };
        }

        /// <summary>
        /// Creates header row
        /// </summary>
        /// <param name="index">The index</param>
        /// <param name="dt">the data table</param>
        /// <returns>The row</returns>
        internal static Row CreateHeaderRow(uint index, DataTable dt)
        {
            Row row = new Row { RowIndex = index };

            foreach (DataColumn col in dt.Columns)
            {
                Cell cell = CreateTextCell(col.ColumnName);
                cell.StyleIndex = 10;
                row.Append(cell);
            }

            return row;
        }

        /// <summary>
        /// Creates footer row
        /// </summary>
        /// <param name="index">The index</param>
        /// <param name="dr">The data row</param>
        /// <returns>The row</returns>
        internal static Row CreateFooterRow(uint index, DataRow dr)
        {
            Row row = new Row { RowIndex = index };

            foreach (object itm in dr.ItemArray)
            {
                Cell cell;

                if (itm.GetType() == Type.GetType("System.Int32"))
                    cell = CreateNumberCell(Convert.ToInt32(itm));
                else if (itm.GetType() == Type.GetType("System.Decimal"))
                    cell = CreateDecimalCell(Convert.ToDecimal(itm));
                else
                {
                    cell = new Cell { DataType = CellValues.InlineString };
                    InlineString istring = new InlineString();
                    Text t = new Text { Text = itm.ToString() };
                    istring.Append(t);
                    cell.StyleIndex = 12;

                    cell.Append(istring);
                }

                row.Append(cell);
            }

            return row;
        }

        /// <summary>
        /// Creates content row
        /// </summary>
        /// <param name="index">The index</param>
        /// <param name="dr">The data row</param>
        /// <returns>The row</returns>
        internal static Row CreateContentRow(uint index, DataRow dr)
        {
            Row row = new Row { RowIndex = index };

            // First element
            Cell cellF;
            if (dr.ItemArray[0].GetType() == Type.GetType("System.Int32"))
                cellF = CreateNumberCell(Convert.ToInt32(dr.ItemArray[0]));
            else if (dr.ItemArray[0].GetType() == Type.GetType("System.Decimal"))
                cellF = CreateDecimalCell(Convert.ToDecimal(dr.ItemArray[0]));
            else
                cellF = CreateTextCell(dr.ItemArray[0].ToString());
            row.Append(cellF);

            for (int i = 1; i < dr.ItemArray.Length; i++)
            {
                Cell cell;
                if (dr.ItemArray[i].GetType() == Type.GetType("System.Int32"))
                    cell = CreateNumberCell(Convert.ToInt32(dr.ItemArray[i]));
                else if (dr.ItemArray[i].GetType() == Type.GetType("System.Decimal"))
                    cell = CreateDecimalCell(Convert.ToDecimal(dr.ItemArray[i]));
                else
                    cell = CreateAlignedTextCell(dr.ItemArray[i].ToString());
                row.Append(cell);
            }

            return row;
        }

        /// <summary>
        /// Creates text cell
        /// </summary>
        /// <param name="text">the text content</param>
        /// <returns>The cell</returns>
        internal static Cell CreateTextCell(string text)
        {
            Cell cell = new Cell { DataType = CellValues.InlineString };
            InlineString istring = new InlineString();
            istring.Append(new Text { Text = text });
            cell.StyleIndex = 8;
            cell.Append(istring);
            return cell;
        }

        /// <summary>
        /// Creates aligned text cell
        /// </summary>
        /// <param name="text">the text content</param>
        /// <returns>The cell</returns>
        internal static Cell CreateAlignedTextCell(string text)
        {
            Cell cell = new Cell { DataType = CellValues.InlineString };
            InlineString istring = new InlineString();
            istring.Append(new Text { Text = text });
            cell.StyleIndex = 13;
            cell.Append(istring);
            return cell;
        }

        /// <summary>
        /// Creates text cell
        /// </summary>
        /// <param name="header">the header</param>
        /// <param name="index">the index</param>
        /// <param name="text">the text content</param>
        /// <returns>The cell</returns>
        internal static Cell CreateTextCell(string header, uint index, string text)
        {
            Cell cell = new Cell { DataType = CellValues.InlineString, CellReference = header + index };
            InlineString istring = new InlineString();
            istring.Append(new Text { Text = text });
            cell.StyleIndex = 8;
            cell.Append(istring);
            return cell;
        }

        /// <summary>
        /// Creates number cell
        /// </summary>
        /// <param name="number">the number</param>
        /// <returns>The cell</returns>
        internal static Cell CreateNumberCell(int number)
        {
            Cell cell = new Cell();
            CellValue cellValue = new CellValue { Text = number.ToString() };
            cell.StyleIndex = 11;
            cell.Append(cellValue);
            return cell;
        }

        /// <summary>
        /// Creates the number cell
        /// </summary>
        /// <param name="header">the header</param>
        /// <param name="index">the index</param>
        /// <param name="number">the number</param>
        /// <returns>The cell</returns>
        internal static Cell CreateNumberCell(string header, uint index, int number)
        {
            Cell cell = new Cell { CellReference = header + index };
            CellValue cellValue = new CellValue { Text = number.ToString() };
            cell.StyleIndex = 8;
            cell.Append(cellValue);
            return cell;
        }

        /// <summary>
        /// Creates decimal cell
        /// </summary>
        /// <param name="number">the number</param>
        /// <returns>The cell</returns>
        internal static Cell CreateDecimalCell(decimal number)
        {
            Cell cell = new Cell();
            CellValue cellValue = new CellValue { Text = number.ToString() };
            cell.StyleIndex = 9;
            cell.Append(cellValue);
            return cell;
        }

        /// <summary>
        /// Creates decimal celle
        /// </summary>
        /// <param name="header">the header</param>
        /// <param name="index">the index</param>
        /// <param name="number">the number</param>
        /// <returns>The cell</returns>
        internal static Cell CreateDecimalCell(string header, uint index, decimal number)
        {
            Cell cell = new Cell { CellReference = header + index };
            CellValue cellValue = new CellValue { Text = number.ToString() };
            cell.StyleIndex = 9;
            cell.Append(cellValue);
            return cell;
        }

        /// <summary>
        /// Creates style sheet
        /// </summary>
        /// <returns>The style sheet</returns>
        internal static Stylesheet CreateStylesheet()
        {
            Stylesheet styleSheet = new Stylesheet();

            CellFormats cellFormats = new CellFormats();
            CellFormat cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            };
            cellFormats.Append(cellFormat);

            uint iExcelIndex = 164;
            NumberingFormats numberingFormats = new NumberingFormats();
            NumberingFormat nfDateTime = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            };
            numberingFormats.Append(nfDateTime);

            NumberingFormat nf4decimal = new NumberingFormat();
            nf4decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nf4decimal.FormatCode = StringValue.FromString("#,##0.0000");
            numberingFormats.Append(nf4decimal);

            // #,##0.00 is also Excel style index 4
            NumberingFormat nf2decimal = new NumberingFormat();
            nf2decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nf2decimal.FormatCode = StringValue.FromString("#,##0.00");
            numberingFormats.Append(nf2decimal);

            // @ is also Excel style index 49
            NumberingFormat nfForcedText = new NumberingFormat();
            nfForcedText.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfForcedText.FormatCode = StringValue.FromString("@");
            numberingFormats.Append(nfForcedText);

            NumberingFormat nfInteger = new NumberingFormat();
            nfInteger.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfInteger.FormatCode = StringValue.FromString("#");
            numberingFormats.Append(nfInteger);

            // index 1
            // Format dd/mm/yyyy
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 14;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 2
            // Format #,##0.00
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 4;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 3
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfDateTime.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 4
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nf4decimal.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 5
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nf2decimal.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 6
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfForcedText.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 7
            // Header text
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfForcedText.NumberFormatId;
            cellFormat.FontId = 1;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 8
            // column text
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfForcedText.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 1;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 9
            // coloured 2 decimal text
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nf2decimal.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 1;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 10
            // coloured column text
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfForcedText.NumberFormatId;
            cellFormat.FontId = 1;
            cellFormat.FillId = 2;
            cellFormat.BorderId = 1;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 11
            // coloured 2 decimal text
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfInteger.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 1;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            // index 12
            // coloured footer text
            cellFormat = new CellFormat
            {
                NumberFormatId = nfForcedText.NumberFormatId,
                FontId = 1,
                FillId = 3,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center },
                ApplyAlignment = true,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // index 13
            // column text
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = nfForcedText.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 1;
            cellFormat.FormatId = 0;
            cellFormat.Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };
            cellFormat.ApplyAlignment = true;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.Append(cellFormat);

            numberingFormats.Count = UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);

            styleSheet.Append(numberingFormats);
            CreateFonts(styleSheet);
            CreateFills(styleSheet);
            CreateBorders(styleSheet);
            CreateCellStyleFormats(styleSheet);
            styleSheet.Append(cellFormats);
            CreateCellStyles(styleSheet);
            styleSheet.Append(new DifferentialFormats { Count = 0 });
            styleSheet.Append(new TableStyles
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium9",
                DefaultPivotStyle = "PivotStyleLight16"
            });

            return styleSheet;
        }

        #endregion

        #region private methods

        /// <summary>
        /// Creates fonts
        /// </summary>
        /// <param name="styleSheet">The style sheet.</param>
        static void CreateFonts(Stylesheet styleSheet)
        {
            Fonts fonts = new Fonts();
            fonts.Append(new Font
            {
                FontName = new FontName { Val = StringValue.FromString("Calibri") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) }
            });
            fonts.Append(new Font
            {
                FontName = new FontName { Val = StringValue.FromString("Calibri") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) },
                Bold = new Bold { Val = BooleanValue.FromBoolean(true) }
            });

            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);

            styleSheet.Append(fonts);
        }

        /// <summary>
        /// Creates fills
        /// </summary>
        /// <param name="styleSheet">The style sheet.</param>
        static void CreateFills(Stylesheet styleSheet)
        {
            Fills fills = new Fills();
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); //0 
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); //1
            fills.Append(new Fill //2
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("01DF3A") },
                    BackgroundColor = new BackgroundColor { Rgb = HexBinaryValue.FromString("01DF3A") },
                }
            });
            fills.Append(new Fill //3
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFFFF") },
                    BackgroundColor = new BackgroundColor { Rgb = HexBinaryValue.FromString("FFFFFF") },
                }
            });

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);

            styleSheet.Append(fills);
        }

        /// <summary>
        /// Creayes borders
        /// </summary>
        /// <param name="styleSheet">The style sheet.</param>
        static void CreateBorders(Stylesheet styleSheet)
        {
            Borders borders = new Borders();
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });

            //Border Index 1
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });

            //Border Index 2
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });

            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);

            styleSheet.Append(borders);
        }

        /// <summary>
        /// Creates cell style formats
        /// </summary>
        /// <param name="styleSheet">The style sheet.</param>
        static void CreateCellStyleFormats(Stylesheet styleSheet)
        {
            CellStyleFormats cellStyleFormats = new CellStyleFormats();
            CellFormat cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellStyleFormats.Append(cellFormat);
            cellStyleFormats.Count = UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);

            styleSheet.Append(cellStyleFormats);
        }

        /// <summary>
        /// Creates cell styles
        /// </summary>
        /// <param name="styleSheet">The style sheet.</param>
        static void CreateCellStyles(Stylesheet styleSheet)
        {
            CellStyles cellStyles = new CellStyles();
            cellStyles.Append(new CellStyle
            {
                Name = "Normal",
                FormatId = 0,
                BuiltinId = 0
            });
            cellStyles.Count = (uint)cellStyles.ChildElements.Count;
            styleSheet.Append(cellStyles);
        }

        #endregion

    }
}
