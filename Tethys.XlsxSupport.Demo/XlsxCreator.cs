// ---------------------------------------------------------------------------
// <copyright file="XlsxCreator.cs" company="Tethys">
//   Copyright (C) 2022-2023 T. Graf
// </copyright>
//
// Licensed under the Apache License, Version 2.0.
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
// either express or implied.
// SPDX-License-Identifier: Apache-2.0
// ---------------------------------------------------------------------------

namespace Tethys.XlsxSupport.Demo
{
    using System;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.ExtendedProperties;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public class XlsxCreator
    {
        /// <summary>
        /// Generates the report.
        /// </summary>
        /// <param name="filename">The filename.</param>
        public static void Generate(string filename)
        {
            using (var spreadsheet = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                // add WorkBookPart
                var workbookPart = spreadsheet.AddWorkbookPart();
                var workbook = new Workbook();
                workbookPart.Workbook = workbook;

                // Add a WorksheetPart to the WorkbookPart.
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                var sheet = new Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Table1",
                };

                var row = new Row() { RowIndex = 1 };
                var header1 = new Cell() { CellReference = "A1", CellValue = new CellValue("Test"), DataType = CellValues.String };
                row.Append(header1);

                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                sheetData.Append(row);

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                spreadsheet.PackageProperties.Title = $"My Title";
                spreadsheet.PackageProperties.Subject = "My Subject";
                spreadsheet.PackageProperties.Creator = "Me";

                spreadsheet.AddExtendedFilePropertiesPart();
                spreadsheet.ExtendedFilePropertiesPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                spreadsheet.ExtendedFilePropertiesPart.Properties.Company = new Company("Tethys");

                var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                GenerateCreateExcelStyles(stylesPart);

                var firstSheet = workbookPart.Workbook.Descendants<Sheet>().First();
                var worksheet = ((WorksheetPart)workbookPart.GetPartById(firstSheet.Id)).Worksheet;

                // set columns width
                var columns = new Columns();
                // Project
                columns.Append(new Column { Min = 1, Max = 1, Width = 34, Style = 0, CustomWidth = true });

                // Requested
                columns.Append(new Column { Min = 2, Max = 2, Width = 32, Style = 0, CustomWidth = true });

                // Comment
                columns.Append(new Column { Min = 5, Max = 5, Width = 30, Style = 0, CustomWidth = true });

                worksheet.InsertBefore(columns, sheetData);

                UInt32Value rowIndex = 1;

                // title row
                row = new Row { RowIndex = rowIndex++ };
                var text = $"Some text";
                var header = BasicExcelSupport.ConstructCell(
                    "A1", text, CellValues.String, 0);
                row.Append(header);
                sheetData.Append(row);

                rowIndex++;
                // header row
                row = new Row { RowIndex = rowIndex++ };
                row.Append(BasicExcelSupport.ConstructTextCell(
                    "Project", 1));
                row.Append(BasicExcelSupport.ConstructTextCell(
                    "Requested", 1));
                row.Append(BasicExcelSupport.ConstructTextCell(
                    "Comment", 1));
                sheetData.Append(row);

                // data - first row
                row = new Row { RowIndex = rowIndex++ };
                row.Append(BasicExcelSupport.ConstructTextCell("Project A", 0));
                row.Append(new Cell
                {
                    CellValue = new CellValue(DateTime.Now),
                    DataType = new EnumValue<CellValues>(CellValues.Date),
                    StyleIndex = 2,
                });
                row.Append(BasicExcelSupport.ConstructTextCell("Comment", 0));

                sheetData.Append(row);

                // data - second row
                row = new Row { RowIndex = rowIndex++ };
                row.Append(BasicExcelSupport.ConstructTextCell("Project B", 0));
                row.Append(BasicExcelSupport.ConstructTextCell(DateTime.Now.ToString("O"), 2));
                row.Append(BasicExcelSupport.ConstructTextCell("Comment X", 0));

                sheetData.Append(row);

                // (auto) filtering
                var autoFilter = new AutoFilter
                {
                    Reference = $"A3:C{rowIndex}",
                };
                // The mandatory order in the sheet seems to be
                // 1. auto filter
                // 2. columns
                // 3. styles
                worksheet.Append(autoFilter);

                // merge title columns
                var mergeCells = new MergeCells();
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:C1") });
                worksheet.Append(mergeCells);

                workbookPart.Workbook.Save();
                spreadsheet.Save();
            } // using
        } // Generate()

        /// <summary>
        /// Creates the style cell formats.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        private static void CreateStyleCellFormats(Stylesheet stylesheet)
        {
            var cellStyleFormats = new CellStyleFormats { Count = 1U };

            var cellFormatDefault = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
            cellStyleFormats.Append(cellFormatDefault);

            var cellFormats = new CellFormats { Count = 10U };

            // index 0 - Calibri 10pt, borders = none, the DEFAULT format
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
            });

            // index 1 - Calibri, 11pt, bold, borders = bottom
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 1U,
                FillId = 0U,
                BorderId = 2U,
                FormatId = 0U,
            });

            // index 2 - Calibri 10pt, with number format
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 164U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyNumberFormat = true,
            });

            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
        } // CreateStyleCellFormats()

        /// <summary>
        /// Generates the content of the workbook styles part.
        /// </summary>
        /// <param name="workbookStylesPart">The workbook styles part.</param>
        private static void GenerateCreateExcelStyles(WorkbookStylesPart workbookStylesPart)
        {
            var stylesheet1 = new Stylesheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //// ----------------------------------------

            // the order is IMPORTANT - otherwise Excel will report errors
            // 1. Number formats
            // 2. Font styles
            // 3. Fill styles
            // 4. Border styles
            // 5. Cell style formats
            // 6. Cell formats
            CreateNumberFormats(stylesheet1);
            CreateStyleFonts(stylesheet1);
            CreateStyleFills(stylesheet1);
            CreateStyleBorders(stylesheet1);
            CreateStyleCellFormats(stylesheet1);

            //// ----------------------------------------

            var cellStyles1 = new CellStyles { Count = 1U };
            var cellStyle1 = new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U };

            cellStyles1.Append(cellStyle1);
            var differentialFormats1 = new DifferentialFormats { Count = 0U };
            var tableStyles1 = new TableStyles
            {
                Count = 0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16",
            };

            var stylesheetExtensionList1 = new StylesheetExtensionList();

            var stylesheetExtension1 = new StylesheetExtension
            {
                Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
            };
            stylesheetExtension1.AddNamespaceDeclaration(
                "x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            stylesheetExtensionList1.Append(stylesheetExtension1);

            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart.Stylesheet = stylesheet1;
        } // GenerateCreateExcelStyles()

        /// <summary>
        /// Creates the number formats.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        public static void CreateNumberFormats(Stylesheet stylesheet)
        {
            var numberingFormats = new NumberingFormats { Count = 1U };

            // 164 seems to be a good start index value for custom formats
            // http://polymathprogrammer.com/2009/11/09/how-to-create-stylesheet-in-excel-open-xml/
            // 5 seems also so be useable, see
            // https://docs.microsoft.com/de-de/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
            // NumberFormatId = 5,
            uint numberFormatIndex = 164;
            numberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = numberFormatIndex++,
                FormatCode = "mmmm-YY", // "Juni 20"
            });

            numberingFormats.Append(new NumberingFormat
            {
                // ReSharper disable once RedundantAssignment
                NumberFormatId = numberFormatIndex++,
                FormatCode = "DD.MM.YYYY", // "21.12.2021"
            });

            stylesheet.Append(numberingFormats);
        } // CreateNumberFormats()

        /// <summary>
        /// Creates the style fonts.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        public static void CreateStyleFonts(Stylesheet stylesheet)
        {
            var fonts = new Fonts { Count = 4U, KnownFonts = true };

            // Index 0 - Calibri, 10pt ==> first font = default
            var fontCalibri10 = new Font();
            fontCalibri10.Append(new FontSize { Val = 10D });
            fontCalibri10.Append(new Color { Theme = 1U });
            fontCalibri10.Append(new FontName { Val = "Calibri" });
            fontCalibri10.Append(new FontFamilyNumbering { Val = 2 });
            fontCalibri10.Append(new FontScheme { Val = FontSchemeValues.Minor });
            fonts.Append(fontCalibri10);

            // Index 1 - Calibri, 11pt, bold
            var fontCalibri11Bold = new Font();
            fontCalibri11Bold.Append(new Bold());
            fontCalibri11Bold.Append(new FontSize { Val = 11D });
            fontCalibri11Bold.Append(new Color { Theme = 1U });
            fontCalibri11Bold.Append(new FontName { Val = "Calibri" });
            fontCalibri11Bold.Append(new FontFamilyNumbering { Val = 2 });
            fontCalibri11Bold.Append(new FontScheme { Val = FontSchemeValues.Minor });
            fonts.Append(fontCalibri11Bold);

            stylesheet.Append(fonts);
        } // CreateStyleFonts()

        /// <summary>
        /// Creates the style fills.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        public static void CreateStyleFills(Stylesheet stylesheet)
        {
            var fills = new Fills { Count = 5U };

            // Index 0 - always: no fill
            var fillNone = new Fill();
            fillNone.Append(new PatternFill { PatternType = PatternValues.None });
            fills.Append(fillNone);

            // Index 1 - always: Gray125
            var fillPatternGray125 = new Fill();
            fillPatternGray125.Append(new PatternFill { PatternType = PatternValues.Gray125 });
            fills.Append(fillPatternGray125);

            // Index 23 - Siemens White(FFFFFFFF)
            var fillWhite = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.Solid };
            patternFill.Append(new ForegroundColor { Rgb = "FFFFFFFF" });
            patternFill.Append(new BackgroundColor { Indexed = 64U });
            fillWhite.Append(patternFill);
            fills.Append(fillWhite);

            stylesheet.Append(fills);
        } // CreateStyleFills()

        /// <summary>
        /// Creates the style borders.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        public static void CreateStyleBorders(Stylesheet stylesheet)
        {
            var borders = new Borders { Count = 2U };

            // Index 0 - no borders
            var borderNone = new Border();
            borderNone.Append(new LeftBorder());
            borderNone.Append(new RightBorder());
            borderNone.Append(new TopBorder());
            borderNone.Append(new BottomBorder());
            borderNone.Append(new DiagonalBorder());
            borders.Append(borderNone);

            // Index 1 - thin borders
            var borderThinAll = new Border();
            var leftBorderThin = new LeftBorder { Style = BorderStyleValues.Thin };
            leftBorderThin.Append(new Color { Indexed = 64U });
            var rightBorderThin = new RightBorder { Style = BorderStyleValues.Thin };
            rightBorderThin.Append(new Color { Indexed = 64U });
            var topBorderThin = new TopBorder { Style = BorderStyleValues.Thin };
            topBorderThin.Append(new Color { Indexed = 64U });
            var bottomBorderThin = new BottomBorder { Style = BorderStyleValues.Thin };
            bottomBorderThin.Append(new Color { Indexed = 64U });
            var diagonalBorderNone = new DiagonalBorder();

            borderThinAll.Append(leftBorderThin);
            borderThinAll.Append(rightBorderThin);
            borderThinAll.Append(topBorderThin);
            borderThinAll.Append(bottomBorderThin);
            borderThinAll.Append(diagonalBorderNone);

            borders.Append(borderThinAll);

            // Index 2 - thin borders at bottom
            var borderThinBottom = new Border();
            borderThinBottom.Append(new LeftBorder());
            borderThinBottom.Append(new RightBorder());
            borderThinBottom.Append(new TopBorder());
            borderThinBottom.Append(new BottomBorder { Style = BorderStyleValues.Thin });
            borderThinBottom.Append(new DiagonalBorder());
            borders.Append(borderThinBottom);

            // Index 3 - thin borders left and right
            var borderLeftRight = new Border();
            borderLeftRight.Append(new LeftBorder { Style = BorderStyleValues.Thin });
            borderLeftRight.Append(new RightBorder { Style = BorderStyleValues.Thin });
            borderLeftRight.Append(new TopBorder());
            borderLeftRight.Append(new BottomBorder());
            borderLeftRight.Append(new DiagonalBorder());
            borders.Append(borderLeftRight);

            stylesheet.Append(borders);
        } // CreateStyleBorders()
    }
}
