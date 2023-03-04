// ---------------------------------------------------------------------------
// <copyright file="BasicExcelSupport.cs" company="Tethys">
//   Copyright (C) 2021-2023 T. Graf
// </copyright>
//
// Licensed under the Apache License, Version 2.0.
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
// either express or implied.
// SPDX-License-Identifier: Apache-2.0
// ---------------------------------------------------------------------------

/*****************************************************************************
 * Required NuGet Packages
 * -----------------------
 * - DocumentFormat.OpenXml 2.12.3
 *
 ****************************************************************************/

namespace Tethys.XlsxSupport
{
    using System;
    using System.Diagnostics;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Validation;

    using Tethys.Logging;

    /// <summary>
    /// Generator methods for Excel documents.
    /// </summary>
    public class BasicExcelSupport
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The logger for this class.
        /// </summary>
        private static readonly ILog Log = LogManager.GetLogger(typeof(BasicExcelSupport));
        #endregion // PRIVATE PROPERTIES

        //// ---------------------------------------------------------------------

        #region PUBLIC METHODS
        /// <summary>
        /// Opens the document in Excel.
        /// </summary>
        /// <param name="filename">The fileName.</param>
        /// <remarks>
        /// Opening an Excel sheet this work works perfectly for .NET Framework 4.8
        /// application, but fails for .NET 5.0, independently whether this is a console
        /// or WinForms .NET 5 application.
        /// </remarks>
        public static void OpenDocumentInExcel(string filename)
        {
            Log.InfoFormat("Opening Microsoft Excel for file '{0}'", filename);

            try
            {
                var process = new Process();
                process.StartInfo.UseShellExecute = true;
                process.StartInfo.FileName = "EXCEL.EXE";
                process.StartInfo.Arguments = $"\"{filename}\"";
                process.Start();
            }
            catch (Exception ex)
            {
                Log.Error("Error opening Microsoft Excel", ex);
            } // catch
        } // OpenDocumentInExcel()

        /// <summary>
        /// Validates the word document.
        /// </summary>
        /// <param name="filepath">The file path.</param>
        /// <returns>The number of errors found.</returns>
        public static int ValidateExcelDocument(string filepath)
        {
            // https://docs.microsoft.com/de-de/office/open-xml/how-to-validate-a-word-processing-document
            int count;
            using (var doc = SpreadsheetDocument.Open(filepath, true))
            {
                count = ValidateExcelDocument(doc);
            } // using

            return count;
        } // ValidateExcelDocument()

        /// <summary>
        /// Validates the given Excel document.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns>
        /// The number of errors found.
        /// </returns>
        public static int ValidateExcelDocument(SpreadsheetDocument doc)
        {
            var count = 0;
            try
            {
                Log.Debug("Validating document...");
                var validator = new OpenXmlValidator();
                foreach (var error in validator.Validate(doc))
                {
                    count++;
                    Log.Error("Error " + count);
                    Log.Error("Description: " + error.Description);
                    Log.Error("ErrorType: " + error.ErrorType);
                    Log.Error("Node: " + error.Node);
                    Log.Error("Path: " + error.Path.XPath);
                    Log.Error("Part: " + error.Part.Uri);
                    Log.Error("-------------------------------------------");
                } // foreach

                if (count > 0)
                {
                    Log.Error($"Total issue count={count}");
                } // if
            }
            catch (Exception ex)
            {
                Log.Error("Error validating document: " + ex.Message);
            } // catch

            return count;
        } // ValidateExcelDocument()

        /// <summary>
        /// Creates the simple excel sheet.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns>A <see cref="SpreadsheetDocument"/>.</returns>
        public static SpreadsheetDocument CreateSimpleExcelSheet(string filename)
        {
            var spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);

            // add WorkBookPart
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            var workbook = new Workbook();
            workbookPart.Workbook = workbook;

            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
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

            return spreadsheetDocument;
        } // CreateSimpleExcelSheet()

        /// <summary>
        /// Constructs a cell with the specified contents.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="dataType">Type of the data.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns>A <see cref="Cell"/>.</returns>
        public static Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex,
            };
        } // ConstructCell()

        /// <summary>
        /// Constructs a cell with the specified contents.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns>A <see cref="Cell"/>.</returns>
        public static Cell ConstructTextCell(string value, uint styleIndex = 0)
        {
            return ConstructCell(value, CellValues.String, styleIndex);
        } // ConstructTextCell()

        /// <summary>
        /// Constructs a cell with the specified contents.
        /// </summary>
        /// <param name="reference">The reference.</param>
        /// <param name="value">The value.</param>
        /// <param name="dataType">Type of the data.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns>
        /// A <see cref="Cell" />.
        /// </returns>
        public static Cell ConstructCell(StringValue reference, string value, CellValues dataType, uint styleIndex = 0)
        {
            var cell = new Cell
            {
                CellReference = reference,
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
            };

            if (styleIndex != 0)
            {
                cell.StyleIndex = styleIndex;
            } // if

            return cell;
        } // ConstructCell()

        /// <summary>
        /// Constructs a cell with the specified contents.
        /// </summary>
        /// <param name="reference">The reference.</param>
        /// <param name="value">The value.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns>
        /// A <see cref="Cell" />.
        /// </returns>
        public static Cell ConstructTextCell(StringValue reference, string value, uint styleIndex = 0)
        {
            return ConstructCell(reference, value, CellValues.String, styleIndex);
        } // ConstructTextCell()

        /// <summary>
        /// Gets the row with the given index.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns>A <see cref="Row"/>.</returns>
        public static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().First(r => r.RowIndex == rowIndex);
        } // GetRow()

        /// <summary>
        /// Gets the specified cell.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns>A <see cref="Cell"/>.</returns>
        public static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            var row = GetRow(worksheet, rowIndex);

            return (row?.Elements<Cell>() ?? throw new InvalidOperationException()).First(c =>
                string.Compare(
                    c.CellReference.Value,
                    columnName + rowIndex,
                    StringComparison.OrdinalIgnoreCase) == 0);
        } // GetCell()

        /// <summary>
        /// Gets the shared string item by identifier.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="id">The identifier.</param>
        /// <returns>A <see cref="SharedStringItem"/>.</returns>
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        } // GetSharedStringItemById()

        /// <summary>
        /// Gets the worksheet part for the sheet with teh specified name.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>A <see cref="WorksheetPart"/>.</returns>
        public WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        } // GetWorksheetPart()
        #endregion // PUBLIC METHODS
    } // BasicExcelSupport
}
