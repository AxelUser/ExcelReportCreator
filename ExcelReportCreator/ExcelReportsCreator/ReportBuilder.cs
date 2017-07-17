using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportCreator
{
    class ReportBuilderInternal
    {
        public class CustomReport<T> where T : new()
        {
            /// <summary>
            /// List of column-infos.
            /// </summary>
            List<Func<T, ExcelColumn>> _rowsCreator;

            /// <summary>
            /// Default header style for columns.
            /// </summary>
            private CellStyle _defHeaderStyle;

            /// <summary>
            /// Default data-cell style for columns.
            /// </summary>
            private CellStyle _defDataCellStyle;

            /// <summary>
            /// Title of report.
            /// </summary>
            public string ReportsTitle { get; set; }

            public CustomReport()
            {
                _rowsCreator = new List<Func<T, ExcelColumn>>();
                _defHeaderStyle = new CellStyle();
                _defDataCellStyle = new CellStyle();
            }

            /// <summary>
            /// Add column info with data-mapping and style.
            /// </summary>
            /// <param name="columnCreator">Data mapping and style.</param>
            /// <returns>Updated reporter.</returns>
            public CustomReport<T> AddColumn(Func<T, ExcelColumn> columnCreator)
            {
                _rowsCreator.Add(columnCreator);
                return this;
            }

            /// <summary>
            /// Initialize default header style.
            /// </summary>
            /// <param name="headerStyle">Style for column's headers.</param>
            /// <returns>Updated reporter.</returns>
            public CustomReport<T> SetDefHeaderStyle(CellStyle headerStyle)
            {
                _defHeaderStyle = headerStyle;
                return this;
            }

            /// <summary>
            /// Initialize default data-cell style.
            /// </summary>
            /// <param name="dataCellStyle">Style for column's data-cells.</param>
            /// <returns>Updated reporter.</returns>
            public CustomReport<T> SetDefDataCellStyle(CellStyle dataCellStyle)
            {
                _defDataCellStyle = dataCellStyle;
                return this;
            }

            /// <summary>
            /// Convert report from provided data.
            /// </summary>
            /// <param name="collection">Collection of entities.</param>
            /// <returns>Report in binary format.</returns>
            public byte[] Create(IEnumerable<T> collection)
            {
                using (ExcelPackage excellPack = new ExcelPackage())
                {
                    //TODO нужно обработать случай для пустой коллекции.
                    var workSheet = excellPack.Workbook.Worksheets.Add(ReportsTitle);
                    T dummy = new T();
                    List<ExcelColumn> columnsInfos = _rowsCreator.Select(c => c(dummy)).ToList();

                    int headerColumnsCount = 0;
                    int headerRowsCount = 0;
                    ComputeHeaderSize(out headerRowsCount, out headerColumnsCount, columnsInfos);

                    CreateTitle(workSheet, ReportsTitle, 2, headerColumnsCount);
                    CreateHeader(workSheet, columnsInfos, 3);

                    int rowIndex = 4;
                    int dataRowsCount = 0;
                    int dataColumnsCount = 0;
                    ComputeRowSize(out dataRowsCount, out dataColumnsCount, columnsInfos);
                    foreach (T item in collection)
                    {
                        List<ExcelColumn> dataForRow = _rowsCreator.Select(cr => cr(item)).ToList();
                        CreateRow(workSheet, dataForRow, rowIndex);
                        rowIndex += dataRowsCount;
                    }
                    return excellPack.GetAsByteArray();
                }
            }

            /// <summary>
            /// Create title section.
            /// </summary>
            /// <param name="wSheet">Current excel-worksheet.</param>
            /// <param name="title">Worksheet title.</param>
            /// <param name="rowsToMerge">How many rows should be occupied by this section.</param>
            /// <param name="colsToMerge">How many columns should be occupied by this section.</param>
            private void CreateTitle(ExcelWorksheet wSheet, string title, int rowsToMerge, int colsToMerge)
            {
                var headerCells = wSheet.Cells[1, 1, rowsToMerge, colsToMerge];
                headerCells.Merge = true;
                headerCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                headerCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                headerCells.Style.Font.Bold = true;
                headerCells.Value = title;
            }

            /// <summary>
            /// Create header section.
            /// </summary>
            /// <param name="wSheet">Current excel-worksheet.</param>
            /// <param name="columns">Column's infos.</param>
            /// <param name="rowIndex">Row to start drawing header section.</param>
            private void CreateHeader(ExcelWorksheet wSheet, List<ExcelColumn> columns, int rowIndex)
            {
                for (int i = 0; i < columns.Count; i++)
                {
                    ExcelColumn column = columns[i];
                    var hStyle = column.HeaderStyle ?? _defHeaderStyle;
                    ExcelRange cell = null;
                    if (hStyle.CellsToMergeHorizontally > 1 || hStyle.CellsToMergeUpright > 1)
                    {
                        cell = wSheet.Cells[rowIndex, i + 1, i + hStyle.CellsToMergeUpright, rowIndex + hStyle.CellsToMergeHorizontally - 1];
                        cell.Merge = true;
                    }
                    else
                    {
                        cell = wSheet.Cells[rowIndex, i + 1];
                    }
                    if (hStyle.Border)
                    {
                        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    if (!hStyle.CellsColor.IsEmpty)
                    {
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(hStyle.CellsColor);
                    }
                    cell.Style.WrapText = hStyle.WordWrap;

                    int colsCount = cell.End.Column - cell.Start.Column + 1;
                    for (int j = cell.Start.Column; j <= cell.End.Column; j++)
                    {
                        wSheet.Column(j).Width = column.Width / colsCount;
                    }

                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.Font.Bold = hStyle.BoldText;

                    cell.Value = column.Name;
                }
            }

            /// <summary>
            /// Create section for new row.
            /// </summary>
            /// <param name="wSheet">Current excel-worksheet.</param>
            /// <param name="dataForRow">Column's infos for current row.</param>
            /// <param name="rowIndex">Row to start drawing section.</param>
            private void CreateRow(ExcelWorksheet wSheet, List<ExcelColumn> dataForRow, int rowIndex)
            {
                for (int i = 0; i < dataForRow.Count; i++)
                {
                    ExcelColumn column = dataForRow[i];
                    var cStyle = column.CellStyle ?? _defDataCellStyle;
                    ExcelRange cell = null;
                    if (cStyle.CellsToMergeHorizontally > 1 || cStyle.CellsToMergeUpright > 1)
                    {
                        cell = wSheet.Cells[rowIndex, i + 1, i + cStyle.CellsToMergeUpright, rowIndex + cStyle.CellsToMergeHorizontally - 1];
                        cell.Merge = true;
                    }
                    else
                    {
                        cell = wSheet.Cells[rowIndex, i + 1];
                    }
                    if (cStyle.Border)
                    {
                        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    }
                    if (!cStyle.CellsColor.IsEmpty)
                    {
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(cStyle.CellsColor);
                    }
                    cell.Style.WrapText = cStyle.WordWrap;
                    cell.Style.Font.Bold = cStyle.BoldText;

                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    wSheet.Cells[rowIndex, i + 1].Value = column.Value;
                }
            }

            /// <summary>
            /// Compute sizes for header.
            /// </summary>
            /// <param name="rowsCount">Rows for header.</param>
            /// <param name="columnsCount">Columns for header.</param>
            /// <param name="columnsInfos">Collection of column's infos.</param>
            private void ComputeHeaderSize(out int rowsCount, out int columnsCount, List<ExcelColumn> columnsInfos)
            {
                rowsCount = 1;
                columnsCount = 0;
                foreach (var info in columnsInfos)
                {
                    CellStyle style = info.HeaderStyle ?? _defHeaderStyle;
                    if (style.CellsToMergeUpright > rowsCount)
                    {
                        rowsCount = style.CellsToMergeUpright;
                    }
                    columnsCount += style.CellsToMergeHorizontally;
                }
            }

            /// <summary>
            /// Compute sizes for data-row.
            /// </summary>
            /// <param name="rowsCount">Rows for data.</param>
            /// <param name="columnsCount">Columns for data.</param>
            /// <param name="columnsInfos">Collection of column's infos.</param>
            private void ComputeRowSize(out int rowsCount, out int columnsCount, List<ExcelColumn> columnsInfos)
            {
                rowsCount = 1;
                columnsCount = 0;
                foreach (var info in columnsInfos)
                {
                    CellStyle style = info.CellStyle ?? _defDataCellStyle;
                    if (style.CellsToMergeUpright > rowsCount)
                    {
                        rowsCount = style.CellsToMergeUpright;
                    }
                    columnsCount += style.CellsToMergeHorizontally;
                }
            }
        }
    }
}


