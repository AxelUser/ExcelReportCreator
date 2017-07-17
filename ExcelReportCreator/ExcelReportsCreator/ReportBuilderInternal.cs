using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelReportsCreator
{
    static class ReportBuilderInternal
    {
        /// <summary>
        /// Create title section.
        /// </summary>
        /// <param name="wSheet">Current excel-worksheet.</param>
        /// <param name="title">Worksheet title.</param>
        /// <param name="rowsToMerge">How many rows should be occupied by this section.</param>
        /// <param name="colsToMerge">How many columns should be occupied by this section.</param>
        public static void CreateTitle(ExcelWorksheet wSheet, string title, int rowsToMerge, int colsToMerge)
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
        public static void CreateHeader(ExcelWorksheet wSheet, List<ReportColumn> columns, int rowIndex, CellStyle defHeaderStyle)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                ReportColumn column = columns[i];
                var hStyle = column.HeaderStyle ?? defHeaderStyle;
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

                cell.Value = column.Title;
            }
        }

        /// <summary>
        /// Create section for new row.
        /// </summary>
        /// <param name="wSheet">Current excel-worksheet.</param>
        /// <param name="dataForRow">Column's infos for current row.</param>
        /// <param name="rowIndex">Row to start drawing section.</param>
        public static void CreateRow(ExcelWorksheet wSheet, List<ReportColumn> dataForRow, int rowIndex, CellStyle defDataCellStyle)
        {
            for (int i = 0; i < dataForRow.Count; i++)
            {
                ReportColumn column = dataForRow[i];
                var cStyle = column.CellStyle ?? defDataCellStyle;
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
        public static void ComputeHeaderSize(out int rowsCount, out int columnsCount, List<ReportColumn> columnsInfos, CellStyle defHeaderStyle)
        {
            rowsCount = 1;
            columnsCount = 0;
            foreach (var info in columnsInfos)
            {
                CellStyle style = info.HeaderStyle ?? defHeaderStyle;
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
        public static void ComputeRowSize(out int rowsCount, out int columnsCount, List<ReportColumn> columnsInfos, CellStyle defDataCellStyle)
        {
            rowsCount = 1;
            columnsCount = 0;
            foreach (var info in columnsInfos)
            {
                CellStyle style = info.CellStyle ?? defDataCellStyle;
                if (style.CellsToMergeUpright > rowsCount)
                {
                    rowsCount = style.CellsToMergeUpright;
                }
                columnsCount += style.CellsToMergeHorizontally;
            }
        }
    }
}
