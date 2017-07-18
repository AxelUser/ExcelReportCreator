using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportsCreator
{
    public class ReportBuilder<T> where T : new()
    {
        /// <summary>
        /// List of column-infos.
        /// </summary>
        protected List<Func<T, ReportColumn>> _rowsCreator;

        /// <summary>
        /// Default header style for columns.
        /// </summary>
        private CellStyle _defHeaderStyle;

        /// <summary>
        /// Default data-cell style for columns.
        /// </summary>
        private CellStyle _defDataCellStyle;

        /// <summary>
        /// Template for row generation.
        /// </summary>
        public List<Func<T, ReportColumn>> RowTemplate
        {
            get
            {
                return _rowsCreator;
            }
        }

        public CellStyle DefaultHeaderStyle
        {
            get
            {
                return _defHeaderStyle;
            }
        }

        public CellStyle DefaultDataCellStyle
        {
            get
            {
                return _defDataCellStyle;
            }
        }

        /// <summary>
        /// Title of report.
        /// </summary>
        public string ReportsTitle { get; set; }

        public ReportBuilder(string title)
        {
            _rowsCreator = new List<Func<T, ReportColumn>>();
            _defHeaderStyle = new CellStyle();
            _defDataCellStyle = new CellStyle();
            ReportsTitle = title;
        }

        /// <summary>
        /// Start creating report.
        /// </summary>
        /// <returns></returns>
        public static ReportBuilder<T> Create(string title)
        {
            return new ReportBuilder<T>(title);
        }

        /// <summary>
        /// Add column info with data-mapping and style.
        /// </summary>
        /// <param name="columnCreator">Data mapping and style.</param>
        /// <returns>Updated reporter.</returns>
        public ReportBuilder<T> AddColumn(Func<T, ReportColumn> columnCreator)
        {
            _rowsCreator.Add(columnCreator);
            return this;
        }

        /// <summary>
        /// Initialize default header style.
        /// </summary>
        /// <param name="headerStyle">Style for column's headers.</param>
        /// <returns>Updated reporter.</returns>
        public ReportBuilder<T> SetDefHeaderStyle(CellStyle headerStyle)
        {
            _defHeaderStyle = headerStyle;
            return this;
        }

        /// <summary>
        /// Initialize default data-cell style.
        /// </summary>
        /// <param name="dataCellStyle">Style for column's data-cells.</param>
        /// <returns>Updated reporter.</returns>
        public ReportBuilder<T> SetDefDataCellStyle(CellStyle dataCellStyle)
        {
            _defDataCellStyle = dataCellStyle;
            return this;
        }

        /// <summary>
        /// Convert report from provided data.
        /// </summary>
        /// <param name="collection">Collection of entities.</param>
        /// <returns>Report in binary format.</returns>
        public byte[] Build(IEnumerable<T> collection)
        {
            if(collection == null || !collection.Any())
            {
                return null;
            }

            if(!_rowsCreator.Any())
            {
                throw new ReportBuilderException("Report must have at least one column.");
            }

            using (ExcelPackage excellPack = new ExcelPackage())
            {
                //TODO нужно обработать случай для пустой коллекции.
                var workSheet = excellPack.Workbook.Worksheets.Add(ReportsTitle);
                T dummy = new T();
                List<ReportColumn> columnsInfos = _rowsCreator.Select(c => c(dummy)).ToList();

                int headerColumnsCount = 0;
                int headerRowsCount = 0;
                ReportBuilderInternal.ComputeHeaderSize(out headerRowsCount, out headerColumnsCount, columnsInfos, _defHeaderStyle);

                ReportBuilderInternal.CreateTitle(workSheet, ReportsTitle, 2, headerColumnsCount);
                ReportBuilderInternal.CreateHeader(workSheet, columnsInfos, 3, _defHeaderStyle);

                int rowIndex = 4;
                int dataRowsCount = 0;
                int dataColumnsCount = 0;
                ReportBuilderInternal.ComputeRowSize(out dataRowsCount, out dataColumnsCount, columnsInfos, _defDataCellStyle);
                foreach (T item in collection)
                {
                    List<ReportColumn> dataForRow = _rowsCreator.Select(cr => cr(item)).ToList();
                    ReportBuilderInternal.CreateRow(workSheet, dataForRow, rowIndex, _defDataCellStyle);
                    rowIndex += dataRowsCount;
                }
                return excellPack.GetAsByteArray();
            }
        }

 
    }
}


