namespace ExcelReportsCreator
{
    public class ReportColumn
    {
        public string Title { get; set; }

        public object Value { get; set; }

        public CellStyle HeaderStyle { get; set; }

        public CellStyle CellStyle { get; set; }

        public int Width { get; set; } = 40;

        public ReportColumn() { }
    }
}
