namespace ExcelReportsCreator
{
    public class ReportColumn
    {
        public const int DefaultWidth = 40;

        public string Title { get; set; }

        public object Value { get; set; }

        public CellStyle HeaderStyle { get; set; }

        public CellStyle CellStyle { get; set; }

        public int Width { get; set; } = DefaultWidth;

        public ReportColumn() { }

        public ReportColumn(string title, object value, int width = DefaultWidth)
        {
            Title = title;
            Value = value;
            Width = width;
        }
    }
}
