namespace ExcelReportCreator
{
    public class ExcelColumn
    {
        public string Name { get; set; }

        public object Value { get; set; }

        public CellStyle HeaderStyle { get; set; }

        public CellStyle CellStyle { get; set; }

        public int Width { get; set; }

        public ExcelColumn() { }
    }
}
