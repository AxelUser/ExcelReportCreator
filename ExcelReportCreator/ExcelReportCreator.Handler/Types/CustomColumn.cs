using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportCreator.Handler.Types
{
    public class CustomColumn
    {
        public string Name { get; set; }
        public object Value { get; set; }
        public CellStyle HeaderStyle { get; set; }
        public CellStyle CellStyle { get; set; }
        public int Width { get; set; }
        public CustomColumn()
        {
        }
    }
}
