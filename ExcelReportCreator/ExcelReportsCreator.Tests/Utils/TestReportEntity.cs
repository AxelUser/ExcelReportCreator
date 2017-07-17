using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportsCreator.Tests.Utils
{
    public class TestReportEntity
    {
        public Guid Id { get; set; }

        public string Title { get; set; }

        public int Count { get; set; }

        public DateTime Date { get; set; }
    }
}
