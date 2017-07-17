using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportsCreator.Tests.Utils
{
    public static class TestUtils
    {
        static Random _random;

        static TestUtils()
        {
            _random = new Random();
        }

        public static List<TestReportEntity> CreateTestReportEntities(int count = 5)
        {
            List<TestReportEntity> entities = new List<TestReportEntity>();

            for (int i = 0; i < count; i++)
            {
                entities.Add(new TestReportEntity()
                {
                    Id = Guid.NewGuid(),
                    Title = $"{nameof(TestReportEntity)} {i + 1}",
                    Count = _random.Next(0, 1000),
                    Date = DateTime.Now.AddHours(_random.Next(-10000, 10000))
                });
            }

            return entities;
        }
    }
}
