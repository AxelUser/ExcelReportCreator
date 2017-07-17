using ExcelReportsCreator.Tests.Utils;
using System;
using Xunit;

namespace ExcelReportsCreator.Tests.Reports
{
    public class ReportBuilderContentTests
    {
        [Theory]
        [InlineData(10)]
        [InlineData(100000)]
        public void Build_PutCollection_EqualsCount(int count)
        {
            var entities = TestUtils.CreateTestReportEntities(count);

            throw new NotImplementedException();
        }

        [Fact]
        public void Build_PutEmptyCollection_ReturnNull()
        {
            throw new NotImplementedException();
        }

    }
}
