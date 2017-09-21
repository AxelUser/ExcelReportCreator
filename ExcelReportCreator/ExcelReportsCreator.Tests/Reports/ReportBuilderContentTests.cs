using ExcelReportsCreator.Tests.Utils;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace ExcelReportsCreator.Tests.Reports
{
    public class ReportBuilderContentTests
    {
        [Fact]
        public void Build_AddColumn_ReturnsBinary()
        {
            var entities = TestUtils.CreateTestReportEntities(10);
            var binary = TestUtils.CreateReportWithTitleColumn(nameof(Build_AddColumn_ReturnsBinary), entities);

            Assert.NotNull(binary);
            Assert.NotEmpty(binary);
        }

        [Fact]
        public void Build_PutEmptyCollection_ReturnsNull()
        {
            var binary = TestUtils.CreateReportWithTitleColumn(nameof(Build_PutEmptyCollection_ReturnsNull), 
                new List<TestReportEntity>());

            Assert.Null(binary);
        }

        [Fact]
        public void Build_PutNull_ReturnsNull()
        {
            var binary = TestUtils.CreateReportWithTitleColumn(nameof(Build_PutNull_ReturnsNull), null);

            Assert.Null(binary);
        }

        [Theory]
        [InlineData(10)]
        public void Build_PutCollection_EqualsCount(int count)
        {
            string title;
            string header;
            List<string> rowValues = new List<string>();

            var entities = TestUtils.CreateTestReportEntities(count);
            var binary = TestUtils.CreateReportWithTitleColumn(nameof(Build_PutCollection_EqualsCount), entities);

            Assert.NotNull(binary);

            using (MemoryStream ms = new MemoryStream(binary))
            {
                using (ExcelPackage package = new ExcelPackage(ms))
                {
                    ExcelWorksheet w = package.Workbook.Worksheets[1];
                    title = w.GetValue(1, 1).ToString();
                    header = w.GetValue(3, 1).ToString();

                    
                    for(int i=4; i < count+4; i++)
                    {
                        rowValues.Add(w.GetValue(i, 1).ToString());
                    }
                }
            }

            Assert.Equal(nameof(Build_PutCollection_EqualsCount), title);
            Assert.Equal(nameof(TestReportEntity.Title), header);
            Assert.Equal(entities.Select(e => e.Title), rowValues);
        }

        [Fact]
        public void Build_NoColumns_ThrowsReportBuilderException()
        {
            var entities = TestUtils.CreateTestReportEntities(10);
            var builder = ReportBuilder<TestReportEntity>
                .Create(nameof(Build_AddColumn_ReturnsBinary));

            Assert.Throws<ReportBuilderException>(() => builder.Build(entities));
        }
    }
}
