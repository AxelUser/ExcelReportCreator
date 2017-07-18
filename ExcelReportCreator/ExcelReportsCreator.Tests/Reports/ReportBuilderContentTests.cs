using ExcelReportsCreator.Tests.Utils;
using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelReportsCreator.Tests.Reports
{
    public class ReportBuilderContentTests
    {
        [Fact]
        public void Build_AddColumn_ReturnsBinary()
        {
            var entities = TestUtils.CreateTestReportEntities(10);
            var binary = ReportBuilder<TestReportEntity>
                .Create(nameof(Build_AddColumn_ReturnsBinary))
                .AddColumn(entity => new ReportColumn()
                {
                    Title = "Title",
                    Value = entity.Title
                })
                .Build(entities);

            Assert.NotNull(binary);
            Assert.NotEmpty(binary);
        }

        [Fact]
        public void Build_PutEmptyCollection_ReturnsNull()
        {
            var binary = ReportBuilder<TestReportEntity>
                .Create(nameof(Build_PutEmptyCollection_ReturnsNull))
                .AddColumn(entity => new ReportColumn()
                {
                    Title = "Title",
                    Value = entity.Title
                })
                .Build(new List<TestReportEntity>());

            Assert.Null(binary);
        }

        [Fact]
        public void Build_PutNull_ReturnsNull()
        {
            var binary = ReportBuilder<TestReportEntity>
                .Create(nameof(Build_PutEmptyCollection_ReturnsNull))
                .AddColumn(entity => new ReportColumn()
                {
                    Title = "Title",
                    Value = entity.Title
                })
                .Build(null);

            Assert.Null(binary);
        }

        [Theory]
        [InlineData(10)]
        [InlineData(100000)]
        public void Build_PutCollection_EqualsCount(int count)
        {
            var entities = TestUtils.CreateTestReportEntities(count);

            throw new NotImplementedException();
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
