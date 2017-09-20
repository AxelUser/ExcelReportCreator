using ExcelReportsCreator.Tests.Utils;
using System.IO;
using Xunit;

namespace ExcelReportsCreator.Tests.Reports
{
    public class ReportBuilderFunctionalTests: IClassFixture<ReportBuilderFunctionalTestsFixture>
    {
        private ReportBuilderFunctionalTestsFixture _fixture;

        public ReportBuilderFunctionalTests(ReportBuilderFunctionalTestsFixture fixture)
        {
            _fixture = fixture;
        }

        [Theory, InlineData(5)]
        public void Build_Run(int entitiesCount)
        {
            string reportName = $"{nameof(Build_Run)}.xlsx";

            var entities = TestUtils.CreateTestReportEntities(entitiesCount);
            var binary = ReportBuilder<TestReportEntity>
                .Create(nameof(Build_Run))
                .AddColumn(entity => new ReportColumn()
                {
                    Title = "Title",
                    Value = entity.Title
                })
                .AddColumn(entity => new ReportColumn("Date", entity.Date.ToShortDateString()))
                .Build(entities);

            File.WriteAllBytes(Path.Combine(_fixture.ReportsDirectory.FullName, reportName), binary);
        }
    }

    public class ReportBuilderFunctionalTestsFixture
    {
        private const string ReportsDirectoryName = "Reports";

        public DirectoryInfo ReportsDirectory { get; private set; }

        public ReportBuilderFunctionalTestsFixture()
        {
            if (Directory.Exists(ReportsDirectoryName))
            {
                Directory.Delete(ReportsDirectoryName, true);
            }

            ReportsDirectory = Directory.CreateDirectory(ReportsDirectoryName);
        }
    }
}
