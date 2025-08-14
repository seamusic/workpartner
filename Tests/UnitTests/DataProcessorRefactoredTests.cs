using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner.Tests.UnitTests
{
    public class DataProcessorRefactoredTests
    {
        private readonly DataProcessorConfig _defaultConfig;

        public DataProcessorRefactoredTests()
        {
            _defaultConfig = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.05,
                RandomSeed = 42,
                TimeFactorWeight = 1.0,
                MinimumAdjustment = 0.001
            };
        }

        [Fact]
        public void ProcessMissingData_WithCumulativeColumns_ShouldCalculateCorrectly()
        {
            var files = CreateTestFilesWithCumulativeData();
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            
            Assert.NotNull(result);
            Assert.Equal(3, result.Count);
            
            var secondFile = result[1];
            var g10Row = secondFile.DataRows.FirstOrDefault(r => r.Name == "G10");
            Assert.NotNull(g10Row);
            Assert.Equal(15.0, g10Row.Values[0]);
        }

        [Fact]
        public void ProcessMissingData_WithConsecutiveMissingData_ShouldProcessCorrectly()
        {
            var files = CreateTestFilesWithConsecutiveMissingData();
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            
            Assert.NotNull(result);
            var secondFile = result[1];
            var testRow = secondFile.DataRows.FirstOrDefault(r => r.Name == "TestData");
            Assert.NotNull(testRow);
            Assert.True(testRow.Values[0].HasValue);
        }

        [Fact]
        public void ProcessMissingData_WithSupplementFiles_ShouldAdjustCorrectly()
        {
            var files = CreateTestFilesWithSupplementFiles();
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            
            Assert.NotNull(result);
            var supplementFiles = result.Where(f => f.IsSupplementFile).ToList();
            Assert.True(supplementFiles.Any());
        }

        private List<ExcelFile> CreateTestFilesWithCumulativeData()
        {
            var files = new List<ExcelFile>();
            
            var file1 = new ExcelFile
            {
                FileName = "test1.xls",
                Date = DateTime.Today,
                Hour = 0,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "G10", Values = new List<double?> { 10.0, 20.0, 30.0 } },
                    new DataRow { Name = "D10", Values = new List<double?> { 5.0, 10.0, 15.0 } }
                }
            };
            
            var file2 = new ExcelFile
            {
                FileName = "test2.xls",
                Date = DateTime.Today,
                Hour = 8,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "G10", Values = new List<double?> { null, null, null } },
                    new DataRow { Name = "D10", Values = new List<double?> { 5.0, 10.0, 15.0 } }
                }
            };
            
            var file3 = new ExcelFile
            {
                FileName = "test3.xls",
                Date = DateTime.Today,
                Hour = 16,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "G10", Values = new List<double?> { null, null, null } },
                    new DataRow { Name = "D10", Values = new List<double?> { 8.0, 12.0, 18.0 } }
                }
            };
            
            files.AddRange(new[] { file1, file2, file3 });
            return files;
        }

        private List<ExcelFile> CreateTestFilesWithConsecutiveMissingData()
        {
            var files = new List<ExcelFile>();
            
            var file1 = new ExcelFile
            {
                FileName = "test1.xls",
                Date = DateTime.Today,
                Hour = 0,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "TestData", Values = new List<double?> { 100.0, 200.0, 300.0 } }
                }
            };
            
            var file2 = new ExcelFile
            {
                FileName = "test2.xls",
                Date = DateTime.Today,
                Hour = 8,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "TestData", Values = new List<double?> { null, null, null } }
                }
            };
            
            files.AddRange(new[] { file1, file2 });
            return files;
        }

        private List<ExcelFile> CreateTestFilesWithSupplementFiles()
        {
            var files = new List<ExcelFile>();
            
            var originalFile = new ExcelFile
            {
                FileName = "original.xls",
                Date = DateTime.Today,
                Hour = 0,
                IsSupplementFile = false,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "TestData", Values = new List<double?> { 100.0, 200.0, 300.0 } }
                }
            };
            
            var supplementFile = new ExcelFile
            {
                FileName = "supplement.xls",
                Date = DateTime.Today,
                Hour = 8,
                IsSupplementFile = true,
                SupplementSource = new SupplementFileInfo
                {
                    SourceFile = originalFile,
                    AdjustmentParams = new AdjustmentParameters
                    {
                        AdjustmentRange = 0.05,
                        RandomSeed = 42,
                        MinimumAdjustment = 0.001,
                        MaintainDataCorrelation = true,
                        CorrelationWeight = 0.7
                    }
                },
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "TestData", Values = new List<double?> { 100.0, 200.0, 300.0 } }
                }
            };
            
            files.AddRange(new[] { originalFile, supplementFile });
            return files;
        }
    }
}
