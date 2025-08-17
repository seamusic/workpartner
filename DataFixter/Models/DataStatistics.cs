using System;
using System.Collections.Generic;
using System.Linq;

namespace DataFixter.Models
{
    /// <summary>
    /// 数据统计模型类
    /// 支持按点名、按文件、按调整类型的统计
    /// </summary>
    public class DataStatistics
    {
        /// <summary>
        /// 统计ID（唯一标识）
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();

        /// <summary>
        /// 统计时间
        /// </summary>
        public DateTime StatisticsTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 统计维度
        /// </summary>
        public StatisticsDimension Dimension { get; set; }

        /// <summary>
        /// 统计项名称（点名、文件名或调整类型）
        /// </summary>
        public string? ItemName { get; set; }

        /// <summary>
        /// 总数据量
        /// </summary>
        public int TotalCount { get; set; }

        /// <summary>
        /// 验证通过的数据量
        /// </summary>
        public int ValidCount { get; set; }

        /// <summary>
        /// 验证失败的数据量
        /// </summary>
        public int InvalidCount { get; set; }

        /// <summary>
        /// 需要调整的数据量
        /// </summary>
        public int NeedsAdjustmentCount { get; set; }

        /// <summary>
        /// 已调整的数据量
        /// </summary>
        public int AdjustedCount { get; set; }

        /// <summary>
        /// 未验证的数据量
        /// </summary>
        public int NotValidatedCount { get; set; }

        /// <summary>
        /// 按调整类型统计的详细数据
        /// </summary>
        public Dictionary<AdjustmentType, int> AdjustmentTypeCounts { get; set; } = new Dictionary<AdjustmentType, int>();

        /// <summary>
        /// 按数据方向统计的详细数据
        /// </summary>
        public Dictionary<DataDirection, int> DataDirectionCounts { get; set; } = new Dictionary<DataDirection, int>();

        /// <summary>
        /// 按验证严重程度统计的详细数据
        /// </summary>
        public Dictionary<ValidationSeverity, int> SeverityCounts { get; set; } = new Dictionary<ValidationSeverity, int>();

        /// <summary>
        /// 统计描述
        /// </summary>
        public string? Description { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        public DataStatistics()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="dimension">统计维度</param>
        /// <param name="itemName">统计项名称</param>
        public DataStatistics(StatisticsDimension dimension, string itemName)
        {
            Dimension = dimension;
            ItemName = itemName;
        }

        /// <summary>
        /// 计算验证通过率
        /// </summary>
        public double ValidationSuccessRate => TotalCount > 0 ? (double)ValidCount / TotalCount * 100 : 0;

        /// <summary>
        /// 计算验证失败率
        /// </summary>
        public double ValidationFailureRate => TotalCount > 0 ? (double)InvalidCount / TotalCount * 100 : 0;

        /// <summary>
        /// 计算需要调整率
        /// </summary>
        public double NeedsAdjustmentRate => TotalCount > 0 ? (double)NeedsAdjustmentCount / TotalCount * 100 : 0;

        /// <summary>
        /// 计算已调整率
        /// </summary>
        public double AdjustedRate => TotalCount > 0 ? (double)AdjustedCount / TotalCount * 100 : 0;

        /// <summary>
        /// 计算未验证率
        /// </summary>
        public double NotValidatedRate => TotalCount > 0 ? (double)NotValidatedCount / TotalCount * 100 : 0;

        /// <summary>
        /// 检查是否有验证失败的数据
        /// </summary>
        public bool HasInvalidData => InvalidCount > 0;

        /// <summary>
        /// 检查是否有需要调整的数据
        /// </summary>
        public bool HasDataNeedingAdjustment => NeedsAdjustmentCount > 0;

        /// <summary>
        /// 检查是否有已调整的数据
        /// </summary>
        public bool HasAdjustedData => AdjustedCount > 0;

        /// <summary>
        /// 检查是否所有数据都验证通过
        /// </summary>
        public bool AllDataValid => InvalidCount == 0 && NeedsAdjustmentCount == 0;

        /// <summary>
        /// 添加验证通过的数据
        /// </summary>
        /// <param name="count">数量</param>
        public void AddValidData(int count = 1)
        {
            ValidCount += count;
            TotalCount += count;
        }

        /// <summary>
        /// 添加验证失败的数据
        /// </summary>
        /// <param name="count">数量</param>
        public void AddInvalidData(int count = 1)
        {
            InvalidCount += count;
            TotalCount += count;
        }

        /// <summary>
        /// 添加需要调整的数据
        /// </summary>
        /// <param name="count">数量</param>
        public void AddNeedsAdjustmentData(int count = 1)
        {
            NeedsAdjustmentCount += count;
            TotalCount += count;
        }

        /// <summary>
        /// 添加已调整的数据
        /// </summary>
        /// <param name="count">数量</param>
        public void AddAdjustedData(int count = 1)
        {
            AdjustedCount += count;
            TotalCount += count;
        }

        /// <summary>
        /// 添加未验证的数据
        /// </summary>
        /// <param name="count">数量</param>
        public void AddNotValidatedData(int count = 1)
        {
            NotValidatedCount += count;
            TotalCount += count;
        }

        /// <summary>
        /// 添加调整类型统计
        /// </summary>
        /// <param name="adjustmentType">调整类型</param>
        /// <param name="count">数量</param>
        public void AddAdjustmentTypeCount(AdjustmentType adjustmentType, int count = 1)
        {
            if (AdjustmentTypeCounts.ContainsKey(adjustmentType))
            {
                AdjustmentTypeCounts[adjustmentType] += count;
            }
            else
            {
                AdjustmentTypeCounts[adjustmentType] = count;
            }
        }

        /// <summary>
        /// 添加数据方向统计
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <param name="count">数量</param>
        public void AddDataDirectionCount(DataDirection direction, int count = 1)
        {
            if (DataDirectionCounts.ContainsKey(direction))
            {
                DataDirectionCounts[direction] += count;
            }
            else
            {
                DataDirectionCounts[direction] = count;
            }
        }

        /// <summary>
        /// 添加验证严重程度统计
        /// </summary>
        /// <param name="severity">严重程度</param>
        /// <param name="count">数量</param>
        public void AddSeverityCount(ValidationSeverity severity, int count = 1)
        {
            if (SeverityCounts.ContainsKey(severity))
            {
                SeverityCounts[severity] += count;
            }
            else
            {
                SeverityCounts[severity] = count;
            }
        }

        /// <summary>
        /// 设置统计描述
        /// </summary>
        /// <param name="description">描述</param>
        public void SetDescription(string description)
        {
            Description = description;
        }

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>统计摘要字符串</returns>
        public string GetSummary()
        {
            var summary = $"{Dimension}: {ItemName} - 总计:{TotalCount}";
            
            if (ValidCount > 0)
                summary += $", 通过:{ValidCount}({ValidationSuccessRate:F1}%)";
            
            if (InvalidCount > 0)
                summary += $", 失败:{InvalidCount}({ValidationFailureRate:F1}%)";
            
            if (NeedsAdjustmentCount > 0)
                summary += $", 需调整:{NeedsAdjustmentCount}({NeedsAdjustmentRate:F1}%)";
            
            if (AdjustedCount > 0)
                summary += $", 已调整:{AdjustedCount}({AdjustedRate:F1}%)";
            
            if (NotValidatedCount > 0)
                summary += $", 未验证:{NotValidatedCount}({NotValidatedRate:F1}%)";
            
            return summary;
        }

        /// <summary>
        /// 获取详细统计信息
        /// </summary>
        /// <returns>详细统计信息字符串</returns>
        public string GetDetailedInfo()
        {
            var info = $"统计ID: {Id}\n";
            info += $"统计时间: {StatisticsTime:yyyy-MM-dd HH:mm:ss}\n";
            info += $"统计维度: {Dimension}\n";
            info += $"统计项: {ItemName}\n";
            info += $"总数据量: {TotalCount}\n";
            info += $"验证通过: {ValidCount} ({ValidationSuccessRate:F2}%)\n";
            info += $"验证失败: {InvalidCount} ({ValidationFailureRate:F2}%)\n";
            info += $"需要调整: {NeedsAdjustmentCount} ({NeedsAdjustmentRate:F2}%)\n";
            info += $"已调整: {AdjustedCount} ({AdjustedRate:F2}%)\n";
            info += $"未验证: {NotValidatedCount} ({NotValidatedRate:F2}%)\n";
            
            if (!string.IsNullOrEmpty(Description))
                info += $"统计描述: {Description}\n";
            
            if (AdjustmentTypeCounts.Count > 0)
            {
                info += $"按调整类型统计:\n";
                foreach (var kvp in AdjustmentTypeCounts.OrderByDescending(x => x.Value))
                {
                    info += $"  {kvp.Key}: {kvp.Value}\n";
                }
            }
            
            if (DataDirectionCounts.Count > 0)
            {
                info += $"按数据方向统计:\n";
                foreach (var kvp in DataDirectionCounts.OrderByDescending(x => x.Value))
                {
                    info += $"  {kvp.Key}: {kvp.Value}\n";
                }
            }
            
            if (SeverityCounts.Count > 0)
            {
                info += $"按严重程度统计:\n";
                foreach (var kvp in SeverityCounts.OrderByDescending(x => x.Value))
                {
                    info += $"  {kvp.Key}: {kvp.Value}\n";
                }
            }
            
            return info;
        }

        /// <summary>
        /// 获取CSV格式的统计信息
        /// </summary>
        /// <returns>CSV格式的统计信息</returns>
        public string GetCsvFormat()
        {
            var csv = $"统计维度,统计项,总数据量,验证通过,验证失败,需要调整,已调整,未验证,通过率(%),失败率(%),需调整率(%),已调整率(%),未验证率(%)\n";
            csv += $"{Dimension},{ItemName},{TotalCount},{ValidCount},{InvalidCount},{NeedsAdjustmentCount},{AdjustedCount},{NotValidatedCount},";
            csv += $"{ValidationSuccessRate:F2},{ValidationFailureRate:F2},{NeedsAdjustmentRate:F2},{AdjustedRate:F2},{NotValidatedRate:F2}";
            return csv;
        }

        /// <summary>
        /// 重写ToString方法
        /// </summary>
        public override string ToString()
        {
            return GetSummary();
        }

        /// <summary>
        /// 创建深拷贝
        /// </summary>
        /// <returns>深拷贝对象</returns>
        public DataStatistics Clone()
        {
            return new DataStatistics
            {
                Id = Id,
                StatisticsTime = StatisticsTime,
                Dimension = Dimension,
                ItemName = ItemName,
                TotalCount = TotalCount,
                ValidCount = ValidCount,
                InvalidCount = InvalidCount,
                NeedsAdjustmentCount = NeedsAdjustmentCount,
                AdjustedCount = AdjustedCount,
                NotValidatedCount = NotValidatedCount,
                AdjustmentTypeCounts = new Dictionary<AdjustmentType, int>(AdjustmentTypeCounts),
                DataDirectionCounts = new Dictionary<DataDirection, int>(DataDirectionCounts),
                SeverityCounts = new Dictionary<ValidationSeverity, int>(SeverityCounts),
                Description = Description
            };
        }

        /// <summary>
        /// 合并另一个统计数据
        /// </summary>
        /// <param name="other">另一个统计数据</param>
        public void Merge(DataStatistics other)
        {
            if (other == null) return;
            
            TotalCount += other.TotalCount;
            ValidCount += other.ValidCount;
            InvalidCount += other.InvalidCount;
            NeedsAdjustmentCount += other.NeedsAdjustmentCount;
            AdjustedCount += other.AdjustedCount;
            NotValidatedCount += other.NotValidatedCount;
            
            foreach (var kvp in other.AdjustmentTypeCounts)
            {
                AddAdjustmentTypeCount(kvp.Key, kvp.Value);
            }
            
            foreach (var kvp in other.DataDirectionCounts)
            {
                AddDataDirectionCount(kvp.Key, kvp.Value);
            }
            
            foreach (var kvp in other.SeverityCounts)
            {
                AddSeverityCount(kvp.Key, kvp.Value);
            }
        }
    }
}
