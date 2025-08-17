using System;

namespace DataFixter.Models
{
    /// <summary>
    /// 调整记录模型类
    /// 记录每次数据调整的详细信息，支持调整前后的数据对比
    /// </summary>
    public class AdjustmentRecord
    {
        /// <summary>
        /// 调整ID（唯一标识）
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();

        /// <summary>
        /// 调整时间
        /// </summary>
        public DateTime AdjustmentTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 调整类型
        /// </summary>
        public AdjustmentType AdjustmentType { get; set; }

        /// <summary>
        /// 数据方向
        /// </summary>
        public DataDirection DataDirection { get; set; }

        /// <summary>
        /// 调整前的值
        /// </summary>
        public double OriginalValue { get; set; }

        /// <summary>
        /// 调整后的值
        /// </summary>
        public double AdjustedValue { get; set; }

        /// <summary>
        /// 调整幅度
        /// </summary>
        public double AdjustmentAmount => AdjustedValue - OriginalValue;

        /// <summary>
        /// 调整原因
        /// </summary>
        public string? Reason { get; set; }

        /// <summary>
        /// 调整描述
        /// </summary>
        public string? Description { get; set; }

        /// <summary>
        /// 相关文件名
        /// </summary>
        public string? FileName { get; set; }

        /// <summary>
        /// 相关点名
        /// </summary>
        public string? PointName { get; set; }

        /// <summary>
        /// 相关行号
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 是否成功调整
        /// </summary>
        public bool IsSuccessful { get; set; } = true;

        /// <summary>
        /// 调整失败的原因（如果调整失败）
        /// </summary>
        public string? FailureReason { get; set; }

        /// <summary>
        /// 调整的约束条件
        /// </summary>
        public string? Constraints { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        public AdjustmentRecord()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="adjustmentType">调整类型</param>
        /// <param name="dataDirection">数据方向</param>
        /// <param name="originalValue">调整前的值</param>
        /// <param name="adjustedValue">调整后的值</param>
        /// <param name="reason">调整原因</param>
        public AdjustmentRecord(AdjustmentType adjustmentType, DataDirection dataDirection, 
            double originalValue, double adjustedValue, string reason)
        {
            AdjustmentType = adjustmentType;
            DataDirection = dataDirection;
            OriginalValue = originalValue;
            AdjustedValue = adjustedValue;
            Reason = reason;
        }

        /// <summary>
        /// 设置调整失败
        /// </summary>
        /// <param name="failureReason">失败原因</param>
        public void SetFailure(string failureReason)
        {
            IsSuccessful = false;
            FailureReason = failureReason;
        }

        /// <summary>
        /// 设置调整成功
        /// </summary>
        public void SetSuccess()
        {
            IsSuccessful = true;
            FailureReason = null;
        }

        /// <summary>
        /// 设置相关文件信息
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="pointName">点名</param>
        /// <param name="rowNumber">行号</param>
        public void SetFileInfo(string fileName, string pointName, int rowNumber)
        {
            FileName = fileName;
            PointName = pointName;
            RowNumber = rowNumber;
        }

        /// <summary>
        /// 设置调整描述
        /// </summary>
        /// <param name="description">描述</param>
        public void SetDescription(string description)
        {
            Description = description;
        }

        /// <summary>
        /// 设置约束条件
        /// </summary>
        /// <param name="constraints">约束条件</param>
        public void SetConstraints(string constraints)
        {
            Constraints = constraints;
        }

        /// <summary>
        /// 获取调整幅度的绝对值
        /// </summary>
        public double AdjustmentAmountAbs => Math.Abs(AdjustmentAmount);

        /// <summary>
        /// 获取调整幅度的百分比
        /// </summary>
        public double AdjustmentPercentage => OriginalValue != 0 ? (AdjustmentAmount / Math.Abs(OriginalValue)) * 100 : 0;

        /// <summary>
        /// 检查调整是否在允许范围内
        /// </summary>
        /// <param name="maxAdjustment">最大允许调整幅度</param>
        /// <returns>是否在允许范围内</returns>
        public bool IsWithinAllowedRange(double maxAdjustment)
        {
            return AdjustmentAmountAbs <= maxAdjustment;
        }

        /// <summary>
        /// 获取调整摘要
        /// </summary>
        /// <returns>调整摘要字符串</returns>
        public string GetSummary()
        {
            if (IsSuccessful)
            {
                return $"{AdjustmentType} - {DataDirection}方向: {OriginalValue:F3} → {AdjustedValue:F3} (调整{AdjustmentAmount:F3})";
            }
            else
            {
                return $"{AdjustmentType} - {DataDirection}方向: 调整失败 - {FailureReason}";
            }
        }

        /// <summary>
        /// 获取详细调整信息
        /// </summary>
        /// <returns>详细调整信息字符串</returns>
        public string GetDetailedInfo()
        {
            var info = $"调整ID: {Id}\n";
            info += $"调整时间: {AdjustmentTime:yyyy-MM-dd HH:mm:ss}\n";
            info += $"调整类型: {AdjustmentType}\n";
            info += $"数据方向: {DataDirection}\n";
            info += $"调整前值: {OriginalValue:F3}\n";
            info += $"调整后值: {AdjustedValue:F3}\n";
            info += $"调整幅度: {AdjustmentAmount:F3}\n";
            info += $"调整百分比: {AdjustmentPercentage:F2}%\n";
            info += $"调整原因: {Reason}\n";
            
            if (!string.IsNullOrEmpty(Description))
                info += $"调整描述: {Description}\n";
            
            if (!string.IsNullOrEmpty(Constraints))
                info += $"约束条件: {Constraints}\n";
            
            info += $"文件名: {FileName}\n";
            info += $"点名: {PointName}\n";
            info += $"行号: {RowNumber}\n";
            info += $"调整状态: {(IsSuccessful ? "成功" : "失败")}\n";
            
            if (!IsSuccessful && !string.IsNullOrEmpty(FailureReason))
                info += $"失败原因: {FailureReason}\n";
            
            return info;
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
        public AdjustmentRecord Clone()
        {
            return new AdjustmentRecord
            {
                Id = Id,
                AdjustmentTime = AdjustmentTime,
                AdjustmentType = AdjustmentType,
                DataDirection = DataDirection,
                OriginalValue = OriginalValue,
                AdjustedValue = AdjustedValue,
                Reason = Reason,
                Description = Description,
                FileName = FileName,
                PointName = PointName,
                RowNumber = RowNumber,
                IsSuccessful = IsSuccessful,
                FailureReason = FailureReason,
                Constraints = Constraints
            };
        }
    }
}
