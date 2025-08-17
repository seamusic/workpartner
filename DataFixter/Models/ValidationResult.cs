using System;
using System.Collections.Generic;
using System.Linq;

namespace DataFixter.Models
{
    /// <summary>
    /// 数据验证结果模型类
    /// 记录验证失败的具体原因和位置，支持按点名、按文件、按类型分类
    /// </summary>
    public class ValidationResult
    {
        /// <summary>
        /// 验证ID（唯一标识）
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();

        /// <summary>
        /// 验证时间
        /// </summary>
        public DateTime ValidationTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 验证状态
        /// </summary>
        public ValidationStatus Status { get; set; }

        /// <summary>
        /// 验证类型
        /// </summary>
        public string? ValidationType { get; set; }

        /// <summary>
        /// 验证描述
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
        /// 数据方向
        /// </summary>
        public DataDirection? DataDirection { get; set; }

        /// <summary>
        /// 验证失败的具体原因
        /// </summary>
        public List<string> ErrorDetails { get; set; } = new List<string>();

        /// <summary>
        /// 验证失败的数据值
        /// </summary>
        public Dictionary<string, object> FailedValues { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// 期望的数据值
        /// </summary>
        public Dictionary<string, object> ExpectedValues { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// 验证规则
        /// </summary>
        public string? ValidationRule { get; set; }

        /// <summary>
        /// 严重程度
        /// </summary>
        public ValidationSeverity Severity { get; set; } = ValidationSeverity.Warning;

        /// <summary>
        /// 是否已修复
        /// </summary>
        public bool IsFixed { get; set; } = false;

        /// <summary>
        /// 修复时间
        /// </summary>
        public DateTime? FixedTime { get; set; }

        /// <summary>
        /// 修复方式
        /// </summary>
        public string? FixMethod { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        public ValidationResult()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="status">验证状态</param>
        /// <param name="validationType">验证类型</param>
        /// <param name="description">验证描述</param>
        public ValidationResult(ValidationStatus status, string validationType, string description)
        {
            Status = status;
            ValidationType = validationType;
            Description = description;
        }

        /// <summary>
        /// 设置验证失败
        /// </summary>
        /// <param name="errorDetail">错误详情</param>
        /// <param name="severity">严重程度</param>
        public void SetFailure(string errorDetail, ValidationSeverity severity = ValidationSeverity.Warning)
        {
            Status = ValidationStatus.Invalid;
            Severity = severity;
            if (!string.IsNullOrEmpty(errorDetail))
            {
                ErrorDetails.Add(errorDetail);
            }
        }

        /// <summary>
        /// 设置验证通过
        /// </summary>
        public void SetSuccess()
        {
            Status = ValidationStatus.Valid;
            ErrorDetails.Clear();
            FailedValues.Clear();
            ExpectedValues.Clear();
        }

        /// <summary>
        /// 设置需要调整
        /// </summary>
        public void SetNeedsAdjustment()
        {
            Status = ValidationStatus.NeedsAdjustment;
        }

        /// <summary>
        /// 添加错误详情
        /// </summary>
        /// <param name="errorDetail">错误详情</param>
        public void AddErrorDetail(string errorDetail)
        {
            if (!string.IsNullOrEmpty(errorDetail))
            {
                ErrorDetails.Add(errorDetail);
            }
        }

        /// <summary>
        /// 添加失败的数据值
        /// </summary>
        /// <param name="key">键</param>
        /// <param name="value">值</param>
        public void AddFailedValue(string key, object value)
        {
            if (!string.IsNullOrEmpty(key))
            {
                FailedValues[key] = value;
            }
        }

        /// <summary>
        /// 添加期望的数据值
        /// </summary>
        /// <param name="key">键</param>
        /// <param name="value">值</param>
        public void AddExpectedValue(string key, object value)
        {
            if (!string.IsNullOrEmpty(key))
            {
                ExpectedValues[key] = value;
            }
        }

        /// <summary>
        /// 设置文件信息
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
        /// 设置数据方向
        /// </summary>
        /// <param name="direction">数据方向</param>
        public void SetDataDirection(DataDirection direction)
        {
            DataDirection = direction;
        }

        /// <summary>
        /// 设置验证规则
        /// </summary>
        /// <param name="rule">验证规则</param>
        public void SetValidationRule(string rule)
        {
            ValidationRule = rule;
        }

        /// <summary>
        /// 标记为已修复
        /// </summary>
        /// <param name="fixMethod">修复方式</param>
        public void MarkAsFixed(string? fixMethod = null)
        {
            IsFixed = true;
            FixedTime = DateTime.Now;
            FixMethod = fixMethod;
        }

        /// <summary>
        /// 检查是否有错误详情
        /// </summary>
        public bool HasErrorDetails => ErrorDetails.Count > 0;

        /// <summary>
        /// 获取错误详情数量
        /// </summary>
        public int ErrorDetailCount => ErrorDetails.Count;

        /// <summary>
        /// 检查是否有失败的数据值
        /// </summary>
        public bool HasFailedValues => FailedValues.Count > 0;

        /// <summary>
        /// 检查是否有期望的数据值
        /// </summary>
        public bool HasExpectedValues => ExpectedValues.Count > 0;

        /// <summary>
        /// 获取验证摘要
        /// </summary>
        /// <returns>验证摘要字符串</returns>
        public string GetSummary()
        {
            var summary = $"{ValidationType}: {Status}";
            
            if (Status == ValidationStatus.Invalid)
            {
                summary += $" - {ErrorDetailCount}个错误";
                if (Severity != ValidationSeverity.Warning)
                {
                    summary += $" ({Severity})";
                }
            }
            
            if (!string.IsNullOrEmpty(PointName))
            {
                summary += $" - 点名:{PointName}";
            }
            
            if (!string.IsNullOrEmpty(FileName))
            {
                summary += $" - 文件:{System.IO.Path.GetFileName(FileName)}";
            }
            
            if (RowNumber > 0)
            {
                summary += $" - 行号:{RowNumber}";
            }
            
            return summary;
        }

        /// <summary>
        /// 获取详细验证信息
        /// </summary>
        /// <returns>详细验证信息字符串</returns>
        public string GetDetailedInfo()
        {
            var info = $"验证ID: {Id}\n";
            info += $"验证时间: {ValidationTime:yyyy-MM-dd HH:mm:ss}\n";
            info += $"验证状态: {Status}\n";
            info += $"验证类型: {ValidationType}\n";
            info += $"验证描述: {Description}\n";
            info += $"严重程度: {Severity}\n";
            
            if (!string.IsNullOrEmpty(FileName))
                info += $"文件名: {FileName}\n";
            
            if (!string.IsNullOrEmpty(PointName))
                info += $"点名: {PointName}\n";
            
            if (RowNumber > 0)
                info += $"行号: {RowNumber}\n";
            
            if (DataDirection.HasValue)
                info += $"数据方向: {DataDirection.Value}\n";
            
            if (!string.IsNullOrEmpty(ValidationRule))
                info += $"验证规则: {ValidationRule}\n";
            
            if (HasErrorDetails)
            {
                info += $"错误详情 ({ErrorDetailCount}个):\n";
                foreach (var error in ErrorDetails)
                {
                    info += $"  - {error}\n";
                }
            }
            
            if (HasFailedValues)
            {
                info += $"失败的数据值:\n";
                foreach (var kvp in FailedValues)
                {
                    info += $"  {kvp.Key}: {kvp.Value}\n";
                }
            }
            
            if (HasExpectedValues)
            {
                info += $"期望的数据值:\n";
                foreach (var kvp in ExpectedValues)
                {
                    info += $"  {kvp.Key}: {kvp.Value}\n";
                }
            }
            
            info += $"是否已修复: {(IsFixed ? "是" : "否")}\n";
            if (IsFixed)
            {
                info += $"修复时间: {FixedTime:yyyy-MM-dd HH:mm:ss}\n";
                if (!string.IsNullOrEmpty(FixMethod))
                    info += $"修复方式: {FixMethod}\n";
            }
            
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
        public ValidationResult Clone()
        {
            return new ValidationResult
            {
                Id = Id,
                ValidationTime = ValidationTime,
                Status = Status,
                ValidationType = ValidationType,
                Description = Description,
                FileName = FileName,
                PointName = PointName,
                RowNumber = RowNumber,
                DataDirection = DataDirection,
                ErrorDetails = new List<string>(ErrorDetails),
                FailedValues = new Dictionary<string, object>(FailedValues),
                ExpectedValues = new Dictionary<string, object>(ExpectedValues),
                ValidationRule = ValidationRule,
                Severity = Severity,
                IsFixed = IsFixed,
                FixedTime = FixedTime,
                FixMethod = FixMethod
            };
        }
    }

    /// <summary>
    /// 验证严重程度枚举
    /// </summary>
    public enum ValidationSeverity
    {
        /// <summary>
        /// 信息
        /// </summary>
        Info,
        
        /// <summary>
        /// 警告
        /// </summary>
        Warning,
        
        /// <summary>
        /// 错误
        /// </summary>
        Error,
        
        /// <summary>
        /// 严重错误
        /// </summary>
        Critical
    }
}
