using System;
using System.Collections.Generic;

namespace DataFixter.Models
{
    /// <summary>
    /// 单期数据模型类
    /// 包含文件信息、时间、所有列数据，支持X、Y、Z三个方向的变化量
    /// </summary>
    public class PeriodData
    {
        /// <summary>
        /// 文件信息
        /// </summary>
        public ExcelFileInfo? FileInfo { get; set; }

        /// <summary>
        /// 数据行号（在Excel中的行号，从1开始）
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 序号
        /// </summary>
        public int SequenceNumber { get; set; }

        /// <summary>
        /// 点名（唯一标识）
        /// </summary>
        public string? PointName { get; set; }

        /// <summary>
        /// 里程
        /// </summary>
        public double Mileage { get; set; }

        /// <summary>
        /// 本期变化量（X方向）
        /// </summary>
        public double CurrentPeriodX { get; set; }

        /// <summary>
        /// 本期变化量（Y方向）
        /// </summary>
        public double CurrentPeriodY { get; set; }

        /// <summary>
        /// 本期变化量（Z方向）
        /// </summary>
        public double CurrentPeriodZ { get; set; }

        /// <summary>
        /// 累计变化量（X方向）
        /// </summary>
        public double CumulativeX { get; set; }

        /// <summary>
        /// 累计变化量（Y方向）
        /// </summary>
        public double CumulativeY { get; set; }

        /// <summary>
        /// 累计变化量（Z方向）
        /// </summary>
        public double CumulativeZ { get; set; }

        /// <summary>
        /// 日变化量（X方向）
        /// </summary>
        public double DailyX { get; set; }

        /// <summary>
        /// 日变化量（Y方向）
        /// </summary>
        public double DailyY { get; set; }

        /// <summary>
        /// 日变化量（Z方向）
        /// </summary>
        public double DailyZ { get; set; }

        /// <summary>
        /// 验证状态
        /// </summary>
        public ValidationStatus ValidationStatus { get; set; } = ValidationStatus.NotValidated;

        /// <summary>
        /// 验证失败的原因
        /// </summary>
        public List<string> ValidationErrors { get; set; } = new List<string>();

        /// <summary>
        /// 调整类型
        /// </summary>
        public AdjustmentType AdjustmentType { get; set; } = AdjustmentType.None;

        /// <summary>
        /// 调整记录
        /// </summary>
        public List<AdjustmentRecord> Adjustments { get; set; } = new List<AdjustmentRecord>();

        /// <summary>
        /// 是否可以修正
        /// 如果原始数据中已经存在的，则尽可能不要修改原来的数据
        /// </summary>
        public bool CanAdjustment { get; set; } = false;

        /// <summary>
        /// 构造函数
        /// </summary>
        public PeriodData()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileInfo">文件信息</param>
        /// <param name="rowNumber">行号</param>
        public PeriodData(ExcelFileInfo fileInfo, int rowNumber)
        {
            FileInfo = fileInfo;
            RowNumber = rowNumber;
        }

        /// <summary>
        /// 获取指定方向的本期变化量
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <returns>本期变化量</returns>
        public double GetCurrentPeriodValue(DataDirection direction)
        {
            return direction switch
            {
                DataDirection.X => CurrentPeriodX,
                DataDirection.Y => CurrentPeriodY,
                DataDirection.Z => CurrentPeriodZ,
                _ => throw new ArgumentException($"不支持的数据方向: {direction}")
            };
        }

        /// <summary>
        /// 设置指定方向的本期变化量
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <param name="value">值</param>
        public void SetCurrentPeriodValue(DataDirection direction, double value)
        {
            switch (direction)
            {
                case DataDirection.X:
                    CurrentPeriodX = value;
                    break;
                case DataDirection.Y:
                    CurrentPeriodY = value;
                    break;
                case DataDirection.Z:
                    CurrentPeriodZ = value;
                    break;
                default:
                    throw new ArgumentException($"不支持的数据方向: {direction}");
            }
        }

        /// <summary>
        /// 获取指定方向的累计变化量
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <returns>累计变化量</returns>
        public double GetCumulativeValue(DataDirection direction)
        {
            return direction switch
            {
                DataDirection.X => CumulativeX,
                DataDirection.Y => CumulativeY,
                DataDirection.Z => CumulativeZ,
                _ => throw new ArgumentException($"不支持的数据方向: {direction}")
            };
        }

        /// <summary>
        /// 设置指定方向的累计变化量
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <param name="value">值</param>
        public void SetCumulativeValue(DataDirection direction, double value)
        {
            switch (direction)
            {
                case DataDirection.X:
                    CumulativeX = value;
                    break;
                case DataDirection.Y:
                    CumulativeY = value;
                    break;
                case DataDirection.Z:
                    CumulativeZ = value;
                    break;
                default:
                    throw new ArgumentException($"不支持的数据方向: {direction}");
            }
        }

        /// <summary>
        /// 获取指定方向的日变化量
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <returns>日变化量</returns>
        public double GetDailyValue(DataDirection direction)
        {
            return direction switch
            {
                DataDirection.X => DailyX,
                DataDirection.Y => DailyY,
                DataDirection.Z => DailyZ,
                _ => throw new ArgumentException($"不支持的数据方向: {direction}")
            };
        }

        /// <summary>
        /// 设置指定方向的日变化量
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <param name="value">值</param>
        public void SetDailyValue(DataDirection direction, double value)
        {
            switch (direction)
            {
                case DataDirection.X:
                    DailyX = value;
                    break;
                case DataDirection.Y:
                    DailyY = value;
                    break;
                case DataDirection.Z:
                    DailyZ = value;
                    break;
                default:
                    throw new ArgumentException($"不支持的数据方向: {direction}");
            }
        }

        /// <summary>
        /// 添加验证错误
        /// </summary>
        /// <param name="error">错误信息</param>
        public void AddValidationError(string error)
        {
            ValidationErrors.Add(error);
            ValidationStatus = ValidationStatus.Invalid;
        }

        /// <summary>
        /// 清除验证错误
        /// </summary>
        public void ClearValidationErrors()
        {
            ValidationErrors.Clear();
            ValidationStatus = ValidationStatus.Valid;
        }

        /// <summary>
        /// 添加调整记录
        /// </summary>
        /// <param name="adjustment">调整记录</param>
        public void AddAdjustment(AdjustmentRecord adjustment)
        {
            Adjustments.Add(adjustment);
            AdjustmentType = adjustment.AdjustmentType;
        }

        /// <summary>
        /// 检查是否已调整
        /// </summary>
        /// <returns>是否已调整</returns>
        public bool HasBeenAdjusted => Adjustments.Count > 0;

        /// <summary>
        /// 获取调整次数
        /// </summary>
        public int AdjustmentCount => Adjustments.Count;

        /// <summary>
        /// 获取格式化的时间字符串
        /// </summary>
        public string FormattedTime => FileInfo?.FormattedDateTime ?? "未知时间";

        /// <summary>
        /// 获取项目名称
        /// </summary>
        public string ProjectName => FileInfo?.ProjectName ?? "未知项目";

        /// <summary>
        /// 重写ToString方法
        /// </summary>
        public override string ToString()
        {
            return $"{PointName} - {FormattedTime} - 序号:{SequenceNumber}";
        }

        /// <summary>
        /// 创建深拷贝
        /// </summary>
        /// <returns>深拷贝对象</returns>
        public PeriodData Clone()
        {
            return new PeriodData
            {
                FileInfo = FileInfo,
                RowNumber = RowNumber,
                SequenceNumber = SequenceNumber,
                PointName = PointName,
                Mileage = Mileage,
                CurrentPeriodX = CurrentPeriodX,
                CurrentPeriodY = CurrentPeriodY,
                CurrentPeriodZ = CurrentPeriodZ,
                CumulativeX = CumulativeX,
                CumulativeY = CumulativeY,
                CumulativeZ = CumulativeZ,
                DailyX = DailyX,
                DailyY = DailyY,
                DailyZ = DailyZ,
                ValidationStatus = ValidationStatus,
                ValidationErrors = new List<string>(ValidationErrors),
                AdjustmentType = AdjustmentType,
                Adjustments = new List<AdjustmentRecord>(Adjustments)
            };
        }
    }
}
